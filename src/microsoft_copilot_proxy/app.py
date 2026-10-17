from __future__ import annotations

import json
import time
import uuid
from collections.abc import AsyncIterator, Callable

from fastapi import Depends, FastAPI, HTTPException, Request
from fastapi.responses import JSONResponse, StreamingResponse

from .config import AVAILABLE_MODELS, Settings
from .session_store import PersistentSession, PersistentSessionStore
from .substrate_client import SubstrateCopilotClient, SubstrateCopilotError
from .token_store import AccessTokenStore
from .models import AnthropicMessagesRequest, OpenAIChatRequest, OpenAIResponsesRequest
from .translator import translate_anthropic_request, translate_openai_request, translate_responses_request

_PERSIST_MODEL_SUFFIX = ":persist"
_SESSION_ID_HEADER = "x-m365-session-id"
_ANTHROPIC_MODEL_PREFIX = "claude-"
_ANTHROPIC_DEFAULT_CONTEXT_WINDOW = 200_000


def create_app(
    settings: Settings | None = None,
    copilot_client_factory: Callable[[], SubstrateCopilotClient] | None = None,
) -> FastAPI:
    app = FastAPI(title="Microsoft 365 Copilot OpenAI Proxy")
    resolved_settings = settings or Settings()
    app.state.settings = resolved_settings
    app.state.token_store = AccessTokenStore(resolved_settings.access_token)
    app.state.session_store = PersistentSessionStore()
    app.state.copilot_client_factory = copilot_client_factory or (
        lambda: SubstrateCopilotClient(app.state.token_store.get())
    )

    def get_settings() -> Settings:
        return app.state.settings

    def get_copilot_client() -> SubstrateCopilotClient:
        return app.state.copilot_client_factory()

    @app.get("/healthz")
    async def healthz() -> dict:
        return {"status": "ok", "token": app.state.token_store.status()}

    @app.get("/v1/token/status")
    async def token_status() -> dict:
        return app.state.token_store.status()

    @app.get("/v1/models")
    async def list_models() -> dict:
        return _anthropic_models_response()

    @app.post("/v1/chat/completions")
    async def chat_completions(
        raw_request: Request,
        request: OpenAIChatRequest,
        settings: Settings = Depends(get_settings),
        client: SubstrateCopilotClient = Depends(get_copilot_client),
    ):
        try:
            translated = translate_openai_request(request)
            session = _persistent_session(app, raw_request, request.model, request.user)
            if request.stream:
                return StreamingResponse(
                    _openai_stream(
                        settings.model_alias,
                        client,
                        translated.prompt,
                        translated.additional_context,
                        request.model,
                        session,
                    ),
                    media_type="text/event-stream",
                )
            text = await client.chat(
                translated.prompt,
                translated.additional_context,
                request.model,
                session,
            )
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        except SubstrateCopilotError as exc:
            raise HTTPException(status_code=502, detail=str(exc)) from exc

        return JSONResponse({
            "id": f"chatcmpl_{uuid.uuid4().hex}",
            "object": "chat.completion",
            "created": int(time.time()),
            "model": settings.model_alias,
            "choices": [
                {
                    "index": 0,
                    "message": {"role": "assistant", "content": text},
                    "finish_reason": "stop",
                }
            ],
        })

    @app.post("/v1/responses")
    async def openai_responses(
        raw: Request,
        settings: Settings = Depends(get_settings),
        client: SubstrateCopilotClient = Depends(get_copilot_client),
    ):
        body = await raw.json()
        try:
            request = OpenAIResponsesRequest.model_validate(body)
            translated = translate_responses_request(request)
            session = _persistent_session(app, raw, request.model)
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc

        if request.stream:
            return StreamingResponse(
                _responses_stream(
                    settings.model_alias,
                    client,
                    translated.prompt,
                    translated.additional_context,
                    request.model,
                    session,
                ),
                media_type="text/event-stream",
            )

        try:
            text = await client.chat(
                translated.prompt,
                translated.additional_context,
                request.model,
                session,
            )
        except SubstrateCopilotError as exc:
            raise HTTPException(status_code=502, detail=str(exc)) from exc

        return JSONResponse({
            "id": f"resp_{uuid.uuid4().hex}",
            "object": "response",
            "created_at": int(time.time()),
            "model": settings.model_alias,
            "output": [{
                "type": "message",
                "id": f"msg_{uuid.uuid4().hex}",
                "role": "assistant",
                "content": [{"type": "output_text", "text": text}],
            }],
            "usage": {"input_tokens": 0, "output_tokens": 0, "total_tokens": 0},
        })

    @app.post("/v1/messages")
    async def anthropic_messages(
        raw_request: Request,
        request: AnthropicMessagesRequest,
        settings: Settings = Depends(get_settings),
        client: SubstrateCopilotClient = Depends(get_copilot_client),
    ):
        try:
            translated = translate_anthropic_request(request)
            routed_model = _resolve_anthropic_model(request.model)
            session = _persistent_session(app, raw_request, routed_model)
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc

        if request.stream:
            return StreamingResponse(
                _anthropic_stream(
                    settings.model_alias,
                    client,
                    translated.prompt,
                    translated.additional_context,
                    routed_model,
                    session,
                ),
                media_type="text/event-stream",
            )

        try:
            text = await client.chat(
                translated.prompt,
                translated.additional_context,
                routed_model,
                session,
            )
        except SubstrateCopilotError as exc:
            raise HTTPException(status_code=502, detail=str(exc)) from exc

        return JSONResponse({
            "id": f"msg_{uuid.uuid4().hex}",
            "type": "message",
            "role": "assistant",
            "model": settings.model_alias,
            "content": [{"type": "text", "text": text}],
            "stop_reason": "end_turn",
            "stop_sequence": None,
            "usage": {"input_tokens": 0, "output_tokens": 0},
        })

    return app


def _openai_models_response() -> dict:
    return {
        "object": "list",
        "data": [
            {
                "id": model,
                "object": "model",
                "owned_by": "Microsoft_Copilot",
            }
            for model in AVAILABLE_MODELS
        ],
    }


def _anthropic_models_response() -> dict:
    data = [_anthropic_model(model) for model in AVAILABLE_MODELS]
    return {
        "data": data,
        "first_id": data[0]["id"] if data else None,
        "has_more": False,
        "last_id": data[-1]["id"] if data else None,
    }


def _anthropic_model(model: str) -> dict:
    public_id = model
    return {
        "id": public_id,
        "type": "model",
        "display_name": model,
        "created_at": "1970-01-01T00:00:00Z",
        "max_input_tokens": _ANTHROPIC_DEFAULT_CONTEXT_WINDOW,
        "max_tokens": 0,
        "capabilities": _anthropic_capabilities(),
    }


def _anthropic_capabilities() -> dict:
    return {
        "audio_input": {"supported": False},
        "batch": {"supported": False},
        "citations": {"supported": True},
        "code_execution": {"supported": False},
        "context_management": {
            "clear_thinking_20251015": {"supported": True},
            "clear_tool_uses_20250919": {"supported": True},
            "compact_20260112": {"supported": False},
            "max_input_tokens": _ANTHROPIC_DEFAULT_CONTEXT_WINDOW,
            "supported": True,
        },
        "context_window": {
            "max_input_tokens": _ANTHROPIC_DEFAULT_CONTEXT_WINDOW,
            "one_million_context_variant": False,
            "supported": True,
            "supports_1m_context": False,
        },
        "effort": {
            "high": {"supported": True},
            "low": {"supported": True},
            "max": {"supported": True},
            "medium": {"supported": True},
            "supported": True,
            "ultra": {"supported": True},
            "xhigh": {"supported": True},
        },
        "image_input": {"supported": False},
        "pdf_input": {"supported": False},
        "structured_outputs": {"supported": False},
        "thinking": {
            "supported": True,
            "types": {
                "adaptive": {"supported": True},
                "enabled": {"supported": True},
            },
        },
        "tool_use": {"supported": True},
        "video_input": {"supported": False},
    }


def _resolve_anthropic_model(model: str) -> str:
    normalized = model.strip()
    if not normalized.lower().startswith(_ANTHROPIC_MODEL_PREFIX):
        return model

    candidate = normalized[len(_ANTHROPIC_MODEL_PREFIX):]
    model_lookup = {available.lower(): available for available in AVAILABLE_MODELS}
    return model_lookup.get(candidate.lower(), model)


def _persistent_session(
    app: FastAPI,
    raw_request: Request,
    model: str,
    fallback_key: str | None = None,
) -> PersistentSession | None:
    header_key = (
        raw_request.headers.get(_SESSION_ID_HEADER)
        or fallback_key
        or "default"
    )

    return app.state.session_store.get(header_key)


async def _openai_stream(
    model_alias: str,
    client: SubstrateCopilotClient,
    prompt: str,
    additional_context: list[str],
    model: str,
    session: PersistentSession | None = None,
) -> AsyncIterator[str]:
    completion_id = f"chatcmpl_{uuid.uuid4().hex}"
    created = int(time.time())
    first_chunk = {
        "id": completion_id,
        "object": "chat.completion.chunk",
        "created": created,
        "model": model_alias,
        "choices": [{"index": 0, "delta": {"role": "assistant"}, "finish_reason": None}],
    }
    yield f"data: {json.dumps(first_chunk)}\n\n"
    try:
        async for delta in client.chat_stream(prompt, additional_context, model, session):
            chunk = {
                "id": completion_id,
                "object": "chat.completion.chunk",
                "created": created,
                "model": model_alias,
                "choices": [{"index": 0, "delta": {"content": delta}, "finish_reason": None}],
            }
            yield f"data: {json.dumps(chunk)}\n\n"
    except SubstrateCopilotError as exc:
        yield f"data: {json.dumps({'error': {'message': str(exc), 'type': 'upstream_error'}})}\n\n"
        yield "data: [DONE]\n\n"
        return
    final_chunk = {
        "id": completion_id,
        "object": "chat.completion.chunk",
        "created": created,
        "model": model_alias,
        "choices": [{"index": 0, "delta": {}, "finish_reason": "stop"}],
    }
    yield f"data: {json.dumps(final_chunk)}\n\n"
    yield "data: [DONE]\n\n"


async def _responses_stream(
    model_alias: str,
    client: SubstrateCopilotClient,
    prompt: str,
    additional_context: list[str],
    model: str,
    session: PersistentSession | None = None,
) -> AsyncIterator[str]:
    resp_id = f"resp_{uuid.uuid4().hex}"
    item_id = f"msg_{uuid.uuid4().hex}"
    created = int(time.time())

    yield f"data: {json.dumps({'type': 'response.created', 'response': {'id': resp_id, 'object': 'response', 'created_at': created, 'model': model_alias, 'status': 'in_progress', 'output': []}})}\n\n"
    yield f"data: {json.dumps({'type': 'response.output_item.added', 'output_index': 0, 'item': {'id': item_id, 'type': 'message', 'role': 'assistant', 'content': []}})}\n\n"
    yield f"data: {json.dumps({'type': 'response.content_part.added', 'item_id': item_id, 'output_index': 0, 'content_index': 0, 'part': {'type': 'output_text', 'text': ''}})}\n\n"

    full_text = ""
    try:
        async for delta in client.chat_stream(prompt, additional_context, model, session):
            full_text += delta
            yield f"data: {json.dumps({'type': 'response.output_text.delta', 'item_id': item_id, 'output_index': 0, 'content_index': 0, 'delta': delta})}\n\n"
    except SubstrateCopilotError as exc:
        yield f"data: {json.dumps({'type': 'error', 'error': {'message': str(exc), 'type': 'upstream_error'}})}\n\n"
        return

    yield f"data: {json.dumps({'type': 'response.output_text.done', 'item_id': item_id, 'output_index': 0, 'content_index': 0, 'text': full_text})}\n\n"
    yield f"data: {json.dumps({'type': 'response.completed', 'response': {'id': resp_id, 'object': 'response', 'created_at': created, 'model': model_alias, 'status': 'completed', 'output': [{'id': item_id, 'type': 'message', 'role': 'assistant', 'content': [{'type': 'output_text', 'text': full_text}]}], 'usage': {'input_tokens': 0, 'output_tokens': 0, 'total_tokens': 0}}})}\n\n"


async def _anthropic_stream(
    model_alias: str,
    client: SubstrateCopilotClient,
    prompt: str,
    additional_context: list[str],
    model: str,
    session: PersistentSession | None = None,
) -> AsyncIterator[str]:
    msg_id = f"msg_{uuid.uuid4().hex}"

    def sse(event: str, data: dict) -> str:
        return f"event: {event}\ndata: {json.dumps(data)}\n\n"

    yield sse("message_start", {"type": "message_start", "message": {"id": msg_id, "type": "message", "role": "assistant", "content": [], "model": model_alias, "stop_reason": None, "stop_sequence": None, "usage": {"input_tokens": 0, "output_tokens": 0}}})
    yield sse("content_block_start", {"type": "content_block_start", "index": 0, "content_block": {"type": "text", "text": ""}})
    yield sse("ping", {"type": "ping"})

    try:
        async for delta in client.chat_stream(prompt, additional_context, model, session):
            yield sse("content_block_delta", {"type": "content_block_delta", "index": 0, "delta": {"type": "text_delta", "text": delta}})
    except SubstrateCopilotError as exc:
        yield sse("error", {"type": "error", "error": {"type": "upstream_error", "message": str(exc)}})
        return

    yield sse("content_block_stop", {"type": "content_block_stop", "index": 0})
    yield sse("message_delta", {"type": "message_delta", "delta": {"stop_reason": "end_turn", "stop_sequence": None}, "usage": {"output_tokens": 0}})
    yield sse("message_stop", {"type": "message_stop"})
