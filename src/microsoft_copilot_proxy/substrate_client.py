from __future__ import annotations

import json
import time
import uuid
from collections.abc import AsyncIterator
from urllib.parse import quote

import websockets

from .session_store import PersistentSession
from .token_store import (
    decode_jwt_payload,
    is_substrate_token_claims,
    read_env_value,
    token_serialization_type,
)

SIGNALR_SEP = "\x1e"
_WS_BASE = "wss://substrate.office.com/m365Copilot/Chathub"

_VARIANTS = (
    "EnableMcpServerWidgets,feature.EnableMcpServerWidgets,"
    "feature.EnableImageGenInsufficientTokensThrottled,"
    "feature.EnableImageGenSystemCapacityThrottled,"
    "feature.EnableLuForChatCIQ,feature.enableChatCIQPlugin,"
    "EnableRequestPlugins,feature.EnableSensitivityLabels,"
    "EnableUnsupportedUrlDetector,feature.IsCustomEngineCopilotEnabled,"
    "feature.bizchatfluxv3,feature.enablechatpages,feature.enableCodeCanvas,"
    "feature.turnOnDARecommendation,"
    "feature.IsStreamingModeInChatRequestEnabled,"
    "IncludeSourceAttributionsConcise,SkipPublishEmptyMessage,"
    "feature.EnableDeduplicatingSourceAttributions,"
    "feature.IsCitationsReferencesOutputEnabled,"
    "feature.enableDeltaStreamingForReferences,"
    "feature.enableIncludeReferencesInDeltaResponse,"
    "feature.enablereferencesforagents,"
    "feature.EnableCodeInterpreterConversion,"
    "agt_module_attr_enableReferencesForCodeInterpreter,"
    "agt_module_enableCodeInterpreterHallucinatedUrlFilter,"
    "agt_module_attr_enableCodeInterpreterFilePreviewReference,"
    "cdxcipreviewmsg,Enable3PActionProgressMessages,"
    "feature.enableClientWebRtc,"
    "feature.EnableMeetingRecapOfSeriesMeetingWithCiq,"
    "feature.cwcfluxv3fe,feature.cwcfluxv3fem,"
    "feature.EnableReferencesListCompleteSignal,"
    "feature.StorageMessageSplitDisabled,SingletonEnvOn,"
    "cdxenablefccinmainline,"
    "agt_bizchat_enableSearchResultProgressMessages,"
    "EnableComposeWidget,agt_researcheragent_enableMemoryRead,"
    "cdxweb_search_citations_answer_cards,"
    "cdxenable_strongly_typed_bing_grounding_conversion,"
    "cdxweb_search_citations_video_answer_card,"
    "cdxweb_search_citations_places_answer_card,"
    "cdxweb_search_citations_sports_answer_card,"
    "cdxweb_search_citations_weather_answer_card,"
    "cdxweb_search_citations_image_answer_card,"
    "agt_bizchat_enableImageListSdlCard,"
    "agt_webagent_enableImageListSdlCard,cdximage_sdl_card,"
    "cdxshopping_berry_answer_card_msa,"
    "unitab_msa_enableEducationLearningTool,"
    "skds_msa_EnableGlobalIndex,feature.cwcallowedos,"
    "feature.EnableMergingPureDeltas,feature.disabledisallowedmsgs,"
    "feature.enableCitationsForSynthesisData,"
    "feature.EnableConversationShareApis,"
    "feature.EnableConversationShareApisForMsa,"
    "feature.EnableGoodbyeDrainGate,"
    "feature.enableGenerateGraphicArtOptionsSet,cdximagen,"
    "feature.EnableUpdatedUXForConfirmationDialog,"
    "feature.EnableContentApiandDocTypeHtmlInRichAnswers,"
    "cdxgrounding_api_v2_rich_web_answers_reference_bottom_force,"
    "cdxenablerenderforisocomp,"
    "feature.EnableClientFileURLSupportForOfficeWebPaidCopilot,"
    "feature.EnableDesignEditorImageGrounding,"
    "feature.EnableDesignerEditor,"
    "feature.EnableSkipRehydrationForSpeCIdImages,"
    "feature.sourcescontrolmainline,feature.sourcescontrolmainlineal,"
    "feature.EnableConnectorExecutionControlsAllowlist,"
    "feature.EnableBizchatMainlineExecutionControlsResolution,"
    "rich_responses,feature.EnableBase64DataInMessageAnnotations,"
    "feature.EnableSkipEmittingMessageOnFlush,"
    "feature.EnableRemoveEmptySourceAttributions,"
    "feature.EnableRemoveStreamingMode"
)

_OPTIONS_SETS = [
    "search_result_progress_messages_with_search_queries",
    "cwc_flux_image",
    "cwc_code_interpreter",
    "cwc_code_interpreter_amsfix",
    "cwcfluxgptv",
    "flux_v3_gptv_enable_upload_multi_image_in_turn_wo_ch",
    "cwc_code_interpreter_citation_fix",
    "code_interpreter_interactive_charts",
    "cwc_code_interpreter_interactive_charts_inline_image",
    "code_interpreter_matplotlib_patching",
    "cwc_fileupload_odb",
    "update_memory_plugin",
    "add_custom_instructions",
    "cwc_flux_v3",
    "flux_v3_progress_messages",
    "enable_batch_token_processing",
    "enable_gg_gpt",
    "flux_v3_image_gen_enable_dimensions",
    "flux_v3_image_gen_enable_icon_dimensions",
    "flux_v3_image_gen_enable_system_text_with_params",
    "flux_v3_image_gen_enable_designer_dimensions_meta_prompting_in_system_prompts",
]

_ALLOWED_MESSAGE_TYPES = [
    "Chat",
    "Suggestion",
    "InternalSearchQuery",
    "Disengaged",
    "InternalLoaderMessage",
    "Progress",
    "GeneratedCode",
    "RenderCardRequest",
    "AdsQuery",
    "SemanticSerp",
    "GenerateContentQuery",
    "GenerateGraphicArt",
    "SearchQuery",
    "ConfirmationCard",
    "AuthError",
    "DeveloperLogs",
    "TriggerPlugin",
    "HintInvocation",
    "MemoryUpdate",
    "EndOfRequest",
    "TriggerConfirmation",
    "ResumeInvokeAction",
    "ResumeUserInputRequest",
    "TriggerUserInputRequest",
    "EscapeHatch",
    "TriggerPluginAuth",
    "ResumePluginAuth",
    "SideBySide",
    "ReferencesListComplete",
    "SwitchRespondingEndpoint",
]


class SubstrateCopilotError(RuntimeError):
    pass


class SubstrateCopilotClient:
    def __init__(
        self,
        access_token: str,
    ):
        if not access_token:
            raise SubstrateCopilotError(
                "M365_ACCESS_TOKEN is missing. Run capture-token --headed first."
            )

        self._token = access_token
        token_type = token_serialization_type(access_token)

        if token_type == "jwt":
            try:
                claims = decode_jwt_payload(access_token)
            except Exception as exc:
                raise SubstrateCopilotError(
                    f"Cannot decode access token: {exc}"
                ) from exc

            if not is_substrate_token_claims(claims):
                raise SubstrateCopilotError(
                    "Access token is not a substrate.office.com token."
                )

            if time.time() > claims.get("exp", 0):
                raise SubstrateCopilotError(
                    "Access token expired. Run capture-token again."
                )

            self._oid = str(claims["oid"])
            self._tid = str(claims["tid"])

        elif token_type == "jwe":
            self._oid = read_env_value(
                ".env",
                "M365_ROUTE_USER_ID",
            ) or ""
            self._tid = read_env_value(
                ".env",
                "M365_ROUTE_TENANT_ID",
            ) or ""

            if not self._oid or not self._tid:
                raise SubstrateCopilotError(
                    "Encrypted token requires M365_ROUTE_USER_ID and "
                    "M365_ROUTE_TENANT_ID in .env. Run capture-token again."
                )

        else:
            raise SubstrateCopilotError(
                "Access token is not a compact JWT or JWE."
            )

    def _ws_url(
        self,
        conv_id: str,
        session_id: str,
        req_id: str,
    ) -> str:
        token = quote(self._token, safe="")

        # The current consumer Copilot request uses the same 32-character
        # routing key for these three query parameters.
        routing_key = session_id.replace("-", "")

        return (
            f"{_WS_BASE}/{self._oid}@{self._tid}"
            f"?chatsessionid={routing_key}"
            f"&XRoutingParameterSessionKey={routing_key}"
            f"&clientrequestid={routing_key}"
            f"&X-SessionId={session_id}"
            f"&ConversationId={conv_id}"
            f"&access_token={token}"
            f"&variants={_VARIANTS}"
            f"&source=%22officeweb%22"
            f"&product=Office"
            f"&agentHost=Bizchat.FullScreen"
            f"&licenseType=Starter"
            f"&isEdu=false"
            f"&agent=web"
            f"&scenario=OfficeWebPaidConsumerCopilot"
        )

    def _chat_invoke(
        self,
        text: str,
        tone: str,
        conv_id: str,
        session_id: str,
        req_id: str,
        is_start_of_session: bool,
    ) -> str:
        print(f"Tone: {tone}")
        payload = {
            "arguments": [
                {
                    "source": "officeweb",
                    "clientCorrelationId": req_id,
                    "sessionId": session_id,
                    "optionsSets": _OPTIONS_SETS,
                    "streamingMode": "ConciseWithPadding",
                    "spokenTextMode": "None",
                    "options": {},
                    "extraExtensionParameters": {},
                    "allowedMessageTypes": _ALLOWED_MESSAGE_TYPES,
                    "sliceIds": [],
                    "threadLevelGptId": {},
                    "traceId": req_id,
                    "isStartOfSession": is_start_of_session,
                    "clientInfo": {
                        "clientPlatform": "mcmcopilot-web",
                        "clientAppName": "Office",
                        "clientEntrypoint": "mcmcopilot-officeweb",
                        "clientSessionId": session_id,
                        "clientAppType": "Web",
                        "deviceOS": "Windows",
                        "deviceType": "Desktop",
                    },
                    "message": {
                        "author": "user",
                        "inputMethod": "Keyboard",
                        "text": text,
                        "entityAnnotationTypes": [
                            "People",
                            "File",
                            "Event",
                            "Email",
                            "TeamsMessage",
                        ],
                        "requestId": req_id,
                        "locale": "en-us",
                        "messageType": "Chat",
                        "experienceType": "Default",
                        "adaptiveCards": [],
                        "clientPreferences": {},
                    },
                    "plugins": [
                        {
                            "Id": "BingWebSearch",
                            "Source": "BuiltIn",
                        }
                    ],
                    "isSbsSupported": True,
                    "tone": tone,
                    "renderReferencesBehindEOS": True,
                }
            ],
            "invocationId": "0",
            "target": "chat",
            "type": 4,
        }
        return json.dumps(payload, ensure_ascii=False) + SIGNALR_SEP

    async def chat_stream(
        self,
        prompt: str,
        additional_context: list[str],
        model: str,
        session: PersistentSession | None = None,
    ) -> AsyncIterator[str]:
        text = _combine_text(prompt, additional_context)

        if session is None:
            async for chunk in self._chat_stream_for_turn(
                text=text,
                model=model,
                conv_id=str(uuid.uuid4()),
                session_id=str(uuid.uuid4()),
                is_start_of_session=True,
            ):
                yield chunk
            return

        async with session.lock:
            turn = session.reserve_turn()
            async for chunk in self._chat_stream_for_turn(
                text=text,
                model=model,
                conv_id=turn.conversation_id,
                session_id=turn.client_session_id,
                is_start_of_session=turn.is_start_of_session,
            ):
                yield chunk

    async def _chat_stream_for_turn(
        self,
        text: str,
        model: str,
        conv_id: str,
        session_id: str,
        is_start_of_session: bool,
    ) -> AsyncIterator[str]:
        req_id = str(uuid.uuid4())
        url = self._ws_url(
            conv_id,
            session_id,
            req_id,
        )

        redacted_url = url.split("&access_token=", 1)[0]
        # print(
        #     "Opening Substrate WebSocket:",
        #     redacted_url + "&access_token=<redacted>",
        # )

        try:
            async with websockets.connect(
                url,
                additional_headers={
                    "Origin": "https://m365.cloud.microsoft",
                    "Pragma": "no-cache",
                    "Cache-Control": "no-cache",
                },
            ) as ws:
                await ws.send(
                    json.dumps(
                        {
                            "protocol": "json",
                            "version": 1,
                        }
                    )
                    + SIGNALR_SEP
                )

                handshake = await ws.recv()
                # print("SignalR handshake:", repr(handshake))

                await ws.send(
                    self._chat_invoke(
                        text,
                        model,
                        conv_id,
                        session_id,
                        req_id,
                        is_start_of_session,
                    )
                )

                fallback_text = ""
                yielded_any = False

                async for raw in ws:
                    for part in raw.split(SIGNALR_SEP):
                        part = part.strip()
                        if not part:
                            continue

                        try:
                            msg = json.loads(part)
                        except json.JSONDecodeError:
                            continue

                        message_type = msg.get("type")

                        if message_type == 6:
                            continue

                        if (
                            message_type == 1
                            and msg.get("target") == "update"
                        ):
                            args = (
                                msg.get("arguments")
                                or [{}]
                            )[0]

                            delta = args.get("writeAtCursor")
                            if delta:
                                if not yielded_any and fallback_text:
                                    yield fallback_text
                                yielded_any = True
                                yield delta

                            messages = args.get("messages")
                            if messages:
                                entries = (
                                    messages
                                    if isinstance(messages, list)
                                    else [messages]
                                )
                                for entry in reversed(entries):
                                    if entry.get("author") != "user":
                                        fallback_text = entry.get(
                                            "text",
                                            "",
                                        )
                                        break

                        if message_type == 2:
                            item_messages = (
                                msg.get("item") or {}
                            ).get("messages") or []

                            for entry in reversed(item_messages):
                                if entry.get("author") != "user":
                                    fallback_text = entry.get(
                                        "text",
                                        "",
                                    )
                                    break

                        if message_type == 3:
                            if not yielded_any and fallback_text:
                                yield fallback_text
                            return

        except websockets.exceptions.ConnectionClosed as exc:
            raise SubstrateCopilotError(
                "Substrate WebSocket closed: "
                f"code={exc.code}, reason={exc.reason!r}"
            ) from exc
        except SubstrateCopilotError:
            raise
        except Exception as exc:
            raise SubstrateCopilotError(str(exc)) from exc

    async def chat(
        self,
        prompt: str,
        additional_context: list[str],
        model: str,
        session: PersistentSession | None = None,
    ) -> str:
        chunks: list[str] = []

        async for chunk in self.chat_stream(
            prompt,
            additional_context,
            model,
            session,
        ):
            chunks.append(chunk)

        return "".join(chunks)


def _combine_text(
    prompt: str,
    context: list[str],
) -> str:
    if not context:
        return prompt

    return "\n\n".join(context) + "\n\n---\n\n" + prompt
