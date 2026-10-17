# Microsoft Copilot Proxy

Use Microsoft Copilot through different clients, local scripts, and coding tools.

This project runs a local FastAPI proxy that talks to the same `substrate.office.com` WebSocket API used by the Microsoft Copilot web UI, then exposes it as different HTTP endpoints.

No Azure app registration. No admin consent. Sign in with your normal Microsoft Copilot browser session.

## Why Use This

- Use Microsoft Copilot from different clients
- Works with your existing signed-in Copilot web session
- Runs locally on `127.0.0.1` by default
- Auto-captures and refreshes the short-lived browser token
- Supports persistent Copilot sessions across turns
- Supports OpenAI Chat Completions, OpenAI Responses, and Anthropic Messages style requests

## Quick Start

```powershell
uv sync
uv run playwright install chromium
uv run microsoft-copilot-proxy capture-token --headed
uv run microsoft-copilot-proxy serve
```

The server starts at:

```text
http://127.0.0.1:8000
```

On first run, the proxy opens a dedicated browser window. Sign in to Microsoft Copilot there once. The proxy will capture the required Substrate token and write it to `.env`.

The dedicated browser profile is stored at:

```text
%USERPROFILE%\.m365-copilot-openai-proxy\edge-profile
```


## API Endpoints

| Endpoint | Description |
|---|---|
| `GET /healthz` | Service health plus token status |
| `GET /v1/token/status` | Token validity, expiry time, and seconds remaining |
| `GET /v1/models` | OpenAI-compatible model list |
| `POST /v1/chat/completions` | OpenAI Chat Completions, streaming supported |
| `POST /v1/responses` | OpenAI Responses API, streaming supported |
| `POST /v1/messages` | Anthropic Messages API style endpoint |


## License

Apache License 2.0. See [LICENSE](LICENSE).
