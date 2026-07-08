import asyncio
import io
import os
from pathlib import Path
from urllib.parse import unquote

import httpx
import pytest
from fastapi.testclient import TestClient

os.environ.setdefault("MPLCONFIGDIR", "/tmp/matplotlib-cache")

from app.main import app
from app.routes import reports
from app.services.report_session_store import ReportSessionStore


class _SyncASGITransport(httpx.BaseTransport):
    def __init__(self, app):
        self.app = app

    def handle_request(self, request: httpx.Request) -> httpx.Response:
        body = request.read()
        response_started = False
        response_complete: asyncio.Event | None = None
        response_status = 500
        response_headers: list[tuple[bytes, bytes]] = []
        response_body = io.BytesIO()

        async def call_app() -> None:
            nonlocal response_complete
            response_complete = asyncio.Event()
            request_complete = False

            scope = {
                "type": "http",
                "asgi": {"version": "3.0", "spec_version": "2.4"},
                "http_version": "1.1",
                "method": request.method,
                "path": unquote(request.url.path),
                "raw_path": request.url.raw_path.split(b"?", 1)[0],
                "root_path": "",
                "scheme": request.url.scheme,
                "query_string": request.url.query,
                "headers": [
                    (key.lower().encode(), value.encode())
                    for key, value in request.headers.multi_items()
                ],
                "client": ("testclient", 50000),
                "server": (request.url.host, request.url.port),
                "state": {},
            }

            async def receive():
                nonlocal request_complete

                if request_complete:
                    await response_complete.wait()
                    return {"type": "http.disconnect"}

                request_complete = True
                return {"type": "http.request", "body": body}

            async def send(message):
                nonlocal response_started, response_status, response_headers

                if message["type"] == "http.response.start":
                    response_started = True
                    response_status = message["status"]
                    response_headers = message.get("headers", [])
                    return

                if message["type"] == "http.response.body":
                    response_body.write(message.get("body", b""))
                    if not message.get("more_body", False):
                        response_complete.set()

            await self.app(scope, receive, send)

        asyncio.run(call_app())

        assert response_started, "TestClient did not receive any response."
        return httpx.Response(
            status_code=response_status,
            headers=response_headers,
            stream=httpx.ByteStream(response_body.getvalue()),
            request=request,
        )


class SandboxTestClient(TestClient):
    def __init__(self, app):
        httpx.Client.__init__(
            self,
            base_url="http://testserver",
            headers={"user-agent": "testclient"},
            transport=_SyncASGITransport(app),
            follow_redirects=True,
        )


@pytest.fixture
def temp_store(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    store = ReportSessionStore(tmp_path / "storage")
    monkeypatch.setattr(reports, "report_session_store", store)
    yield store


@pytest.fixture
def client(temp_store, monkeypatch: pytest.MonkeyPatch):
    monkeypatch.setenv("ADMIN_PASSWORD", "test-admin-password")
    return SandboxTestClient(app)
