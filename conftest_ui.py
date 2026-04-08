"""Shared fixtures for Playwright UI tests.

Extracts the HTML from main.py, injects a mock pywebview.api backed by the
real Python Api class (via a tiny HTTP bridge), and serves everything on localhost.
"""
import json
import threading
import pytest
from http.server import HTTPServer, BaseHTTPRequestHandler
from unittest.mock import MagicMock
import sys

import db_manager

# Mock modules that main.py imports at module level
sys.modules.setdefault("webview", MagicMock())
sys.modules.setdefault("pystray", MagicMock())

import main
from main import Api


def _extract_html():
    """Pull the HTML string from main.py."""
    return main.HTML


def _build_mock_api_js(port):
    """Return a JS snippet that defines window.pywebview.api as a proxy that
    calls back to our Python HTTP bridge for every method invocation."""
    return f"""
    <script>
    (function() {{
        const handler = {{
            get(target, prop) {{
                if (prop === 'then') return undefined;  // not a thenable
                return async function(...args) {{
                    const resp = await fetch('http://localhost:{port}/api', {{
                        method: 'POST',
                        headers: {{'Content-Type': 'application/json'}},
                        body: JSON.stringify({{method: prop, args: args}})
                    }});
                    return resp.json();
                }};
            }}
        }};
        window.pywebview = {{ api: new Proxy({{}}, handler) }};
        // Fire pywebviewready so initApp proceeds
        window.dispatchEvent(new Event('pywebviewready'));
    }})();
    </script>
    """


class _ApiBridgeHandler(BaseHTTPRequestHandler):
    """Tiny HTTP handler that forwards JS calls to the real Api instance."""
    api_instance = None
    db_path = None

    def do_POST(self):
        length = int(self.headers.get("Content-Length", 0))
        body = json.loads(self.rfile.read(length))
        method_name = body["method"]
        args = body.get("args", [])

        # Ensure DB path is set (test isolation)
        if self.db_path:
            db_manager.DB_PATH = self.db_path

        fn = getattr(self.api_instance, method_name, None)
        if fn is None:
            result = None
        else:
            try:
                result = fn(*args)
            except Exception as e:
                result = {"ok": False, "error": str(e)}

        self.send_response(200)
        self.send_header("Content-Type", "application/json")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(json.dumps(result if result is not None else None).encode())

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def log_message(self, format, *args):
        pass  # silence logs


class _PageHandler(BaseHTTPRequestHandler):
    """Serves the app HTML with the mock API injected."""
    html_content = ""

    def do_GET(self):
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.end_headers()
        self.wfile.write(self.html_content.encode())

    def log_message(self, format, *args):
        pass


@pytest.fixture(scope="session")
def _api_bridge_server():
    """Start the API bridge HTTP server once for the whole test session."""
    server = HTTPServer(("127.0.0.1", 0), _ApiBridgeHandler)
    port = server.server_address[1]
    t = threading.Thread(target=server.serve_forever, daemon=True)
    t.start()
    yield port
    server.shutdown()


@pytest.fixture(scope="session")
def _page_server(_api_bridge_server):
    """Start the page-serving HTTP server once for the whole test session."""
    raw_html = _extract_html()
    mock_js = _build_mock_api_js(_api_bridge_server)
    # Inject mock API script right after <head> so it runs before any other JS
    html = raw_html.replace("<head>", "<head>" + mock_js, 1)
    _PageHandler.html_content = html

    server = HTTPServer(("127.0.0.1", 0), _PageHandler)
    port = server.server_address[1]
    t = threading.Thread(target=server.serve_forever, daemon=True)
    t.start()
    yield port
    server.shutdown()


@pytest.fixture(autouse=True)
def ui_db(tmp_path, _api_bridge_server):
    """Fresh temp database for every UI test."""
    db_file = str(tmp_path / "ui_test.db")
    db_manager.DB_PATH = db_file
    db_manager.init_db()
    api = Api()
    _ApiBridgeHandler.api_instance = api
    _ApiBridgeHandler.db_path = db_file
    yield api


@pytest.fixture()
def app_url(_page_server):
    """The URL to load in the browser."""
    return f"http://localhost:{_page_server}"
