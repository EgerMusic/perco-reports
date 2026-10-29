"""Serve the PERCo report app on this computer only."""

import argparse
import functools
import http.server
from pathlib import Path
import urllib.request
import webbrowser


class LocalServer(http.server.ThreadingHTTPServer):
    allow_reuse_address = False
    allow_reuse_port = False


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--no-browser", action="store_true")
    args = parser.parse_args()
    app_dir = Path(__file__).resolve().parent / "dist"
    handler = functools.partial(http.server.SimpleHTTPRequestHandler, directory=str(app_dir))
    url = "http://localhost:8765/"
    try:
        server = LocalServer(("127.0.0.1", 8765), handler)
    except OSError as exc:
        try:
            with urllib.request.urlopen(url, timeout=2) as response:
                page = response.read(8192)
            if b'PERCo' in page and b'files' in page:
                print(f"PERCo service is already running at {url}", flush=True)
                if not args.no_browser:
                    webbrowser.open_new_tab(url)
                return
        except Exception:
            pass
        raise SystemExit(f"Cannot start on localhost:8765: {exc}")
    print(f"PERCo service is running at {url}", flush=True)
    print("Keep this window open. Press Ctrl+C to stop the service.", flush=True)
    if not args.no_browser:
        webbrowser.open_new_tab(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("Stopping the service.", flush=True)
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
