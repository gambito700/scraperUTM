import http.server
import json
import subprocess
import os
import sys
import threading
import webbrowser
import urllib.parse
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_FILE = os.path.join(BASE_DIR, "dashboard_data.json")
SCRAPER = os.path.join(BASE_DIR, "scraper.py")
LOG_FILE = os.path.join(BASE_DIR, "scraper_log.txt")
EXCEL_FILE = os.path.join(BASE_DIR, "indicadores_previsionales.xlsx")
PORT = 8080


class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=BASE_DIR, **kwargs)

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        route = parsed.path
        if route == "/api/data":
            self.send_json_file(JSON_FILE)
        elif route == "/api/refresh":
            self.run_scraper()
        elif route == "/api/status":
            self.send_json(200, {"ok": True, "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
        elif route == "/favicon.ico":
            self.send_response(204)
            self.end_headers()
        elif route == "/api/logs":
            self.send_logs()
        elif route == "/api/download/excel":
            self.download_excel()
        else:
            super().do_GET()

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "*")
        self.end_headers()

    def send_json(self, status, data):
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(json.dumps(data, ensure_ascii=False).encode("utf-8"))

    def send_json_file(self, path):
        if not os.path.exists(path):
            self.send_json(404, {"error": "Archivo no encontrado. Ejecute el scraper primero."})
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                content = f.read()
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(content.encode("utf-8"))
        except Exception as e:
            self.send_json(500, {"error": str(e)})

    def send_logs(self):
        if not os.path.exists(LOG_FILE):
            self.send_json(200, {"lines": [], "total": 0, "message": "No hay logs aún"})
            return
        all_lines = []
        for enc in ("utf-8", "utf-16", "latin-1"):
            try:
                with open(LOG_FILE, "r", encoding=enc) as f:
                    all_lines = f.readlines()
                break
            except (UnicodeDecodeError, UnicodeError):
                continue
        if not all_lines:
            self.send_json(500, {"error": "No se pudo decodificar el archivo de logs"})
            return
        last_lines = all_lines[-100:]
        self.send_json(200, {"lines": last_lines, "total": len(all_lines)})

    def download_excel(self):
        if not os.path.exists(EXCEL_FILE):
            self.send_json(404, {"error": "Excel no encontrado. Ejecute el scraper primero."})
            return
        try:
            with open(EXCEL_FILE, "rb") as f:
                content = f.read()
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", 'attachment; filename="indicadores_previsionales.xlsx"')
            self.send_header("Access-Control-Allow-Origin", "*")
            self.send_header("Content-Length", str(len(content)))
            self.end_headers()
            self.wfile.write(content)
        except Exception as e:
            self.send_json(500, {"error": str(e)})

    def run_scraper(self):
        def _run():
            try:
                result = subprocess.run(
                    [sys.executable, SCRAPER],
                    capture_output=True, text=True, timeout=120
                )
                self.scraper_result = {
                    "success": result.returncode == 0,
                    "stdout": result.stdout,
                    "stderr": result.stderr,
                    "returncode": result.returncode
                }
            except subprocess.TimeoutExpired:
                self.scraper_result = {"success": False, "error": "Timeout"}
            except Exception as e:
                self.scraper_result = {"success": False, "error": str(e)}

        thread = threading.Thread(target=_run, daemon=True)
        thread.start()
        thread.join()

        ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        sr = self.scraper_result
        if sr.get("success"):
            log_entry = f"[{ts}] Scraping completado exitosamente\n"
        else:
            err = sr.get("error") or sr.get("stderr", "Error desconocido")
            log_entry = f"[{ts}] ERROR: {err}\n"
        try:
            with open(LOG_FILE, "a", encoding="utf-8") as f:
                if sr.get("stdout"):
                    f.write(sr["stdout"] + "\n")
                if sr.get("stderr"):
                    f.write(f"[{ts}] STDERR: {sr['stderr']}\n")
                f.write(log_entry)
        except Exception:
            pass

        if self.scraper_result.get("success"):
            self.send_json(200, {
                "success": True,
                "message": "Datos actualizados correctamente",
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            })
        else:
            self.send_json(500, {
                "success": False,
                "error": self.scraper_result.get("error") or self.scraper_result.get("stderr", ""),
                "stdout": self.scraper_result.get("stdout", "")
            })

    def log_message(self, format, *args):
        msg = format % args if args else format
        print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")


if __name__ == "__main__":
    server = http.server.HTTPServer(("0.0.0.0", PORT), Handler)
    url = f"http://localhost:{PORT}/dashboard.html"
    print(f" Servidor iniciado en http://localhost:{PORT}")
    print(f" Dashboard: {url}")
    print(f" API Refresh: http://localhost:{PORT}/api/refresh")
    print(" Presiona Ctrl+C para detener")

    webbrowser.open(url)

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n Servidor detenido")
        server.server_close()
