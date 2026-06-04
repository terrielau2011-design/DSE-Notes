import http.server
import os

class UTF8Handler(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        if self.path.endswith('.html') or self.path == '/':
            self.send_header('Content-Type', 'text/html; charset=utf-8')
        super().end_headers()

os.chdir('/workspace')
http.server.test(HandlerClass=UTF8Handler, port=8000, bind='0.0.0.0')
