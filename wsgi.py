from asgiref.wsgi import WsgiToAsgi
from Pasted_code import app   # 👈 change if filename is different

# Convert Flask (WSGI) → ASGI
asgi_app = WsgiToAsgi(app)