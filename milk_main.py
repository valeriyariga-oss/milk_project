import os
from flask import Flask
from dotenv import load_dotenv
from db import db
from routes import routes_bp
from settings import settings_bp
from export import export_bp

load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY", "change-this-development-key")

# Регистрация Blueprint'ов
app.register_blueprint(routes_bp)
app.register_blueprint(settings_bp)
app.register_blueprint(export_bp)

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=int(os.getenv("PORT", "5000")))