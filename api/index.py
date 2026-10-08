"""
Entrypoint serverless para Vercel — MP CARGAS Conversor SSW
A Vercel detecta o 'app' Flask exposto aqui como handler WSGI.
"""
import sys
import os

# Garante que os módulos na raiz do projeto sejam encontrados
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Re-exporta o app Flask a partir do web_app principal
from web_app import app
