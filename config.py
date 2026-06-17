"""
config.py
Carrega as variáveis de ambiente do arquivo .env e as expõe como
constantes para uso em todo o projeto.
"""

import os
import logging
from pathlib import Path
from dotenv import load_dotenv

# Garante que o .env é encontrado mesmo quando o script roda de outro diretório
BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")


def _obrigatorio(chave: str) -> str:
    """Lê uma variável obrigatória — encerra com erro claro se estiver ausente."""
    valor = os.getenv(chave)
    if not valor:
        logging.critical(
            f"Variável de ambiente '{chave}' não encontrada. "
            f"Verifique se o arquivo .env existe e está preenchido corretamente."
        )
        raise EnvironmentError(f"Variável obrigatória ausente: {chave}")
    return valor


# --- Caminhos ---
CAMINHO_PLANILHA   = _obrigatorio("CAMINHO_PLANILHA")
PASTA_DOWNLOADS    = _obrigatorio("PASTA_DOWNLOADS")
PASTA_BASE = _obrigatorio("PASTA_BASE")

# --- Serviços externos ---
TWOCAPTCHA_API_KEY = _obrigatorio("TWOCAPTCHA_API_KEY")
