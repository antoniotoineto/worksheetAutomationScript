# =========================
# CONFIGURAÇÕES
# =========================

import os
import sys
from logger import error, log

def base_path():
    if getattr(sys, 'frozen', False):
        # Se estiver rodando como .exe
        return os.path.dirname(sys.executable)
    else:
        # Se estiver rodando como .py
        return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

BASE_DIR = base_path()

PATHS = {
    'template': os.path.join(BASE_DIR, 'base', 'template.xlsx'),
    'infos': os.path.join(BASE_DIR, 'base', 'infos.xlsx'),
    'vilt': os.path.join(BASE_DIR, 'base', 'vilt.xlsx'),
    'output': os.path.join(BASE_DIR, 'output')
}

os.makedirs(PATHS['output'], exist_ok=True)

def validate_files():
    for k, path in PATHS.items():
        if k != 'output' and not os.path.exists(path):
            error(f"Arquivo não encontrado: {path}")
    log("Arquivos encontrados com sucesso")