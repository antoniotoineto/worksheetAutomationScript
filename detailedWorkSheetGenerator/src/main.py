# =========================
# MAIN
# =========================

from config import validate_files
from generator import create_files
from infos import fill_infos
from vilt import filter_vilt

if __name__ == '__main__':
    validate_files()
    fill_infos()
    filter_vilt()
    create_files()
