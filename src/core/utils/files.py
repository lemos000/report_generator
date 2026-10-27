import os
from typing import List


def read_lines_tolerant(file_path: str) -> List[str]:
    """Lê um arquivo texto tentando utf-8 e, em caso de falha, latin1."""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Arquivo não encontrado: {file_path}")

    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return f.readlines()
    except UnicodeDecodeError:
        with open(file_path, 'r', encoding='latin1') as f:
            return f.readlines()
    except Exception as e:
        raise ValueError(f"Erro ao ler CSV {os.path.basename(file_path)}: {str(e)}")
