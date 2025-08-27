import os

# Diretório onde ficam os dados, caso a variável não esteja setada, usa "./data"    
DATA_DIR = os.environ.get("VR_DATA_DIR", "./data")

# Arquivos de saída
RESULT_FILE = os.path.join(DATA_DIR, "VR_MENSAL_05_2025_RESULTADO.xlsx")
VALIDACAO_FILE = os.path.join(DATA_DIR, "VR_MENSAL_05_2025_VALIDACAO.xlsx")