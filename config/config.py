from pathlib import Path
import platform

# Detecta o sistema operacional
IS_WINDOWS = platform.system() == 'Windows'
IS_MAC = platform.system() == 'Darwin'

# Define caminho base do Google Drive
if IS_WINDOWS:
    GOOGLE_DRIVE_PATH = Path("C:/Users/nome_usuario/Google Drive")
elif IS_MAC:
    GOOGLE_DRIVE_PATH = Path("/Users/emiliamargareth/Library/CloudStorage/GoogleDrive-emilia.mga@gmail.com/Meu Drive")


# Define os caminhos base para diferentes pastas
BASE_PATH = GOOGLE_DRIVE_PATH / "Vasconcelos_Rinaldi/Planilhas_Base"
PASTA_CLIENTES = GOOGLE_DRIVE_PATH / "Vasconcelos_Rinaldi/Clientes"

# Define caminhos específicos
ARQUIVO_CLIENTES = BASE_PATH / "clientes.xlsx"
ARQUIVO_FORNECEDORES = BASE_PATH / "base_fornecedores.xlsx"
ARQUIVO_MODELO = BASE_PATH / "MODELO.xlsx"
ARQUIVO_CONTROLE = BASE_PATH / "controle_taxa_adm.xlsx"
PASTA_CLIENTES = BASE_PATH.parent / "Clientes"

# Função para verificar se os arquivos existem
def verificar_arquivos():
    """Verifica se todos os arquivos necessários estão acessíveis"""
    arquivos = [ARQUIVO_CLIENTES, ARQUIVO_FORNECEDORES, ARQUIVO_MODELO, ARQUIVO_CONTROLE]
    for arquivo in arquivos:
        if not arquivo.exists():
            raise FileNotFoundError(f"Arquivo não encontrado: {arquivo}")
