import logging
import sys
import os

# Para funcionar no executável
if getattr(sys, 'frozen', False):
    script_dir = os.path.dirname(sys.executable)
else:
    script_dir = os.path.dirname(__file__)

sys.path.insert(0, script_dir)

# Importar as funções main de cada módulo
from da_sp_att import main as da_sp_main  
from mg_atual import main as mg_main
from sefaz_rs_att import main as rs_main

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler()
    ]
)

try:
    
    logging.info("Baixando guias estaduais referentes a D.A - SP...")
    da_sp_main()
    
    logging.info("Baixando guias estaduais referentes a MG...")
    mg_main()
    
    logging.info("Baixando guias estaduais referentes a RS...")
    rs_main()
    
    logging.info("✅ Processamento completo!")
    input("Pressione ENTER para fechar...")
    
except Exception as e:
    logging.error(f"❌ Erro durante execução: {e}")
    input("Pressione ENTER para fechar...")
