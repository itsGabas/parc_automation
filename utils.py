import os
import re
import time
import logging
import unicodedata
from datetime import datetime
import pandas as pd
from openpyxl.utils import get_column_letter
from selenium.webdriver.common.action_chains import ActionChains
from config import CAMINHO_PLANILHA, PASTA_BASE, PASTA_DOWNLOADS


# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------

MESES_TEXTO = {
    1: "JANEIRO", 2: "FEVEREIRO", 3: "MARÇO",  4: "ABRIL",
    5: "MAIO",    6: "JUNHO",     7: "JULHO",   8: "AGOSTO",
    9: "SETEMBRO", 10: "OUTUBRO", 11: "NOVEMBRO", 12: "DEZEMBRO",
}

MESES_ABREV = {
    1: "JAN", 2: "FEV", 3: "MAR", 4: "ABR",
    5: "MAI", 6: "JUN", 7: "JUL", 8: "AGO",
    9: "SET", 10: "OUT", 11: "NOV", 12: "DEZ",
}

MESES_NUMERO = {v: str(i).zfill(2) for i, v in MESES_ABREV.items()}


# ---------------------------------------------------------------------------
# Planilha — carregamento
# ---------------------------------------------------------------------------

def carregar_dados_planilha_mensal(converters: dict | None = None):
    """
    Lê a aba do mês atual da planilha-mãe e identifica as colunas
    de tipo e número de parcelamento.

    Parâmetros
    ----------
    converters : dict, opcional
        Passado diretamente ao pd.read_excel (útil quando um módulo
        precisa de conversão específica de CNPJ, por exemplo).

    Retorna
    -------
    tuple: (dados, coluna_tipo, coluna_numero, mes_atual_texto)
        Todos os valores são None em caso de falha.
    """
    try:
        hoje = datetime.now()
        mes_atual_texto = MESES_TEXTO[hoje.month]
        nome_aba = f"{mes_atual_texto} {hoje.year}"

        if not os.path.exists(CAMINHO_PLANILHA):
            logging.error(f"Planilha não encontrada: {CAMINHO_PLANILHA}")
            return None, None, None, None

        logging.info(f"Carregando dados da aba '{nome_aba}'")

        kwargs = {"sheet_name": nome_aba}
        if converters:
            kwargs["converters"] = converters

        try:
            dados = pd.read_excel(CAMINHO_PLANILHA, **kwargs)
            dados.columns = [str(c).strip() for c in dados.columns]
            logging.info(f"Registros carregados: {len(dados)}")
        except Exception as e:
            logging.error(f"Erro ao carregar aba '{nome_aba}': {e}")
            return None, None, None, None

        colunas = list(dados.columns)
        logging.info(f"Colunas encontradas: {colunas}")

        coluna_tipo = next(
            (c for c in colunas if "tipo" in c.lower() and "parcelamento" in c.lower()),
            None,
        )
        coluna_numero = next(
            (
                c for c in colunas
                if "parcelamento" in c.lower()
                and any(t in c.lower() for t in ["nº", "no", "numero", "n°", "n ", "nr"])
            ),
            None,
        )

        # Fallback por posição caso as colunas não sejam encontradas por nome
        if not coluna_tipo and len(colunas) >= 4:
            coluna_tipo = colunas[3]
            logging.warning(f"Coluna de tipo identificada por posição: '{coluna_tipo}'")

        if not coluna_numero and len(colunas) >= 5:
            coluna_numero = colunas[4]
            logging.warning(f"Coluna de número identificada por posição: '{coluna_numero}'")

        if not coluna_tipo or not coluna_numero:
            logging.error("Colunas necessárias não encontradas na planilha.")
            return None, None, None, None

        logging.info(f"Coluna tipo: '{coluna_tipo}' | Coluna número: '{coluna_numero}'")
        return dados, coluna_tipo, coluna_numero, mes_atual_texto

    except Exception as e:
        logging.error(f"Erro ao processar a planilha: {e}", exc_info=True)
        return None, None, None, None


# ---------------------------------------------------------------------------
# Planilha — formatação
# ---------------------------------------------------------------------------

def ajustar_largura_colunas(sheet) -> None:
    """Ajusta a largura de cada coluna com base no maior conteúdo da célula."""
    for column_cells in sheet.columns:
        col_letter = get_column_letter(column_cells[0].column)
        max_length = 0
        for cell in column_cells:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except Exception:
                pass
        sheet.column_dimensions[col_letter].width = max_length + 2


# ---------------------------------------------------------------------------
# Selenium — interação robusta
# ---------------------------------------------------------------------------

def click_element_safely(driver, element) -> None:
    """
    Tenta clicar em um elemento com múltiplas estratégias de fallback:
    1. Clique direto
    2. JavaScript click
    3. scrollIntoView + clique direto
    4. ActionChains
    """
    try:
        element.click()
        return
    except Exception:
        pass

    try:
        driver.execute_script("arguments[0].click();", element)
        return
    except Exception:
        pass

    try:
        driver.execute_script("arguments[0].scrollIntoView(true);", element)
        time.sleep(0.5)
        element.click()
        return
    except Exception:
        pass

    try:
        ActionChains(driver).move_to_element(element).pause(0.5).click().perform()
    except Exception as e:
        logging.error(f"Todas as tentativas de clique falharam: {e}")


# ---------------------------------------------------------------------------
# Texto — normalização e extração
# ---------------------------------------------------------------------------

def normalizar_texto(txt: str) -> str:
    """Remove acentos e converte para minúsculas ASCII."""
    return unicodedata.normalize("NFKD", txt.lower()).encode("ascii", "ignore").decode("ascii")


def extrair_numero_parcelamento(texto) -> str | None:
    """
    Extrai o número do parcelamento de strings no formato '12345 / CDA...'.
    Retorna o texto original se o padrão não for encontrado.
    """
    try:
        if pd.isna(texto):
            return None
        texto = str(texto)
        match = re.search(r'^(\d+)\s*/\s*CDA', texto)
        return match.group(1) if match else texto
    except Exception as e:
        logging.error(f"Erro ao extrair número do parcelamento: {e}")
        return str(texto)


def extrair_numero_cda(texto) -> str | None:
    """
    Extrai o número da CDA de strings como 'CDA: 999999' ou retorna
    o valor original se já for um número puro.
    """
    try:
        if pd.isna(texto):
            return None
        texto = str(texto)
        match = re.search(r'CDA:?\s*(\d+)', texto, re.IGNORECASE)
        if match:
            return match.group(1)
        if texto.strip().isdigit():
            return texto.strip()
        return texto
    except Exception as e:
        logging.error(f"Erro ao extrair número CDA: {e}")
        return str(texto)


# ---------------------------------------------------------------------------
# Data — utilitários de mês
# ---------------------------------------------------------------------------

def mes_abrev_atual() -> str:
    """Retorna a abreviação do mês atual em maiúsculas. Ex: 'JUN'"""
    return MESES_ABREV[datetime.now().month]


def numero_mes_atual() -> str:
    """Retorna o número do mês atual com zero à esquerda. Ex: '06'"""
    return str(datetime.now().month).zfill(2)
