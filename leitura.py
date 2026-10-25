"""
Leitura dos dados da DUIMP a partir do texto das telas do Portal Unico.

O portal muda o HTML com frequencia, mas o texto visivel quase nao muda. Por isso
os dados que nao tem um campo estavel sao lidos do texto da pagina com expressoes
regulares, e estas funcoes podem ser testadas sem abrir o navegador.
"""

import re

DUIMP_PATTERN = re.compile(r"\d{2}BR\d{10}-\d")

# Prefixos que a empresa usa nas referencias internas, ex.: "REF. IMP 1099B"
PREFIXOS_REFERENCIA = ("IMP",)

RE_NCM = re.compile(r"\b\d{4}\.\d{2}\.\d{2}\b")
RE_PESO = re.compile(r"Peso\s+bruto[^\n\d]*[\s:]*([\d.]+,\d+|[\d.,]+)", re.IGNORECASE)
RE_DATA_REGISTRO = re.compile(r"Data de Registro[:\s]*(\d{2}/\d{2}/\d{4})", re.IGNORECASE)
RE_CNPJ = re.compile(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}")
RE_ADQUIRENTE = re.compile(r"(?:ENCOMENDANTE|ADQUIRENTE|IMPORTADOR):\s*([^\n\r\-]+)", re.IGNORECASE)
RE_IDENTIFICACAO_DI = re.compile(r"IDENTIFICA[ÇC][AÃ]O\s*DI[\.:\s]+([^\n]+)", re.IGNORECASE)
RE_REF_CLIENTE = re.compile(r"REF\s*CLIENTE[\.:\s]+([^\n]+)", re.IGNORECASE)


def numero_br(texto):
    """Converte numero no formato brasileiro (1.234,56) para float. Vazio ou invalido vira 0."""
    if not texto:
        return 0.0
    limpo = re.sub(r"[^\d,.\-]", "", str(texto))
    if not limpo:
        return 0.0
    if "," in limpo and "." in limpo:
        limpo = limpo.replace(".", "").replace(",", ".")
    elif "," in limpo:
        limpo = limpo.replace(",", ".")
    try:
        return float(limpo)
    except ValueError:
        return 0.0


def numeros_duimp(texto):
    """Numeros de DUIMP no texto, sem repetir e na ordem em que aparecem."""
    return list(dict.fromkeys(DUIMP_PATTERN.findall(texto or "")))


def ncms_no_texto(texto):
    """O portal sempre mostra a NCM formatada (4011.10.00). Procurar o padrao no texto
    e mais confiavel do que achar a coluna certa numa tabela que muda de estrutura."""
    return list(dict.fromkeys(RE_NCM.findall(texto or "")))


def peso_no_texto(texto):
    m = RE_PESO.search(texto or "")
    return m.group(1).strip() if m else ""


def data_registro_no_texto(texto):
    m = RE_DATA_REGISTRO.search(texto or "")
    return m.group(1) if m else ""


def primeiro_cnpj(texto):
    m = RE_CNPJ.search(texto or "")
    return m.group(0) if m else ""


def dados_complementares(texto, prefixos=PREFIXOS_REFERENCIA):
    """Nome do adquirente e referencia interna, lidos das informacoes complementares.

    A referencia e o prefixo da empresa mais o codigo que vem logo depois de "REF.".
    Sem ele, vale a identificacao da DI ou a referencia do cliente.
    """
    texto = texto or ""
    nome, referencia = "", ""
    m = RE_ADQUIRENTE.search(texto)
    if m:
        nome = re.sub(r"\s+", " ", m.group(1).strip())
    prefixo = "|".join(re.escape(p) for p in prefixos)
    m = (re.search(rf"REF\.?\s*((?:{prefixo})\s*[A-Za-z0-9\-/]+)", texto, re.IGNORECASE)
         or RE_IDENTIFICACAO_DI.search(texto)
         or RE_REF_CLIENTE.search(texto))
    if m:
        referencia = re.sub(r"\s+", " ", m.group(1).strip())
    return nome, referencia
