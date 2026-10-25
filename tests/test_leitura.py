import pytest

from leitura import (dados_complementares, data_registro_no_texto, ncms_no_texto, numero_br,
                     numeros_duimp, peso_no_texto, primeiro_cnpj)


@pytest.mark.parametrize("texto, esperado", [
    ("1.234,56", 1234.56),
    ("12.345,678 kg", 12345.678),
    ("1234.5", 1234.5),
    ("350,00", 350.0),
    ("", 0.0),
    ("-", 0.0),
    (None, 0.0),
])
def test_numero_br(texto, esperado):
    assert numero_br(texto) == pytest.approx(esperado)


def test_numeros_duimp_sem_repetir_e_na_ordem():
    texto = "26BR0000012345-6 Registrada\n26BR0000067890-1\n26BR0000012345-6 (repetida)"
    assert numeros_duimp(texto) == ["26BR0000012345-6", "26BR0000067890-1"]


def test_ncms_no_texto():
    texto = "Item 1 NCM 4011.10.00 Pneu\nItem 2 NCM 4011.20.90\nItem 3 NCM 4011.10.00"
    assert ncms_no_texto(texto) == ["4011.10.00", "4011.20.90"]


@pytest.mark.parametrize("texto, esperado", [
    ("Peso bruto (kg)\n12.345,678", "12.345,678"),
    ("Peso bruto total: 8.900,5", "8.900,5"),
    ("Peso liquido 100", ""),
])
def test_peso_no_texto(texto, esperado):
    assert peso_no_texto(texto) == esperado


def test_data_registro():
    assert data_registro_no_texto("Situação: Desembaraçada\nData de Registro: 14/03/2026") == "14/03/2026"
    assert data_registro_no_texto("sem data") == ""


def test_primeiro_cnpj():
    assert primeiro_cnpj("<span>12.345.678/0001-90</span><span>98.765.432/0001-10</span>") == "12.345.678/0001-90"


def test_nome_e_referencia_interna():
    texto = "ENCOMENDANTE: COMERCIAL EXEMPLO LTDA\nREF. IMP 1099B - CONTAINER TSTU1234567"
    assert dados_complementares(texto) == ("COMERCIAL EXEMPLO LTDA", "IMP 1099B")


def test_prefixos_de_referencia_configuraveis():
    texto = "ADQUIRENTE: OUTRA EMPRESA SA\nREF. PRC-0042"
    assert dados_complementares(texto, prefixos=("PRC",))[1] == "PRC-0042"


def test_sem_referencia_usa_identificacao_da_di():
    texto = "IMPORTADOR: EMPRESA X\nIDENTIFICAÇÃO DI: 2026/0042"
    assert dados_complementares(texto) == ("EMPRESA X", "2026/0042")


def test_texto_vazio():
    assert dados_complementares("") == ("", "")
