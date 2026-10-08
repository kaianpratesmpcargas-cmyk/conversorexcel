"""
Script de teste unitário e de integração para validar o parser e a geração do Excel.
"""
import os
from parser_ssw import parse_ssw_table
from excel_export import export_to_excel

# Amostra realista copiada do SSW com Markdown, URLs, caracteres invisíveis, &#xA0;, <br>, vazios, etc.
SAMPLE_SSW = """
| [CTRC](https://sistema.ssw.inf.br/bin/ssw0861#) | [Emissão](https://sistema.ssw.inf.br/bin/ssw0861#) | [Remetente](https://sistema.ssw.inf.br/bin/ssw0861#) | [Destinatário](https://sistema.ssw.inf.br/bin/ssw0861#) | [Pagador](https://sistema.ssw.inf.br/bin/ssw0861#) | [Destino](https://sistema.ssw.inf.br/bin/ssw0861#) | [**N Fiscal**](https://sistema.ssw.inf.br/bin/ssw0861#) | [Peso](https://sistema.ssw.inf.br/bin/ssw0861#) | [Frete](https://sistema.ssw.inf.br/bin/ssw0861#) | [Pv Entr](https://sistema.ssw.inf.br/bin/ssw0861#) | [Ocorrência](https://sistema.ssw.inf.br/bin/ssw0861#) | [Atraso](https://sistema.ssw.inf.br/bin/ssw0861#) | [Localização](https://sistema.ssw.inf.br/bin/ssw0861#) |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| [ETR513286-0](https://sistema.ssw.inf.br/bin/ssw0861#) | 12/08/24 | INDÚSTRIA ALFA LTDA | COMÉRCIO BETA S/A | INDÚSTRIA ALFA | SALVADOR | 032040 | 1.250,00 | 1.850,50 | 18/08/24 | MERCADORIA RETIDA SEFAZ | 47 | DEPÓSITO SALVADOR |
| [MIM007293-1](https://sistema.ssw.inf.br/bin/ssw0861#) | 15/08/24 | DISTRIBUIDORA GAMA&#xA0; | MERCADO DELTA &nbsp; | DISTRIBUIDORA GAMA | SIMÕES FILHO | 472010 | 820 | 950,00 | 20/08/24 | <br>AGUARDANDO CLIENTE<br> | 47 | TRANSITO SSA |
| [SSA008254-6](https://sistema.ssw.inf.br/bin/ssw0861#) | 20/08/24 | FABRICA DE PECAS S/A | LOGISTICA EXPRESS | FABRICA DE PECAS | CAMAÇARI | 003554 | 450,50 | 620,00 | 25/08/24 | | 35 | FILIAL CAMACARI |
| [JFS513335-1](https://sistema.ssw.inf.br/bin/ssw0861#) | 22/08/24 | TEXTIL DO BRASIL | VAREJO POPULAR | VAREJO POPULAR | SALVADOR | 088421 | 2.100,00 | 3.200,00 | 28/08/24 | ENTREGUE AO DESTINATARIO | 0 | BAIXADO |
| [PSA028240-5](https://sistema.ssw.inf.br/bin/ssw0861#) | 25/08/24 | METALURGICA NORDESTE | CONSTRUTORA HORIZONTE | CONSTRUTORA HORIZONTE | LAURO DE FREITAS | 014522 | 609,50 | 1.879,50 | 30/08/24 | | | EM ROTA DE ENTREGA |
| | | | | | | | | | | | | |
| × | Fechar |
"""

def test_full_pipeline():
    print("Iniciando teste do parser...")
    records, err = parse_ssw_table(SAMPLE_SSW)
    assert not err, f"Erro inesperado no parser: {err}"
    assert len(records) == 5, f"Esperado 5 registros, obteve {len(records)}"

    r1 = records[0]
    assert r1['CTRC'] == 'ETR513286-0', f"CTRC inválido: {r1['CTRC']}"
    assert r1['SITUAÇÃO'] == 'ATRASADO', f"Situação r1 incorreta: {r1['SITUAÇÃO']}"
    assert r1['PESO'] == 1250.0, f"Peso incorreto: {r1['PESO']}"
    assert r1['FRETE'] == 1850.50, f"Frete incorreto: {r1['FRETE']}"
    assert r1['DESTINO'] == 'SALVADOR'
    assert r1['DIAS DE ATRASO'] == 47

    r2 = records[1]
    assert r2['REMETENTE'] == 'DISTRIBUIDORA GAMA', f"HTML unescape falhou: '{r2['REMETENTE']}'"
    assert r2['OCORRÊNCIA'] == 'AGUARDANDO CLIENTE', f"<br> remoção falhou: '{r2['OCORRÊNCIA']}'"

    r4 = records[3]
    assert r4['SITUAÇÃO'] == 'NO PRAZO', f"Atraso 0 deveria ser NO PRAZO, obteve: {r4['SITUAÇÃO']}"

    r5 = records[4]
    assert r5['SITUAÇÃO'] == 'SEM INFORMAÇÃO', f"Atraso vazio deveria ser SEM INFORMAÇÃO, obteve: {r5['SITUAÇÃO']}"
    assert r5['DIAS DE ATRASO'] is None

    print("Parser validado com sucesso!")

    test_excel_path = "test_relatorio.xlsx"
    print(f"Testando exportação para {test_excel_path}...")
    export_to_excel(records, test_excel_path)
    assert os.path.exists(test_excel_path), "Arquivo Excel não foi criado!"
    print(f"Excel gerado com tamanho {os.path.getsize(test_excel_path)} bytes.")

    # Remover arquivo de teste
    os.remove(test_excel_path)
    print("Teste concluído com 100% de sucesso!")

if __name__ == '__main__':
    test_full_pipeline()
