"""
Teste com os dados exatos enviados pelo usuário (sem pipes, com quebras de linha puras).
"""
from parser_ssw import parse_ssw_table

USER_INPUT = """
CTRC
Emissão
Remetente
Destinatário
Pagador
Destino
N Fiscal
Peso
Frete
Pv Entr
Ocorrência
Atraso
Localização
EUC513370-0
01/09/26
VIA VAREJO S-A
VIA VAREJO S/A
VIA VAREJO S-A
SSAI-CAMACARI
003602
72
0,01
08/09/26
84-CHEGADA NA UNIDADE DE DESTINO
20
MPC SSA - NO ARMAZEM DA UNIDADE MP CARGAS
SSA009990-2
03/09/26
FUNDACAO JOSE CARVALHO
GMO CENTRO DE PEQUISAS E CONTROLE DE QUALIDADE LTDA
FUNDACAO JOSE CARVALHO
SSAP-SALVADOR
032157
73
0,01
09/09/26
80-CTRC EMITIDO
19
MPC SSA-CT-e autorizado com 2 volumes e 73 Kg. Destino: BA/SALVADOR. Previsao de entrega: 09/09/26
MIM010342-0
09/09/26
HIPER LUCKY FEST COMERCIO EMBA
CMC PAPELARIA E ARMARINHO LTDA
AUCILENE ARAUJO DA SILVA-JR TR
SSAP-SALVADOR
002402
329
489,99
11/09/26
 
17
MPC MIM-Subcontrato autorizado com 17 volumes e 329 Kg. Destino: BA/SALVADOR. Previsao de entrega: 11/09/26
BAR513390-4
19/09/26
SIAN ENGENHARIA LTDA
MAQSERV EQUIPAMENTOS LTDA
SIAN ENGENHARIA LTDA
SSAI-LAURO DE FREITAS
002155
10
82,69
25/09/26
84-CHEGADA NA UNIDADE DE DESTINO
03
MPC SSA - NO 
Não foi possível identificar a tabela do SSW. Verifique se você copiou a tabela completa do SSW.
"""

def run_test():
    records, err = parse_ssw_table(USER_INPUT)
    print("Erro retornado:", err)
    print("Total de registros encontrados:", len(records))
    for r in records:
        print(f"-> CTRC: {r['CTRC']}, Destino: {r['DESTINO']}, Situação: {r['SITUAÇÃO']}, Atraso: {r['DIAS DE ATRASO']}, Frete: {r['FRETE']}")

if __name__ == '__main__':
    run_test()
