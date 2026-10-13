# -*- coding: utf-8 -*-
"""Base de regras: CFOP de entrada esperado a partir do CFOP do fornecedor, CST aceitos na escrituração
da entrada e a classificação de CFOP usada nas regras de crédito (Lucro Real, regime não cumulativo)."""

CFOP_DESC = {
    '1101': 'Compra para industrialização', '1102': 'Compra para comercialização', '1111': 'Compra p/ industrialização (consignação)',
    '1113': 'Compra p/ comercialização (consignação)', '1116': 'Compra p/ industrialização (entrega futura)', '1117': 'Compra p/ comercialização (entrega futura)',
    '1118': 'Compra p/ comercialização (entrega a terceiro)', '1120': 'Compra p/ industrialização (venda à ordem)', '1121': 'Compra p/ comercialização (venda à ordem)',
    '1122': 'Compra p/ industrialização (remessa direta)', '1124': 'Industrialização por outra empresa', '1126': 'Compra p/ prestação de serviço (ICMS)',
    '1128': 'Compra p/ prestação de serviço (ISS)', '1151': 'Transferência p/ industrialização', '1152': 'Transferência p/ comercialização',
    '1201': 'Devolução de venda de produção', '1202': 'Devolução de venda de mercadoria', '1352': 'Frete: estabelecimento industrial',
    '1353': 'Frete: estabelecimento comercial', '1401': 'Compra p/ industrialização com ST', '1403': 'Compra p/ comercialização com ST',
    '1406': 'Compra de ativo com ST', '1407': 'Compra p/ uso e consumo com ST', '1411': 'Devolução de venda com ST',
    '1551': 'Compra de ativo imobilizado', '1556': 'Compra de material de uso e consumo', '1653': 'Compra de combustível p/ consumo',
    '1910': 'Entrada de bonificação/brinde', '1911': 'Entrada de amostra grátis', '1915': 'Entrada p/ conserto', '1916': 'Retorno de conserto',
    '1949': 'Outra entrada não especificada', '1933': 'Aquisição de serviço (ISS)',
    '5101': 'Venda de produção', '5102': 'Venda de mercadoria', '5401': 'Venda de produção com ST (substituto)', '5403': 'Venda com ST (substituto)',
    '5405': 'Venda com ST (substituído)', '5551': 'Venda de ativo', '5910': 'Remessa em bonificação', '5911': 'Remessa de amostra',
    '5915': 'Remessa p/ conserto', '5916': 'Retorno de conserto', '5949': 'Outra saída não especificada', '5202': 'Devolução de compra',
    '5411': 'Devolução de compra com ST', '5353': 'Frete (transporte) a comercial', '5929': 'Venda já registrada em ECF', '5933': 'Serviço (ISS)',
}

CFOP_DESC.update({
    '1124': 'Industrialização efetuada por outra empresa', '1125': 'Industrialização por outra empresa (sem retorno)', '1151': 'Transferência p/ industrialização',
    '1153': 'Transferência de energia elétrica', '1154': 'Transferência p/ uso na prestação de serviço', '1203': 'Devolução de venda (ZFM/ALC)',
    '1204': 'Devolução de venda (ZFM/ALC, terceiros)', '1205': 'Anulação de serviço de comunicação', '1206': 'Anulação de serviço de transporte',
    '1208': 'Devolução de produção transferida', '1209': 'Devolução de mercadoria transferida', '1251': 'Energia elétrica p/ distribuição',
    '1252': 'Energia elétrica: estabelecimento industrial', '1253': 'Energia elétrica: estabelecimento comercial', '1254': 'Energia elétrica: prestador de transporte',
    '1255': 'Energia elétrica: prestador de comunicação', '1256': 'Energia elétrica: produtor rural', '1257': 'Energia elétrica p/ consumo',
    '1301': 'Comunicação: execução de serviço da mesma natureza', '1302': 'Comunicação: estabelecimento industrial', '1303': 'Comunicação: estabelecimento comercial',
    '1304': 'Comunicação: prestador de transporte', '1305': 'Comunicação: geradora/distribuidora de energia', '1306': 'Comunicação: produtor rural',
    '1351': 'Frete: execução de serviço da mesma natureza', '1354': 'Frete: prestador de comunicação', '1355': 'Frete: geradora/distribuidora de energia',
    '1356': 'Frete: produtor rural', '1360': 'Frete: contribuinte substituto', '1408': 'Transferência p/ industrialização com ST', '1409': 'Transferência p/ comercialização com ST',
    '1410': 'Devolução de venda de produção com ST', '1414': 'Retorno de produção remetida p/ venda fora com ST', '1415': 'Retorno de mercadoria remetida p/ venda fora com ST',
    '1552': 'Transferência de ativo imobilizado', '1553': 'Devolução de venda de ativo', '1554': 'Retorno de ativo remetido p/ uso fora', '1555': 'Entrada de ativo de terceiro p/ uso',
    '1557': 'Transferência de uso e consumo', '1651': 'Compra de combustível p/ industrialização', '1652': 'Compra de combustível p/ comercialização',
    '1901': 'Entrada p/ industrialização por encomenda', '1902': 'Retorno de mercadoria remetida p/ industrialização', '1903': 'Entrada de mercadoria remetida p/ industrialização e não aplicada',
    '1904': 'Retorno de remessa p/ venda fora', '1905': 'Entrada p/ depósito fechado/armazém', '1906': 'Retorno de depósito fechado/armazém', '1907': 'Retorno simbólico de depósito',
    '1908': 'Entrada de bem por comodato/locação', '1909': 'Retorno de bem remetido por comodato/locação', '1912': 'Entrada para demonstração', '1913': 'Retorno de demonstração',
    '1914': 'Retorno de exposição/feira', '1917': 'Entrada em consignação mercantil', '1918': 'Devolução de consignação mercantil', '1919': 'Devolução simbólica de consignação',
    '1920': 'Entrada de vasilhame/sacaria', '1921': 'Retorno de vasilhame/sacaria', '1922': 'Simples faturamento de compra p/ entrega futura', '1923': 'Mercadoria recebida por conta e ordem',
    '1924': 'Industrialização por conta e ordem', '1925': 'Retorno de industrialização por conta e ordem', '1926': 'Reclassificação de mercadoria (kit)',
    '1931': 'Crédito de ICMS-ST de transporte', '1932': 'Serviço de transporte iniciado em outra UF', '1934': 'Entrada simbólica de depósito fechado',
    '5103': 'Venda de produção fora do estabelecimento', '5104': 'Venda de mercadoria fora do estabelecimento', '5116': 'Venda de produção p/ entrega futura (remessa)',
    '5117': 'Venda de mercadoria p/ entrega futura (remessa)', '5118': 'Venda de produção entregue a terceiro', '5119': 'Venda de mercadoria entregue a terceiro',
    '5120': 'Venda à ordem (mercadoria de terceiro)', '5122': 'Venda de produção remetida p/ industrialização', '5124': 'Industrialização por encomenda',
    '5125': 'Industrialização por encomenda (sem retorno)', '5151': 'Transferência de produção', '5152': 'Transferência de mercadoria', '5201': 'Devolução de compra p/ industrialização',
    '5208': 'Devolução de produção transferida', '5209': 'Devolução de mercadoria transferida', '5251': 'Venda de energia elétrica', '5301': 'Prestação de comunicação',
    '5351': 'Transporte p/ execução de serviço da mesma natureza', '5352': 'Transporte a estabelecimento industrial', '5402': 'Venda de produção com ST (substituído)',
    '5408': 'Transferência de produção com ST', '5409': 'Transferência de mercadoria com ST', '5410': 'Devolução de compra p/ industrialização com ST',
    '5552': 'Transferência de ativo', '5553': 'Devolução de compra de ativo', '5554': 'Remessa de ativo p/ uso fora', '5555': 'Devolução de ativo de terceiro',
    '5556': 'Devolução de compra de uso e consumo', '5557': 'Transferência de uso e consumo', '5551': 'Venda de ativo imobilizado',
    '5655': 'Venda de combustível p/ consumidor ou usuário final', '5656': 'Venda de combustível de terceiros p/ consumidor final', '5667': 'Venda de combustível a consumidor de outra UF',
    '5901': 'Remessa p/ industrialização por encomenda', '5902': 'Retorno de industrialização por encomenda', '5903': 'Retorno de mercadoria recebida p/ industrialização e não aplicada',
    '5904': 'Remessa p/ venda fora do estabelecimento', '5905': 'Remessa p/ depósito fechado/armazém', '5906': 'Retorno de depósito fechado/armazém', '5907': 'Retorno simbólico de depósito',
    '5908': 'Remessa de bem por comodato/locação', '5909': 'Retorno de bem recebido por comodato/locação', '5912': 'Remessa p/ demonstração', '5913': 'Retorno de demonstração',
    '5914': 'Remessa p/ exposição/feira', '5917': 'Remessa em consignação mercantil', '5918': 'Devolução de consignação', '5919': 'Devolução simbólica de consignação',
    '5920': 'Remessa de vasilhame/sacaria', '5921': 'Devolução de vasilhame/sacaria', '5922': 'Simples faturamento de venda p/ entrega futura', '5923': 'Remessa por conta e ordem de terceiros',
    '5924': 'Remessa p/ industrialização por conta e ordem', '5925': 'Retorno de industrialização por conta e ordem', '5926': 'Reclassificação de mercadoria (kit)',
    '5927': 'Baixa de estoque (perda, roubo, deterioração)', '5928': 'Baixa por encerramento de atividade', '5929': 'Venda já registrada em cupom fiscal', '5931': 'Estorno de ICMS-ST de transporte',
    '5932': 'Transporte iniciado em outra UF', '5934': 'Remessa simbólica a depósito fechado',
})


def desc_cfop(c):
    c = str(c or '')
    if c in CFOP_DESC:
        return CFOP_DESC[c]
    if len(c) == 4 and c[0] in '2367':
        base = {'2': '1', '3': '1', '6': '5', '7': '5'}[c[0]] + c[1:]
        if base in CFOP_DESC:
            return CFOP_DESC[base] + (' (interestadual)' if c[0] in '26' else ' (exterior)')
    return ''

COMPRAS = {'101', '102', '111', '113', '116', '117', '118', '120', '121', '122', '124', '125', '126', '128',
           '551', '556', '401', '403', '406', '407', '651', '652', '653', '949'}
TRANSF = {'151', '152', '153', '154', '155', '156', '408', '409', '552', '557', '949'}

USO_CONSUMO = {'556', '407'}          # sem crédito nenhum (ICMS, IPI, PIS, COFINS)
COMBUSTIVEL_CONSUMO = {'653'}
ATIVO = {'551', '406'}
REVENDA_INDUSTRIA = {'101', '102', '111', '113', '116', '117', '118', '120', '121', '122', '124', '125', '126', '401', '403'}
COM_ST = {'401', '403', '406', '407', '408', '409', '410', '411'}
FRETE = {'351', '352', '353', '354', '355', '356', '360'}


def sufixos_entrada_aceitos(cfop_forn):
    """Sufixos (3 últimos dígitos) aceitos na entrada para um CFOP de saída do fornecedor."""
    s = str(cfop_forn)[1:]
    if s in {'101', '102', '103', '104', '105', '106', '107', '108', '109', '110', '111', '112', '113', '114', '115',
             '116', '117', '118', '119', '120', '121', '122', '123', '124', '125', '401', '402', '403', '404', '405',
             '551', '651', '652', '653', '654', '655', '656', '667', '929'}:
        return COMPRAS
    if s in {'151', '152', '153', '155', '156', '408', '409'}:
        return TRANSF
    if s == '922':
        return {'922'}
    return {s}


# sufixo de saída do fornecedor que não existe como entrada -> entrada mais provável
PREF_ENTRADA = {'405': '403', '402': '401', '103': '102', '104': '102', '105': '102', '106': '102', '109': '102', '110': '102',
                '114': '102', '115': '102', '107': '102', '108': '102', '404': '403', '929': '102', '654': '652', '119': '121', '123': '122', '655': '652', '656': '653', '667': '653'}


def cfop_entrada_esperado(cfop_forn, tem_st=False):
    """CFOP de entrada mais provável: 5->1, 6->2, 7->3; sufixo convertido quando não existe na entrada; ST -> x403."""
    c = str(cfop_forn or '')
    if len(c) != 4:
        return c
    pri = {'5': '1', '6': '2', '7': '3'}.get(c[0], c[0])
    suf = PREF_ENTRADA.get(c[1:], c[1:])
    if tem_st and suf == '102':
        suf = '403'
    if tem_st and suf == '101':
        suf = '401'
    return pri + suf


# CST de ICMS (2 últimos dígitos) aceitos na escrituração da entrada, por CST/CSOSN do fornecedor
CST_ENTRADA = {
    '00': {'00', '90'}, '02': {'02', '90'}, '10': {'10', '60', '70', '90'}, '15': {'15', '61', '90'},
    '20': {'20', '90'}, '30': {'30', '60', '70', '90'}, '40': {'40', '41', '90'}, '41': {'41', '40', '90'},
    '50': {'50', '40', '90'}, '51': {'51', '90'}, '53': {'53', '90'}, '60': {'60', '90'}, '61': {'61', '90'},
    '70': {'70', '60', '10', '90'}, '90': {'90', '00', '10', '20', '60', '70'},
    # CSOSN (fornecedor do Simples)
    '101': {'90', '00'}, '102': {'90', '40', '41'}, '103': {'90', '40', '41'}, '300': {'90', '40', '41'},
    '400': {'90', '40', '41'}, '201': {'10', '60', '70', '90'}, '202': {'10', '60', '70', '90'},
    '203': {'10', '60', '70', '90'}, '500': {'60', '90'}, '900': {'90', '00'},
}
CSOSN = {'101', '102', '103', '201', '202', '203', '300', '400', '500', '900'}
CST_COM_ST = {'10', '30', '60', '70', '201', '202', '203', '500'}

CST_DESC = {
    '00': 'tributada integralmente', '02': 'monofásico combustível', '10': 'tributada com ST', '15': 'monofásico com retenção',
    '20': 'base reduzida', '30': 'isenta com ST', '40': 'isenta', '41': 'não tributada', '50': 'suspensão', '51': 'diferimento',
    '53': 'monofásico diferido', '60': 'ST cobrada anteriormente', '61': 'monofásico retido', '70': 'base reduzida com ST', '90': 'outras',
    '101': 'Simples com crédito', '102': 'Simples sem crédito', '103': 'Simples isenção (faixa)', '201': 'Simples com crédito e ST',
    '202': 'Simples sem crédito e ST', '203': 'Simples isenção e ST', '300': 'Simples imune', '400': 'Simples não tributada',
    '500': 'Simples ST cobrada antes', '900': 'Simples outros',
}

# PIS/COFINS
PC_CREDITO = {'50', '51', '52', '53', '54', '55', '56', '60', '61', '62', '63', '64', '65', '66'}
PC_SEM_CREDITO = {'70', '71', '72', '73', '74', '75', '98', '99'}
PC_FORN_SEM_CREDITO = {'04', '05', '06', '07', '08', '09'}  # monofásico, ST, alíquota zero, isenta, sem incidência, suspensão
ALIQ_PIS_NC, ALIQ_COFINS_NC = 1.65, 7.6
