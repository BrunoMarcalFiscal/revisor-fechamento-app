# -*- coding: utf-8 -*-
"""Consolidação dos apontamentos de cada nota: um apontamento por causa.

1. Mesma regra em várias linhas da nota vira um apontamento só (linhas listadas, valores somados).
2. Regras que são a mesma causa se juntam (crédito de ICMS, IPI e PIS/COFINS em uso e consumo).
3. O que é consequência de outro erro (o total difere porque faltou uma linha; a base difere porque o
   ICMS está errado) entra como "também" dentro do erro de origem, não como erro novo.
Também define a categoria de cada regra, o fundamento (por que é erro) e o checklist de pontos conferidos."""
import re
from .util import brl

CATEGORIAS = [
    ('presenca', 'Lançamento'), ('cadastro', 'Identificação'), ('valores', 'Valores'), ('cfop', 'CFOP'), ('cst', 'CST'),
    ('icms', 'ICMS e crédito'), ('st', 'Substituição tributária'), ('ipi', 'IPI'), ('piscofins', 'PIS/COFINS'),
    ('frete', 'Frete'), ('iss', 'ISS'), ('retencoes', 'Retenções'), ('devolucao', 'Devolução'), ('cancel', 'Cancelamento e numeração'),
]
CAT_ROT = dict(CATEGORIAS)

CATEGORIA = {}
for cat, cods in {
    'presenca': ['NAO_LANCADA', 'NAO_LANCADA_MES_ANT', 'FORA_COMPETENCIA', 'SEM_XML', 'SEM_XML_PDF', 'PDF_NAO_LANCADO', 'DUPLICIDADE',
                 'OUTRA_EMPRESA', 'CASOU_SO_NUMERO', 'PARTE_NAO_LANCADA', 'LINHA_SOBRANDO'],
    'cadastro': ['PARTICIPANTE', 'PARTICIPANTE_FILIAL', 'CHAVE', 'SERIE', 'DATA_ANTERIOR', 'DATA_EMISSAO'],
    'valores': ['VALOR_TOTAL', 'LOCACAO_VALOR', 'VALOR_LINHA', 'SERVICO_LIQUIDO'],
    'cfop': ['CFOP_SAIDA', 'CFOP_UF', 'CFOP_UF_PRESENCIAL', 'CFOP_INCOMPATIVEL', 'CFOP_BRANCO', 'CFOP_FRETE', 'DESTINO_COMBUSTIVEL',
             'DESTINO_CONSUMO', 'DESTINO_ATIVO'],
    'cst': ['CST_ICMS', 'CST_SAIDA', 'CST_ORIGEM'],
    'icms': ['BASE_ICMS', 'ALIQ_ICMS', 'ICMS_SAIDA', 'ICMS_ENTRADA_PROPRIA', 'CREDITO_MAIOR', 'CREDITO_MENOR', 'CREDITO_USO_CONSUMO',
             'CREDITO_ATIVO', 'CREDITO_NAO_TOMADO', 'CREDITO_COMBUSTIVEL', 'CREDITO_DEV_VENDA', 'DEV_USO_CONSUMO_ICMS'],
    'st': ['CFOP_ST', 'CFOP_ST_SEM_ST', 'ST_VALOR', 'ST_MISTO'],
    'ipi': ['IPI_VALOR', 'IPI_USO_CONSUMO'],
    'piscofins': ['PIS_COFINS_VALOR', 'PC_CST_SAIDA', 'PC_MONOFASICO', 'PC_USO_CONSUMO', 'PC_NAO_TOMADO', 'PC_BASE_ICMS', 'PC_ALIQ', 'PC_FRETE'],
    'frete': ['FRETE_NAO_TOMADOR', 'FRETE_NAO_TOMADOR_CREDITO'],
    'iss': ['ISS_VALOR', 'ISS_ALIQ', 'ISS_BASE', 'ISS_RETIDO_NAO_LANCADO', 'ISS_RETIDO_INDEVIDO', 'SERVICO_CODIGO'],
    'retencoes': ['RETENCAO'],
    'devolucao': ['DEV_SEM_REF', 'DEV_ORIGINAL_FORA', 'DEV_VALOR_MAIOR', 'DEV_ITEM', 'DEV_QTD', 'DEV_PRECO', 'DEV_ICMS', 'DEV_IPI', 'DEV_ST',
                  'DEV_CFOP', 'DEVOLVIDA_INTEGRAL', 'DEV_NAO_LANCADA', 'DEV_REPETIDA', 'DEV_FORN_NAO_LANCADA'],
    'cancel': ['CANCELADA_LANCADA', 'CANCELADA_SO_NO_SISTEMA', 'CANCELADA_OK', 'XML_SEM_PROTOCOLO', 'DENEGADA', 'INUTILIZADA_LANCADA', 'INUTILIZADA_COM_NOTA', 'NUMERO_PULADO'],
}.items():
    for c in cods:
        CATEGORIA[c] = cat

# por que é erro (curto, para a equipe e para justificar ao cliente)
FUNDAMENTO = {
    'NAO_LANCADA': 'Todo documento fiscal do período tem de estar escriturado; nota fora do livro distorce apuração e cruzamentos do Fisco (SPED × XML).',
    'PARTE_NAO_LANCADA': 'O lançamento precisa cobrir todos os itens da nota; o que falta some da apuração.',
    'DUPLICIDADE': 'Nota lançada duas vezes duplica valor contábil e crédito/débito.',
    'SEM_XML': 'Lançamento sem documento não sustenta crédito nem custo; o XML é a prova da operação.',
    'PDF_NAO_LANCADO': 'O documento existe (está no PDF) e não foi escriturado.',
    'VALOR_TOTAL': 'O valor contábil tem de ser o total da nota (vNF).',
    'LOCACAO_VALOR': 'A mesma nota de locação não pode ter um valor no livro de serviços e outro no relatório de locação: um dos dois está errado no sistema.',
    'PARTICIPANTE': 'O participante do livro tem de ser o emitente/destinatário do documento (cruzamento por CNPJ).',
    'DATA_ANTERIOR': 'Não existe entrada antes da emissão do documento.',
    'CFOP_UF': 'O 1º dígito do CFOP de entrada segue a origem: 1 = mesmo estado, 2 = outro estado, 3 = exterior (Convênio s/nº de 1970, Anexo).',
    'CFOP_UF_PRESENCIAL': 'Na compra presencial em outro estado o fornecedor emite 5xxx, mas para a empresa a operação vem de outro estado (2xxx).',
    'CFOP_INCOMPATIVEL': 'O CFOP de entrada precisa corresponder à natureza da operação do fornecedor.',
    'CFOP_SAIDA': 'Na saída o CFOP escriturado é o do documento emitido.',
    'CST_ICMS': 'O CST da entrada precisa ser compatível com a tributação da nota (CST/CSOSN do emitente).',
    'CST_SAIDA': 'Na saída o CST escriturado é o do documento emitido.',
    'CREDITO_MAIOR': 'O crédito de ICMS é limitado ao imposto destacado (LC 87/96, art. 23); no Simples, ao permitido em pCredSN (LC 123/06, art. 23).',
    'CREDITO_USO_CONSUMO': 'Uso e consumo (x556/x407) não dá crédito de ICMS (LC 87/96, art. 33, I), de IPI nem de PIS/COFINS.',
    'CREDITO_ATIVO': 'Ativo imobilizado: crédito de ICMS em 1/48 por mês pelo CIAP (LC 87/96, art. 20, §5º).',
    'CREDITO_NAO_TOMADO': 'Compra para revenda/industrialização com ICMS destacado dá crédito; sem crédito, a empresa paga a mais.',
    'CREDITO_DEV_VENDA': 'Na devolução de venda a empresa recupera o ICMS debitado na venda (crédito do destacado na devolução).',
    'CFOP_ST': 'Mercadoria com ST entra em CFOP de ST (x401, x403, x406, x407); o CFOP define a apuração e o ressarcimento.',
    'CFOP_ST_SEM_ST': 'Mercadoria sem ST não pode entrar em CFOP de ST.',
    'ST_MISTO': 'Itens com ST e sem ST da mesma nota vão em linhas separadas, cada uma no seu CFOP.',
    'PC_MONOFASICO': 'Revenda de produto monofásico/alíquota zero não gera crédito de PIS/COFINS (Lei 10.833/03, art. 3º, I, b).',
    'PC_BASE_ICMS': 'Desde 05/2023 o ICMS destacado sai da base do crédito de PIS/COFINS (Lei 14.592/2023).',
    'PC_NAO_TOMADO': 'No Lucro Real, compra tributada para revenda/insumo dá crédito de PIS/COFINS (Leis 10.637/02 e 10.833/03).',
    'ISS_RETIDO_NAO_LANCADO': 'ISS retido na nota: o tomador recolhe; sem marcar a retenção, o ISS é apurado de novo ou não é recolhido.',
    'RETENCAO': 'Retenção destacada na nota precisa ser escriturada para compensação (prestado) ou recolhimento (tomado).',
    'DEV_PRECO': 'A devolução sai pelo mesmo valor unitário e tributação da operação original.',
    'DEV_QTD': 'Não se devolve mais do que foi vendido/comprado.',
    'DEV_ICMS': 'O ICMS da devolução é proporcional ao destacado na nota de origem.',
    'CANCELADA_LANCADA': 'Nota cancelada não gera crédito nem débito; no livro entra sem valor (situação cancelada).',
    'INUTILIZADA_LANCADA': 'Número inutilizado não tem operação; não pode estar no livro com valor.',
    'NUMERO_PULADO': 'Todo número da série tem de ter nota autorizada, cancelada ou inutilizada.',
}

ICMS_RAIZ = {'ICMS_SAIDA', 'ICMS_ENTRADA_PROPRIA', 'CREDITO_MAIOR', 'CREDITO_MENOR', 'CREDITO_DEV_VENDA', 'CREDITO_USO_CONSUMO', 'CREDITO_NAO_TOMADO'}
# (raízes, consequência, escopo): a consequência entra dentro da raiz
ABSORVE = [
    ({'ST_MISTO'}, {'CFOP_ST', 'CFOP_ST_SEM_ST', 'PARTE_NAO_LANCADA', 'LINHA_SOBRANDO', 'VALOR_LINHA', 'ST_VALOR'}, 'nota'),
    ({'PARTE_NAO_LANCADA'}, {'VALOR_TOTAL'}, 'nota'),
    ({'CFOP_ST', 'CFOP_ST_SEM_ST'}, {'ST_VALOR'}, 'par'),
    (ICMS_RAIZ, {'BASE_ICMS', 'ALIQ_ICMS'}, 'par'),
    ({'DEV_PRECO'}, {'DEV_ICMS', 'DEV_IPI', 'DEV_ST', 'DEV_VALOR_MAIOR'}, 'nota'),
    ({'DEV_QTD'}, {'DEV_VALOR_MAIOR'}, 'nota'),
    ({'VALOR_TOTAL'}, {'ISS_BASE', 'ISS_VALOR', 'SERVICO_LIQUIDO', 'VALOR_LINHA'}, 'nota'),
    ({'FRETE_NAO_TOMADOR_CREDITO'}, {'CREDITO_MAIOR', 'CREDITO_NAO_TOMADO', 'PC_FRETE'}, 'nota'),
    ({'CANCELADA_SO_NO_SISTEMA'}, {'DEV_REPETIDA'}, 'nota'),
]
# regras que são a mesma causa: juntam num apontamento só
JUNTA = [
    ({'CREDITO_USO_CONSUMO', 'IPI_USO_CONSUMO', 'PC_USO_CONSUMO'}, 'CREDITO_USO_CONSUMO', 'Crédito em uso e consumo (deve ser zero)'),
    ({'PIS_COFINS_VALOR', 'PC_CST_SAIDA'}, 'PIS_COFINS_VALOR', 'PIS/COFINS diferente do XML'),
    ({'ISS_RETIDO_NAO_LANCADO', 'ISS_RETIDO_INDEVIDO'}, None, None),
]
ORDEM_SEV = {'erro': 0, 'alerta': 1, 'info': 2}


def _pares(a):
    if a.get('pares') is not None:
        return set(a['pares'])
    return set() if a.get('par') is None else {a['par']}


def _uniao_txt(*xs):
    out = []
    for x in xs:
        for p in [t.strip() for t in (x or '').split(' | ')]:
            if p and p not in out:
                out.append(p)
    return ' | '.join(out)


RE_RS = re.compile(r'R\$ ?-?[\d.]+,\d{2}')


def _rs_val(t):
    return float(t.replace('R$', '').replace(' ', '').replace('.', '').replace(',', '.'))


def _junta(a, b):
    """Junta dois textos. Se só os valores em R$ mudam, devolve uma frase com os valores somados."""
    a, b = (a or '').strip(), (b or '').strip()
    if not a or a == b:
        return a or b
    if not b:
        return a
    if RE_RS.sub('#', a) == RE_RS.sub('#', b):
        va, vb = RE_RS.findall(a), RE_RS.findall(b)
        if va and len(va) == len(vb):
            it = iter([brl(_rs_val(x) + _rs_val(y)) for x, y in zip(va, vb)])
            return RE_RS.sub(lambda m: next(it), a)
    return _uniao_txt(a, b)


def _funde(base, outro, titulo=None, cod=None):
    base['pares'] = sorted(_pares(base) | _pares(outro))
    base['campos'] = sorted(set(base.get('campos') or [base.get('campo')]) | set(outro.get('campos') or [outro.get('campo')]) - {'', None})
    base['valor'] = round((base.get('valor') or 0) + (outro.get('valor') or 0), 2)
    base['detalhe'] = _junta(base['detalhe'], outro['detalhe'])
    base['xml'] = _junta(base['xml'], outro['xml'])
    base['sis'] = _junta(base['sis'], outro['sis'])
    base['correcao'] = _junta(base['correcao'], outro['correcao'])
    if ORDEM_SEV.get(outro['sev'], 3) < ORDEM_SEV.get(base['sev'], 3):
        base['sev'] = outro['sev']
    if titulo:
        base['titulo'] = titulo
    if cod:
        base['cod'] = cod
    base['linhas_n'] = base.get('linhas_n', 1) + outro.get('linhas_n', 1)


def consolidar(n):
    aps = [dict(a) for a in n.aponts]
    for a in aps:
        a['pares'] = sorted(_pares(a))
        a['campos'] = [a['campo']] if a.get('campo') else []
        a.setdefault('tambem', [])
    # 1 e 2: mesma regra (ou mesma causa) -> um só
    grupos = {}
    for c_set, alvo, _ in JUNTA:
        if alvo:
            for c in c_set:
                grupos[c] = alvo
    out = []
    por_cod = {}
    for a in aps:
        chave = grupos.get(a['cod'], a['cod'])
        if chave in por_cod and a['sev'] != 'info':
            alvo = por_cod[chave]
            tit = next((t for cs, al, t in JUNTA if al == chave and a['cod'] in cs and alvo['cod'] != a['cod']), None)
            _funde(alvo, a, tit, chave)
        else:
            if chave != a['cod']:
                tit = next((t for cs, al, t in JUNTA if al == chave), None)
                a['cod'], a['titulo'] = chave, tit or a['titulo']
            por_cod[chave] = a
            out.append(a)
    # 3: consequência entra na raiz
    for raizes, filhos, escopo in ABSORVE:
        rs = [a for a in out if a['cod'] in raizes]
        if not rs:
            continue
        sobra = []
        for a in out:
            if a['cod'] in filhos and a not in rs:
                alvo = None
                for r in rs:
                    if escopo == 'nota' or (set(a['pares']) & set(r['pares'])):
                        alvo = r
                        break
                if alvo is not None:
                    alvo['tambem'].append('%s: %s' % (a['titulo'], a['detalhe']) if a['detalhe'] else a['titulo'])
                    alvo['pares'] = sorted(set(alvo['pares']) | set(a['pares']))
                    alvo['campos'] = sorted(set(alvo['campos']) | set(a['campos']))
                    if a['correcao'] and a['correcao'] not in alvo['correcao'] and escopo == 'par':
                        alvo['correcao'] = _uniao_txt(alvo['correcao'], a['correcao'])
                    continue
            sobra.append(a)
        out = sobra
    for a in out:
        a['categoria'] = CATEGORIA.get(a['cod'], 'presenca' if a['sev'] == 'info' else 'valores')
        a['fundamento'] = FUNDAMENTO.get(a['cod'], '')
        a['par'] = a['pares'][0] if a['pares'] else None
        if a.get('linhas_n', 1) > 1 and 'linhas da nota' not in a['detalhe']:
            a['detalhe'] = '%s (somando %d linhas da nota)' % (a['detalhe'], a['linhas_n'])
    n.aponts = sorted(out, key=lambda a: (ORDEM_SEV.get(a['sev'], 3), [c for c, _ in CATEGORIAS].index(a['categoria']) if a['categoria'] in CAT_ROT else 99))


# pontos conferidos por tipo de documento: (categoria, rótulo, colunas da planilha que permitem conferir)
PONTOS = {
    'nfe': [('presenca', 'Lançamento', None), ('valores', 'Valor', ['valor_contabil']), ('cadastro', 'Participante e datas', ['cnpj', 'chave', 'data_lanc', 'data_emissao']),
            ('cfop', 'CFOP e destino', ['cfop']), ('cst', 'CST', ['cst_icms']), ('icms', 'ICMS e crédito', ['valor_icms']), ('st', 'ST', ['cfop']),
            ('ipi', 'IPI', ['valor_ipi']), ('piscofins', 'PIS/COFINS', ['cst_pis', 'valor_pis'])],
    'cte': [('presenca', 'Lançamento', None), ('valores', 'Valor', ['valor_contabil']), ('frete', 'Tomador', None), ('cfop', 'CFOP', ['cfop']),
            ('icms', 'ICMS e crédito', ['valor_icms']), ('piscofins', 'PIS/COFINS', ['cst_pis'])],
    'nfse': [('presenca', 'Lançamento', None), ('valores', 'Valor', ['valor_contabil']), ('cadastro', 'Participante', ['cnpj']),
             ('iss', 'ISS', ['valor_iss', 'aliq_iss', 'iss_retido']), ('retencoes', 'Retenções', ['irrf', 'csll_ret', 'inss_ret', 'pis_ret', 'cofins_ret'])],
}


CAMPO_CAT = {'ipi': 'ipi', 'pis': 'piscofins', 'cofins': 'piscofins', 'cst_pis': 'piscofins', 'pis_bc': 'piscofins', 'icms': 'icms', 'bc': 'icms',
             'aliq': 'icms', 'st': 'st', 'cfop': 'cfop', 'cst': 'cst', 'iss': 'iss', 'aliq_iss': 'iss', 'iss_retido': 'iss', 'base_iss': 'iss',
             'irrf': 'retencoes', 'csll': 'retencoes', 'inss': 'retencoes', 'vc': 'valores'}


def checklist(n, cols, tipo, lancada, natureza):
    pts = list(PONTOS.get(tipo, []))
    if natureza and natureza.startswith('dev'):
        pts.append(('devolucao', 'Devolução × origem', None))
    out = []
    def toca(a, cat):  # o apontamento atinge o ponto pela categoria ou pelos campos envolvidos
        return a.get('categoria') == cat or cat in {CAMPO_CAT.get(c) for c in (a.get('campos') or [])}
    for cat, rot, req in pts:
        afetam = [a for a in n.aponts if a['sev'] in ('erro', 'alerta') and toca(a, cat)]
        cods = [a['k'] for a in afetam]
        if cods:
            st = 'erro' if any(a['sev'] == 'erro' for a in afetam) else 'alerta'
            obs = ''
        elif cat != 'presenca' and not lancada:
            st, obs = 'nao', 'sem lançamento'
        elif req and not (set(req) & set(cols or [])):
            st, obs = 'nao', 'a planilha não tem a coluna'
        else:
            st, obs = 'ok', ''
        out.append({'cat': cat, 'rot': rot, 'st': st, 'obs': obs, 'ks': cods})
    return out
