# -*- coding: utf-8 -*-
"""Cruza os documentos (XML e PDF) com a planilha do sistema e gera os apontamentos por nota.

Fluxo: carregar documentos por card -> identificar a empresa -> ler a planilha -> agrupar lançamentos por nota
-> casar lançamento x documento (chave, número+CNPJ, número+valor, número) -> comparar linha a linha
(CFOP/CST) -> presença (não lançada, sem XML, só no PDF, duplicada, cancelada)."""
import os, re, datetime
from collections import Counter, defaultdict
from . import xml_docs, planilha, pdf_docs, destino, devolucoes, consolidar, apuracao, painel
from .util import num, digitos, numero_doc, raiz, igual, data_iso, brl, TOL, norm
from .regras_base import (sufixos_entrada_aceitos, cfop_entrada_esperado, desc_cfop, CST_ENTRADA, CSOSN, CST_COM_ST, CST_DESC, USO_CONSUMO, ATIVO,
                          REVENDA_INDUSTRIA, COM_ST, FRETE, COMBUSTIVEL_CONSUMO, COMPRAS, PC_CREDITO, PC_SEM_CREDITO, PC_FORN_SEM_CREDITO, ALIQ_PIS_NC, ALIQ_COFINS_NC)

CARDS = ['entrada', 'saida', 'frete', 'prestado', 'tomado', 'diversos', 'planilha', 'serv_entrada', 'serv_saida', 'planilha_servicos', 'planilha_locacoes']
# pasta da planilha -> prefixo da chave dos ajustes de colunas
PLANILHAS = [('planilha', ''), ('serv_entrada', 'serv_entrada/'), ('serv_saida', 'serv_saida/'), ('planilha_servicos', 'servicos/'), ('planilha_locacoes', 'locacoes/')]
SERV_CARDS = {'serv_entrada': 'tomado', 'serv_saida': 'prestado', 'planilha_servicos': None}
MOV_NOME = {'entrada': 'Entradas', 'saida': 'Saídas', 'frete': 'Frete de entrada', 'prestado': 'Serviços prestados',
            'tomado': 'Serviços tomados', 'diversos': 'Documentos diversos', 'indef': 'Sem movimento definido'}

# catálogo de apontamentos: código -> (severidade, título curto)
CATALOGO = {
    'NAO_LANCADA': ('erro', 'Nota não lançada no sistema'),
    'NAO_LANCADA_MES_ANT': ('alerta', 'Nota de mês anterior não encontrada no lançamento'),
    'FORA_COMPETENCIA': ('alerta', 'XML de outra competência'),
    'SEM_XML': ('alerta', 'Lançamento sem XML'),
    'SEM_XML_PDF': ('alerta', 'Lançamento sem XML (consta no PDF)'),
    'PDF_NAO_LANCADO': ('erro', 'Documento do PDF não lançado'),
    'DUPLICIDADE': ('erro', 'Nota lançada mais de uma vez'),
    'CANCELADA_LANCADA': ('erro', 'Nota cancelada lançada com valor'),
    'OUTRA_EMPRESA': ('alerta', 'XML não é desta empresa'),
    'VALOR_TOTAL': ('erro', 'Valor total diferente do XML'),
    'VALOR_LINHA': ('erro', 'Valor da linha diferente do XML'),
    'PARTE_NAO_LANCADA': ('erro', 'Parte da nota sem lançamento'),
    'LINHA_SOBRANDO': ('alerta', 'Linha lançada sem correspondência na nota'),
    'PARTICIPANTE': ('erro', 'CNPJ do participante diferente'),
    'PARTICIPANTE_FILIAL': ('alerta', 'Participante em outra filial (mesma raiz)'),
    'CHAVE': ('erro', 'Chave de acesso diferente do XML'),
    'SERIE': ('alerta', 'Série diferente do XML'),
    'CASOU_SO_NUMERO': ('alerta', 'Casou só pelo número da nota'),
    'MANUAL': ('info', 'Nota lançada à mão'),
    'NFSE_REPETIDA': ('alerta', 'NFS-e repetida: provável cancelada e reemitida'),
    'CANCEL_MANUAL': ('info', 'Marcada como cancelada por você'),
    'LOCACAO_VALOR': ('erro', 'Valor diferente do relatório de locação'),
    'DATA_ANTERIOR': ('erro', 'Entrada lançada antes da emissão'),
    'DATA_EMISSAO': ('alerta', 'Data de emissão diferente do XML'),
    'CFOP_SAIDA': ('erro', 'CFOP diferente do XML'),
    'CFOP_UF': ('erro', 'CFOP estadual/interestadual trocado'),
    'CFOP_UF_PRESENCIAL': ('alerta', 'Compra presencial em outro estado com CFOP estadual'),
    'DESTINO_COMBUSTIVEL': ('alerta', 'Combustível lançado fora do CFOP de combustível'),
    'DESTINO_CONSUMO': ('alerta', 'Item de uso e consumo lançado como revenda/industrialização'),
    'DESTINO_ATIVO': ('alerta', 'Bem de ativo imobilizado lançado em outro CFOP'),
    'IPI_USO_CONSUMO': ('erro', 'Crédito de IPI em uso e consumo'),
    'CREDITO_COMBUSTIVEL': ('alerta', 'Crédito de ICMS em combustível para consumo'),
    'CFOP_INCOMPATIVEL': ('erro', 'CFOP de entrada incompatível com a nota'),
    'CFOP_BRANCO': ('erro', 'CFOP em branco'),
    'CFOP_ST': ('erro', 'Mercadoria com ST lançada em CFOP sem ST'),
    'CFOP_ST_SEM_ST': ('erro', 'Mercadoria sem ST lançada em CFOP de ST'),
    'CFOP_FRETE': ('erro', 'CFOP de frete incorreto'),
    'CST_ICMS': ('erro', 'CST de ICMS incompatível com a nota'),
    'CST_SAIDA': ('erro', 'CST de ICMS diferente do XML'),
    'CST_ORIGEM': ('alerta', 'Origem da mercadoria diferente do XML'),
    'BASE_ICMS': ('erro', 'Base do ICMS diferente do XML'),
    'ALIQ_ICMS': ('erro', 'Alíquota do ICMS diferente do XML'),
    'ICMS_SAIDA': ('erro', 'Valor do ICMS diferente do XML'),
    'CREDITO_MAIOR': ('erro', 'Crédito de ICMS maior que o destacado'),
    'CREDITO_MENOR': ('alerta', 'Crédito de ICMS menor que o destacado'),
    'CREDITO_USO_CONSUMO': ('erro', 'Crédito de ICMS em uso e consumo'),
    'CREDITO_ATIVO': ('alerta', 'Crédito de ativo lançado integral'),
    'CREDITO_NAO_TOMADO': ('alerta', 'ICMS destacado sem crédito'),
    'ST_VALOR': ('alerta', 'ICMS-ST diferente do XML'),
    'IPI_VALOR': ('erro', 'IPI diferente do XML'),
    'PIS_COFINS_VALOR': ('erro', 'PIS/COFINS diferente do XML'),
    'PC_CST_SAIDA': ('erro', 'CST de PIS/COFINS diferente do XML'),
    'PC_MONOFASICO': ('erro', 'Crédito de PIS/COFINS em produto monofásico/alíquota zero'),
    'PC_USO_CONSUMO': ('erro', 'Crédito de PIS/COFINS em uso e consumo'),
    'PC_NAO_TOMADO': ('alerta', 'Compra tributada sem crédito de PIS/COFINS'),
    'PC_BASE_ICMS': ('alerta', 'Base do crédito PIS/COFINS sem excluir o ICMS'),
    'PC_ALIQ': ('alerta', 'Alíquota de PIS/COFINS fora do não cumulativo'),
    'PC_FRETE': ('alerta', 'Frete sem crédito de PIS/COFINS'),
    'FRETE_NAO_TOMADOR': ('alerta', 'Empresa não é a tomadora do frete'),
    'FRETE_NAO_TOMADOR_CREDITO': ('erro', 'Crédito de frete sem ser a tomadora'),
    'ISS_VALOR': ('erro', 'ISS diferente da nota'),
    'ISS_ALIQ': ('erro', 'Alíquota do ISS diferente da nota'),
    'ISS_RETIDO_NAO_LANCADO': ('erro', 'ISS retido na nota não lançado'),
    'ISS_RETIDO_INDEVIDO': ('erro', 'ISS lançado como retido sem retenção na nota'),
    'RETENCAO': ('erro', 'Retenção federal diferente da nota'),
    'SERVICO_CODIGO': ('alerta', 'Código do serviço (LC 116) diferente da nota'),
    'ISS_BASE': ('erro', 'Base do ISS diferente da nota'),
    'SERVICO_LIQUIDO': ('alerta', 'Valor líquido diferente da nota'),
    'CREDITO_DEV_VENDA': ('erro', 'Devolução de venda sem o crédito do ICMS'),
    'ST_MISTO': ('erro', 'Itens com e sem ST lançados numa linha só'),
    'ICMS_ENTRADA_PROPRIA': ('erro', 'ICMS da entrada própria diferente do XML'),
    'DEV_USO_CONSUMO_ICMS': ('alerta', 'Devolução de uso e consumo sem débito de ICMS'),
    'DEV_REPETIDA': ('alerta', 'Duas devoluções iguais'),
    'DEV_FORN_NAO_LANCADA': ('alerta', 'Devolução feita por nota de entrada do fornecedor'),
    'CANCELADA_SO_NO_SISTEMA': ('alerta', 'Cancelada no sistema sem XML de cancelamento'),
    'CANCELADA_OK': ('info', 'Cancelada no XML e no sistema'),
    'XML_SEM_PROTOCOLO': ('alerta', 'XML sem protocolo de autorização'),
    'DENEGADA': ('erro', 'Nota denegada'),
    'DEV_SEM_REF': ('alerta', 'Devolução sem nota de origem referenciada'),
    'DEV_ORIGINAL_FORA': ('info', 'Nota de origem da devolução fora do lote'),
    'DEV_VALOR_MAIOR': ('erro', 'Devolução maior que a nota de origem'),
    'DEV_ITEM': ('erro', 'Item devolvido não está na nota de origem'),
    'DEV_QTD': ('erro', 'Quantidade devolvida maior que a da nota de origem'),
    'DEV_PRECO': ('erro', 'Preço da devolução diferente da nota de origem'),
    'DEV_ICMS': ('erro', 'ICMS da devolução não bate com a nota de origem'),
    'DEV_IPI': ('erro', 'IPI da devolução não bate com a nota de origem'),
    'DEV_ST': ('erro', 'ICMS-ST da devolução não bate com a nota de origem'),
    'DEV_CFOP': ('alerta', 'CFOP da devolução não corresponde à operação de origem'),
    'DEVOLVIDA_INTEGRAL': ('alerta', 'Compra devolvida por inteiro'),
    'DEV_NAO_LANCADA': ('erro', 'Devolução não lançada (a compra foi)'),
    'INUTILIZADA_LANCADA': ('erro', 'Número inutilizado lançado'),
    'INUTILIZADA_COM_NOTA': ('erro', 'Número inutilizado com nota autorizada'),
    'NUMERO_PULADO': ('alerta', 'Numeração pulada sem cancelamento/inutilização'),
}


class Nota:
    """Uma linha do painel: documento (XML/PDF) e/ou lançamento da planilha."""
    def __init__(self, mov, doc=None, lancs=None):
        self.mov = mov
        self.doc = doc
        self.lancs = lancs or []
        self.aponts = []
        self.pares = []

    def ap(self, cod, detalhe='', xml='', sis='', correcao='', par=None, campo='', valor=0.0):
        sev, titulo = CATALOGO[cod]
        self.aponts.append({'cod': cod, 'sev': sev, 'titulo': titulo, 'detalhe': detalhe, 'xml': _s(xml), 'sis': _s(sis),
                            'correcao': correcao, 'par': par, 'campo': campo, 'valor': round(valor or 0, 2)})


def _s(v):
    if isinstance(v, float):
        return brl(v)
    return '' if v is None else str(v)


def _cst(v, csosn_ctx=False):
    """Normaliza CST de ICMS -> (origem, tributação)."""
    d = digitos(v)
    if not d:
        return ('', '')
    if len(d) >= 4:
        d = d[-3:]
    if len(d) == 3:
        if csosn_ctx and d in CSOSN:
            return ('', d)
        return (d[0], d[1:])
    if len(d) == 1:
        return ('', '0' + d)
    return ('', d)


def _pc(v):
    d = digitos(v)
    return d[-2:].zfill(2) if d else ''


def _mes(d):
    return (d or '')[:7]


# --------------------------------------------------------------------------------------------- carga

def _arquivos(pasta, exts):
    if not os.path.isdir(pasta):
        return []
    out = []
    for raiz_, _, fs in os.walk(pasta):
        for f in sorted(fs):
            if f.startswith('.') or f.startswith('~$'):
                continue
            if os.path.splitext(f)[1].lower() in exts:
                out.append(os.path.join(raiz_, f))
    return out


def carregar(sessao):
    docs, eventos, erros = [], [], []
    inventario = {}
    for card in ['entrada', 'saida', 'frete', 'prestado', 'tomado']:
        arqs = _arquivos(os.path.join(sessao, card), {'.xml'})
        n = Counter()
        for a in arqs:
            for d in xml_docs.ler(a):
                if d['tipo'] in ('evento', 'inut'):
                    eventos.append(d); n['eventos'] += 1
                elif d['tipo'] in ('nfe', 'cte', 'nfse'):
                    d['card'] = card
                    docs.append(d); n[d['tipo']] += 1
                else:
                    erros.append({'card': card, 'arquivo': d.get('arquivo'), 'erro': d.get('erro') or ('XML não reconhecido: ' + d.get('raiz', ''))})
                    n['nao_lidos'] += 1
        inventario[card] = {'arquivos': len(arqs), 'lidos': dict(n)}
    pdfs = []
    for a in _arquivos(os.path.join(sessao, 'diversos'), {'.pdf'}):
        pdfs.append(pdf_docs.ler(a))
    # XML soltos no card de diversos também valem
    for a in _arquivos(os.path.join(sessao, 'diversos'), {'.xml'}):
        for d in xml_docs.ler(a):
            if d['tipo'] in ('nfe', 'cte', 'nfse'):
                d['card'] = 'diversos'; docs.append(d)
            elif d['tipo'] in ('evento', 'inut'):
                eventos.append(d)
    inventario['diversos'] = {'arquivos': len(pdfs), 'paginas': sum(p['paginas'] for p in pdfs),
                              'documentos': sum(len(p['docs']) for p in pdfs),
                              'sem_texto': sum(len(p['sem_texto']) for p in pdfs)}
    return docs, eventos, pdfs, erros, inventario


def identificar_empresa(docs, cnpj_inf=''):
    c = Counter()
    nomes, ufs = {}, {}
    for d in docs:
        if d['tipo'] == 'nfe':
            if d['card'] == 'saida' or d.get('tpNF') == '1' and d['card'] != 'entrada':
                c[d['emit']['cnpj']] += 2; nomes[d['emit']['cnpj']] = d['emit']['nome']; ufs[d['emit']['cnpj']] = d['emit']['uf']
            elif d['card'] == 'entrada' and d['dest']['cnpj']:
                c[d['dest']['cnpj']] += 1; nomes.setdefault(d['dest']['cnpj'], d['dest']['nome']); ufs.setdefault(d['dest']['cnpj'], d['dest']['uf'])
        elif d['tipo'] == 'cte' and d.get('tomador_cnpj'):
            c[d['tomador_cnpj']] += 1
        elif d['tipo'] == 'nfse':
            if d['card'] == 'prestado':
                c[d['prestador']['cnpj']] += 2; nomes.setdefault(d['prestador']['cnpj'], d['prestador']['nome'])
            elif d['card'] == 'tomado':
                c[d['tomador']['cnpj']] += 1; nomes.setdefault(d['tomador']['cnpj'], d['tomador']['nome'])
    cnpj = digitos(cnpj_inf)
    if not cnpj and c:
        cnpj = c.most_common(1)[0][0]
    # nome/UF: procura qualquer doc com esse CNPJ
    nome, uf = nomes.get(cnpj, ''), ufs.get(cnpj, '')
    if not nome or not uf:
        for d in docs:
            if d['tipo'] == 'nfe':
                for lado in ('emit', 'dest'):
                    if raiz(d[lado]['cnpj']) == raiz(cnpj):
                        nome = nome or d[lado]['nome']; uf = uf or d[lado]['uf']
    return {'cnpj': cnpj, 'raiz': raiz(cnpj), 'nome': nome, 'uf': uf}


# --------------------------------------------------------------------------------------------- documento unificado

def _unificar(d, emp, cancelados):
    """Transforma o XML lido no registro comum usado no casamento."""
    r = emp['raiz']
    u = {'tipo': d['tipo'], 'arquivo': d['arquivo'], 'card': d['card'], 'raw': d, 'avisos': []}
    if d['tipo'] == 'nfe':
        propria = raiz(d['emit']['cnpj']) == r
        if propria:
            # nota própria: quem manda é o tpNF (0 = entrada própria, 1 = saída), não o quadro onde veio
            mov = 'saida' if d['tpNF'] == '1' else 'entrada'
            if d['card'] == 'entrada' and mov == 'saida':
                u['avisos'].append('XML de saída estava no quadro de Entradas; tratado como saída.')
            if d['card'] == 'saida' and mov == 'entrada':
                u['avisos'].append('Nota de entrada própria (emitida pela empresa) estava no quadro de Saídas; tratada como entrada.')
        else:
            mov = 'entrada'
            if d['card'] == 'saida':
                u['avisos'].append('XML de terceiro estava no quadro de Saídas; tratado como entrada.')
        u['natureza'] = devolucoes.natureza(d, propria, r)
        part = d['dest'] if (mov == 'saida' or propria) else d['emit']
        u.update({'mov': mov, 'chave': d['chave'], 'numero': numero_doc(d['numero']), 'serie': d['serie'], 'modelo': d['modelo'],
                  'emissao': d['emissao'], 'part_cnpj': part['cnpj'], 'part_nome': part['nome'], 'part_uf': part['uf'],
                  'valor': d['totais']['vNF'], 'propria': propria,
                  'empresa_envolvida': propria or raiz(d['dest']['cnpj']) == r,
                  'cancelada': d['chave'] in cancelados or d.get('protocolo') in ('101', '151', '155'),
                  'cancel_fonte': 'evento' if d['chave'] in cancelados else ('protocolo' if d.get('protocolo') in ('101', '151', '155') else ''),
                  'denegada': d.get('protocolo') in ('110', '205', '301', '302', '303'),
                  'sem_protocolo': not d.get('protocolo')})
    elif d['tipo'] == 'cte':
        u.update({'mov': 'frete', 'chave': d['chave'], 'numero': numero_doc(d['numero']), 'serie': d['serie'], 'modelo': d['modelo'],
                  'emissao': d['emissao'], 'part_cnpj': d['emit']['cnpj'], 'part_nome': d['emit']['nome'], 'part_uf': d['emit']['uf'],
                  'valor': d['vTPrest'], 'propria': False,
                  'empresa_envolvida': r in (raiz(d['tomador_cnpj']), raiz(d['rem']['cnpj']), raiz(d['dest']['cnpj'])),
                  'cancelada': d['chave'] in cancelados or d.get('protocolo') in ('101', '151'),
                  'cancel_fonte': 'evento' if d['chave'] in cancelados else ('protocolo' if d.get('protocolo') in ('101', '151') else ''),
                  'denegada': d.get('protocolo') in ('110', '205', '301', '302', '303'), 'sem_protocolo': not d.get('protocolo')})
    else:
        prest_emp = raiz(d['prestador']['cnpj']) == r
        if d['card'] == 'prestado':
            mov = 'prestado'
        elif d['card'] == 'tomado':
            mov = 'tomado'
        else:
            mov = 'prestado' if prest_emp else 'tomado'
        if d['card'] == 'tomado' and prest_emp:
            mov = 'prestado'; u['avisos'].append('NFS-e emitida pela empresa estava no card de Tomados; tratada como prestado.')
        if d['card'] == 'prestado' and not prest_emp and raiz(d['tomador']['cnpj']) == r:
            mov = 'tomado'; u['avisos'].append('NFS-e de terceiro estava no card de Prestados; tratada como tomado.')
        part = d['tomador'] if mov == 'prestado' else d['prestador']
        u.update({'mov': mov, 'chave': d.get('chave', ''), 'numero': numero_doc(d['numero']), 'serie': '', 'modelo': 'NFS-e',
                  'emissao': d['emissao'], 'part_cnpj': part['cnpj'], 'part_nome': part['nome'], 'part_uf': part.get('uf', ''),
                  'valor': d['vServ'], 'propria': mov == 'prestado',
                  'empresa_envolvida': r in (raiz(d['prestador']['cnpj']), raiz(d['tomador']['cnpj'])),
                  'cancelada': d.get('cancelada') or (d.get('chave') in cancelados),
                  'cancel_fonte': 'evento' if d.get('chave') in cancelados else ('xml' if d.get('cancelada') else ''), 'denegada': False, 'sem_protocolo': False})
    return u


# --------------------------------------------------------------------------------------------- planilha

def _mov_hint(reg):
    m = (reg.get('movimento') or '').strip().lower()
    aba = (reg.get('aba', '') + ' ' + reg.get('arquivo', '')).lower()
    if reg.get('mov_forcado'):
        return reg['mov_forcado']
    if reg.get('origem') == 'servicos':
        # relatório de serviços: prestado = saída (S), tomado = entrada (E)
        if 'prest' in m or m.startswith('s') or m in ('1', 'p'):
            return 'prestado'
        if 'toma' in m or m.startswith('e') or m in ('0', 't'):
            return 'tomado'
        if 'prestad' in aba or 'emitid' in aba:
            return 'prestado'
        if 'tomad' in aba or 'recebid' in aba:
            return 'tomado'
        return 'servico'
    mod = (reg.get('modelo') or '').upper()
    cf = reg.get('cfop') or ''
    if mod in ('57', '67', 'CTE', 'CT-E', 'CTRC') or cf[1:] in FRETE:
        return 'frete'
    if any(k in aba for k in ('frete', 'cte', 'ct-e', 'conhecimento')):
        return 'frete'
    if 'prestad' in aba:
        return 'prestado'
    if 'tomad' in aba:
        return 'tomado'
    if m.startswith('e') or m in ('0',):
        return 'entrada'
    if m.startswith('s') or m in ('1',):
        return 'saida'
    if cf[:1] in ('1', '2', '3'):
        return 'entrada'
    if cf[:1] in ('5', '6', '7'):
        return 'saida'
    if 'entrada' in aba:
        return 'entrada'
    if 'saida' in aba or 'saída' in aba:
        return 'saida'
    if 'servi' in aba or 'nfs' in aba or mod.startswith('NFS') or reg.get('valor_iss'):
        return 'servico'
    return ''


def ler_planilhas(sessao, ajustes):
    descs, regs = [], []
    for card, pref in PLANILHAS:
        for a in _arquivos(os.path.join(sessao, card), {'.xlsx', '.xlsm', '.xls', '.csv', '.txt', '.pdf'} if card in SERV_CARDS else {'.xlsx', '.xlsm', '.xls', '.csv', '.txt'}):
            try:
                d = planilha.analisar(a, (ajustes or {}).get(pref + os.path.basename(a)), servicos=(card in SERV_CARDS), locacoes=(card == 'planilha_locacoes'))
            except Exception as e:
                descs.append({'arquivo': os.path.basename(a), 'aba': '', 'erro': str(e), 'utilizavel': False, 'card': card})
                continue
            for x in d:
                x['card'] = card
                x['locacoes'] = card == 'planilha_locacoes'
            descs.extend(d)
            for r in planilha.registros(a, d):
                if card in SERV_CARDS:
                    r['origem'] = 'servicos'
                    if SERV_CARDS[card]:
                        r['mov_forcado'] = SERV_CARDS[card]
                elif card == 'planilha_locacoes':
                    r['origem'] = 'locacoes'
                regs.append(r)
    return descs, regs


def _dbr(d):
    return '%s/%s/%s' % (d[8:10], d[5:7], d[:4]) if d and len(d) >= 10 else (d or '')


def _servicos_relatorio(regs):
    """Relatórios de entrada e saída de serviços (PDF ou Excel) num formato padrão, para os cálculos de PIS/COFINS e da SCP."""
    out = {'saida': [], 'entrada': [], 'alertas': []}
    vistos = {}
    for r in regs:
        if r.get('aba') == 'Livro de serviços':
            continue  # livro Mod. 1: sem CNPJ e às vezes pelo líquido; serve para o confronto nota a nota, os cálculos seguem os XML
        if _sit_cancelada(r.get('situacao')) or r.get('data_cancelamento'):
            out.setdefault('canceladas', []).append({'lado': 'saida' if _mov_hint(r) == 'prestado' else 'entrada', 'numero': numero_doc(r.get('numero')),
                                                     'valor': round(r.get('valor_contabil', 0) or 0, 2), 'participante': r.get('participante') or ''})
            continue  # nota cancelada/substituída no relatório nunca entra nas contas
        lado = 'saida' if _mov_hint(r) == 'prestado' else ('entrada' if _mov_hint(r) == 'tomado' else None)
        if lado is None:
            continue
        ret_iss = r.get('iss_retido')
        x = {'arquivo': r['arquivo'], 'linha': r['linha'], 'numero': numero_doc(r.get('numero')), 'data': r.get('data_emissao') or r.get('data_lanc') or '',
             'participante': r.get('participante') or '', 'cnpj': r.get('cnpj') or '', 'servico': r.get('cod_servico') or '', 'descricao': r.get('descricao') or '',
             'valor': round(r.get('valor_contabil', 0) or 0, 2), 'base_iss': round(r.get('base_iss', 0) or 0, 2), 'aliq_iss': r.get('aliq_iss', 0) or 0,
             'iss': round(r.get('valor_iss', 0) or 0, 2), 'iss_retido': _sim(ret_iss) if ret_iss not in (None, '') else False,
             'irrf': round(r.get('irrf', 0) or 0, 2), 'csll': round(r.get('csll_ret', 0) or 0, 2), 'pis': round(r.get('pis_ret', 0) or 0, 2),
             'cofins': round(r.get('cofins_ret', 0) or 0, 2), 'inss': round(r.get('inss_ret', 0) or 0, 2), 'liquido': round(r.get('valor_liquido', 0) or 0, 2)}
        csrf = r.get('csrf_ret') or 0
        if csrf and not (x['pis'] or x['cofins'] or x['csll']):   # PIS 0,65 + COFINS 3 + CSLL 1 = 4,65%
            x['pis'], x['cofins'], x['csll'] = round(csrf * 0.65 / 4.65, 2), round(csrf * 3 / 4.65, 2), round(csrf * 1 / 4.65, 2)
        x['pf'] = len(x['cnpj']) == 11
        if x['valor'] <= 0:
            out['alertas'].append('%s linha %s: valor do serviço zerado.' % (r['arquivo'], r['linha']))
        if not x['cnpj']:
            out['alertas'].append('%s linha %s: sem CNPJ/CPF do %s.' % (r['arquivo'], r['linha'], 'tomador' if lado == 'saida' else 'prestador'))
        k = (lado, x['numero'], x['cnpj'], x['valor'])
        if x['numero'] and k in vistos:
            out['alertas'].append('%s linhas %s e %s: mesma nota repetida.' % (r['arquivo'], vistos[k], r['linha']))
        vistos[k] = r['linha']
        out[lado].append(x)
    for lado in ('saida', 'entrada'):
        ls = out[lado]
        out['tot_' + lado] = {k: round(sum(x[k] for x in ls), 2) for k in ('valor', 'iss', 'irrf', 'csll', 'pis', 'cofins', 'inss')}
        out['tot_' + lado]['notas'] = len(ls)
    return out


def _locacoes(regs):
    """Planilha mensal de locações: receitas (e despesas) de aluguel que não passam por NF-e/NFS-e."""
    linhas, alertas, vistos = [], [], {}
    for r in regs:
        mv = norm(r.get('movimento'))
        tipo = 'despesa' if mv.startswith(('desp', 'pag', 'tomad', 'entrada')) else 'receita'
        x = {'linha': r['linha'], 'arquivo': r['arquivo'], 'data': r.get('data_lanc') or r.get('data_emissao') or '', 'numero': r.get('numero') or '',
             'participante': r.get('participante') or '', 'cnpj': r.get('cnpj') or '', 'descricao': r.get('descricao') or '', 'valor': round(r.get('valor_contabil', 0), 2),
             'tipo': tipo, 'irrf': round(r.get('irrf', 0) or 0, 2)}
        if x['valor'] <= 0:
            alertas.append('Linha %s: valor zerado ou negativo.' % r['linha'])
        k = (x['numero'], x['cnpj'] or x['participante'], x['valor'], x['data'][:7])
        if x['valor'] and k in vistos:
            alertas.append('Linhas %s e %s: parecem a mesma locação lançada duas vezes.' % (vistos[k], r['linha']))
        vistos[k] = r['linha']
        linhas.append(x)
    sem_cli = [x['linha'] for x in linhas if not x['participante'] and not x['cnpj']]
    if sem_cli and len(sem_cli) < len(linhas):
        alertas.append('Linha(s) %s: sem cliente.' % ', '.join(map(str, sem_cli[:20])))
    return {'linhas': linhas, 'alertas': alertas, 'receita': round(sum(x['valor'] for x in linhas if x['tipo'] == 'receita'), 2),
            'despesa': round(sum(x['valor'] for x in linhas if x['tipo'] == 'despesa'), 2)}


def _sit_cancelada(v):
    t = norm(v)
    return bool(t) and (t.startswith('cancel') or t in ('02', '03', '2', '3', 'c') or 'cancelad' in t or 'denegad' in t or t.startswith('substituid'))


def agrupar_lancamentos(regs):
    """Agrupa as linhas da planilha por nota. Linhas da mesma nota separadas por outras notas = duplicidade."""
    grupos = {}
    ordem = []
    seq = defaultdict(int)
    for r in regs:
        k_aba = (r['arquivo'], r['aba'])
        seq[k_aba] += 1
        r['_seq'] = seq[k_aba]
        ch = r.get('chave') or ''
        if len(ch) not in (44, 50):
            ch = ''
        hint = _mov_hint(r)
        mov_k = hint if hint in ('entrada', 'saida') else ('servico' if hint in ('prestado', 'tomado', 'servico') else hint)
        if ch:
            k = (k_aba, 'ch', ch)
        elif r['aba'] == 'Livro de serviços':
            k = (k_aba, 'l', r['linha'])  # no livro cada linha é uma nota; sem CNPJ, números iguais são de prestadores diferentes
        else:
            k = (k_aba, 'n', numero_doc(r.get('numero')), digitos(r.get('cnpj')), (r.get('serie') or '').strip().lstrip('0'), mov_k if not r.get('cnpj') else '')
        if k not in grupos:
            grupos[k] = {'linhas': [], 'chave': ch}
            ordem.append(k)
        grupos[k]['linhas'].append(r)
    lancs = []
    for k in ordem:
        g = grupos[k]
        ls = g['linhas']
        blocos, atual = [], [ls[0]]
        for a, b in zip(ls, ls[1:]):
            if b['_seq'] - a['_seq'] > 1:
                blocos.append(atual); atual = [b]
            else:
                atual.append(b)
        blocos.append(atual)
        # relatório que espalha as linhas da mesma nota (ordenado por CFOP/item): junta os pedaços;
        # só é duplicidade o bloco que REPETE o conteúdo de outro
        def assin(b):
            return tuple(sorted((r.get('cfop', ''), round(r.get('valor_contabil', 0), 2), round(r.get('valor_icms', 0), 2)) for r in b))
        vistos, unidos, dups = set(), [], []
        for b in blocos:
            sg = assin(b)
            if sg in vistos and len(blocos) > 1:
                dups.append(b)
            else:
                vistos.add(sg); unidos.extend(b)
        blocos = [unidos] + dups
        p = ls[0]
        hints = Counter(_mov_hint(x) for x in ls if _mov_hint(x))
        lancs.append({
            'chave': g['chave'], 'numero': numero_doc(p.get('numero')) or (numero_doc(g['chave'][25:34]) if len(g['chave']) == 44 else ''),
            'serie': (p.get('serie') or '').strip(), 'cnpj': p.get('cnpj') or (g['chave'][6:20] if len(g['chave']) == 44 else ''),
            'cnpj_da_chave': not p.get('cnpj') and len(g['chave']) == 44,
            'participante': p.get('participante', ''), 'data_emissao': p.get('data_emissao', ''), 'data_lanc': p.get('data_lanc', ''),
            'mov_hint': hints.most_common(1)[0][0] if hints else '', 'blocos': blocos, 'linhas': blocos[0],
            'arquivo': p['arquivo'], 'aba': p['aba'], 'valor': round(sum(x.get('valor_contabil', 0) for x in blocos[0]), 2),
            'ref': '%s › %s › linha %s' % (p['arquivo'], p['aba'], ', '.join(str(x['linha']) for x in blocos[0][:6]) + ('…' if len(blocos[0]) > 6 else '')),
            'usado': False,
            'cancelada_sis': any(_sit_cancelada(x.get('situacao')) or bool(x.get('data_cancelamento')) for x in blocos[0]),
            'data_cancelamento': next((x.get('data_cancelamento') for x in blocos[0] if x.get('data_cancelamento')), ''),
        })
    return lancs


# --------------------------------------------------------------------------------------------- casamento

COMPAT = {'entrada': {'entrada', 'frete'}, 'saida': {'saida'}, 'frete': {'frete', 'entrada'}, 'servico': {'prestado', 'tomado'},
          'prestado': {'prestado'}, 'tomado': {'tomado'}, '': {'entrada', 'saida', 'frete', 'prestado', 'tomado'}}


def casar(docs, lancs):
    por_chave = {d['chave']: d for d in docs if d.get('chave')}
    por_num = defaultdict(list)
    for d in docs:
        por_num[d['numero']].append(d)
    ligacoes = defaultdict(list)  # id(doc) -> [(lanc, como)]
    fracos = []
    for L in lancs:
        alvo, como = None, ''
        if L['chave'] and L['chave'] in por_chave:
            alvo, como = por_chave[L['chave']], 'chave'
        if alvo is None and L['numero']:
            cands = [d for d in por_num.get(L['numero'], []) if d['mov'] in COMPAT.get(L['mov_hint'], COMPAT[''])]
            if L['cnpj']:
                c2 = [d for d in cands if digitos(d['part_cnpj']) == digitos(L['cnpj'])] or \
                     [d for d in cands if raiz(d['part_cnpj']) == raiz(L['cnpj']) and raiz(L['cnpj'])]
                if c2:
                    alvo, como = c2[0], 'numero_cnpj'
            if alvo is None:
                c3 = [d for d in cands if igual(d['valor'], L['valor'], max(TOL, 0.005 * d['valor']))]
                if len(c3) == 1:
                    alvo, como = c3[0], 'numero_valor'
            if alvo is None and not L['cnpj'] and len(cands) == 1:
                fracos.append((L, cands[0]))  # só pelo número: espera os vínculos fortes (outra linha com o mesmo número pode ser a certa)
                continue
        if alvo is not None:
            L['usado'] = True
            L['como'] = como
            ligacoes[id(alvo)].append(L)
    for L, d in fracos:
        if L['chave'] == '' and ligacoes.get(id(d)) and any(x.get('aba') == 'Livro de serviços' for x in [L] + ligacoes[id(d)]):
            continue  # no livro, número repetido com outro valor é outra nota (outro prestador)
        L['usado'] = True
        L['como'] = 'numero'
        ligacoes[id(d)].append(L)
    return ligacoes


# --------------------------------------------------------------------------------------------- grupos por CFOP/CST

def _grupos_xml_nfe(d, por_cfop, por_cst):
    gs = {}
    for it in d['itens']:
        ic = it.get('icms') or {}
        cst = ic.get('CST', '')
        com_st = cst in CST_COM_ST or ic.get('vICMSST', 0) > 0
        k = (it['CFOP'] if por_cfop else '*', cst if por_cst else '*', com_st)
        g = gs.setdefault(k, {'com_st': com_st, 'cfop': it['CFOP'], 'cst': cst, 'orig': ic.get('orig', ''), 'vc': 0.0, 'bc': 0.0, 'icms': 0.0, 'bcst': 0.0,
                              'st': 0.0, 'ipi': 0.0, 'pis': 0.0, 'cofins': 0.0, 'pis_bc': 0.0, 'credsn': 0.0, 'aliq': set(),
                              'cst_pis': set(), 'cfops': set(), 'csts': set(), 'itens': 0, 'prod': [], 'det': []})
        ipi = (it.get('ipi') or {}).get('vIPI', 0)
        g['vc'] += it['vProd'] - it['vDesc'] + it['vFrete'] + it['vSeg'] + it['vOutro'] + ipi + ic.get('vICMSST', 0) + ic.get('vFCPST', 0)
        g['bc'] += ic.get('vBC', 0); g['icms'] += ic.get('vICMS', 0); g['bcst'] += ic.get('vBCST', 0); g['st'] += ic.get('vICMSST', 0)
        g['ipi'] += ipi; g['bc_ipi'] = g.get('bc_ipi', 0) + (it.get('ipi') or {}).get('vBC', 0); g['pis'] += (it.get('pis') or {}).get('vPIS', 0); g['cofins'] += (it.get('cofins') or {}).get('vCOFINS', 0)
        g['pis_bc'] += (it.get('pis') or {}).get('vBC', 0)
        g['credsn'] += ic.get('vCredICMSSN', 0)
        if ic.get('pICMS'):
            g['aliq'].add(round(ic['pICMS'], 2))
        if (it.get('pis') or {}).get('CST'):
            g['cst_pis'].add(it['pis']['CST'])
        g['cfops'].add(it['CFOP']); g['csts'].add(cst)
        g['itens'] += 1
        g['det'].append({'xProd': it['xProd'], 'NCM': it['NCM'], 'vUnCom': it['vUnCom'], 'qCom': it['qCom'],
                         'vProd': it['vProd'] - it['vDesc'] + it['vFrete'] + it['vSeg'] + it['vOutro'] + ipi + ic.get('vICMSST', 0)})
        if len(g['prod']) < 3:
            g['prod'].append(it['xProd'][:40])
    for g in gs.values():
        for k in ('vc', 'bc', 'icms', 'bcst', 'st', 'ipi', 'pis', 'cofins', 'pis_bc', 'credsn'):
            g[k] = round(g[k], 2)
    return list(gs.values())


def _grupos_sis(linhas, por_cfop, por_cst, csosn_ctx):
    gs = {}
    for r in linhas:
        o, t = _cst(r.get('cst_icms'), csosn_ctx)
        k = (r.get('cfop', '') if por_cfop else '*', t if por_cst else '*')
        g = gs.setdefault(k, {'cfop': r.get('cfop', ''), 'cst': t, 'orig': o, 'vc': 0.0, 'bc': 0.0, 'icms': 0.0, 'bcst': 0.0, 'st': 0.0,
                              'ipi': 0.0, 'pis': 0.0, 'cofins': 0.0, 'pis_bc': 0.0, 'aliq': set(), 'cst_pis': set(), 'linhas': [],
                              'iss': 0.0, 'aliq_iss': 0.0, 'iss_ret': None, 'ret': defaultdict(float)})
        vc = r.get('valor_contabil', 0)
        if not vc and r.get('continuacao'):
            vc = r.get('base_icms', 0) + r.get('isentas', 0) + r.get('outras', 0)  # livro fiscal: linha de CFOP sem valor contábil
        g['vc'] += vc; g['bc'] += r.get('base_icms', 0); g['icms'] += r.get('valor_icms', 0)
        g['bcst'] += r.get('base_st', 0); g['st'] += r.get('valor_st', 0); g['ipi'] += r.get('valor_ipi', 0)
        g['pis'] += r.get('valor_pis', 0); g['cofins'] += r.get('valor_cofins', 0); g['pis_bc'] += r.get('base_pis', 0)
        if r.get('aliq_icms'):
            g['aliq'].add(round(r['aliq_icms'] * (100 if r['aliq_icms'] < 1 else 1), 2))
        if r.get('cst_pis'):
            g['cst_pis'].add(_pc(r['cst_pis']))
        g['linhas'].append(r['linha'])
    for g in gs.values():
        for k in ('vc', 'bc', 'icms', 'bcst', 'st', 'ipi', 'pis', 'cofins', 'pis_bc'):
            g[k] = round(g[k], 2)
    return list(gs.values())


def _parear(X, S, cfop_ok, cst_ok):
    """Pareia grupos do XML com grupos do sistema (guloso pela melhor nota de afinidade)."""
    if len(X) == 1 and len(S) == 1:
        return [(0, 0)], [], []
    cand = []
    for i, x in enumerate(X):
        for j, s in enumerate(S):
            sc = 0
            if igual(x['vc'], s['vc'], max(TOL, 0.005 * x['vc'])):
                sc += 4
            ok, exato = cfop_ok(x, s)
            sc += 2 if exato else (1 if ok else 0)
            sc += 1 if cst_ok(x, s) else 0
            cand.append((sc, -abs(x['vc'] - s['vc']), i, j))
    cand.sort(reverse=True)
    ux, us, pares = set(), set(), []
    for sc, _, i, j in cand:
        if i in ux or j in us:
            continue
        if sc == 0 and len(X) != len(S):
            continue
        pares.append((i, j)); ux.add(i); us.add(j)
    return sorted(pares), [i for i in range(len(X)) if i not in ux], [j for j in range(len(S)) if j not in us]


# --------------------------------------------------------------------------------------------- regras por tipo

def _par_registro(x, s, extra=None):
    p = {'xml': {'cfop': x.get('cfop', ''), 'cst': (x.get('orig', '') + x.get('cst', '')) if x.get('cst') and x.get('cst') not in CSOSN else x.get('cst', ''),
                 'vc': x.get('vc', 0), 'bc': x.get('bc', 0), 'icms': x.get('icms', 0), 'st': x.get('st', 0), 'ipi': x.get('ipi', 0),
                 'pis': x.get('pis', 0), 'cofins': x.get('cofins', 0), 'cst_pis': ','.join(sorted(x.get('cst_pis', []))),
                 'aliq': ','.join('%g' % a for a in sorted(x.get('aliq', []))), 'bc_ipi': round(x.get('bc_ipi', 0), 2),
                 'orig': x.get('orig', ''), 'cst_n': x.get('cst', ''), 'paliq': max(x.get('aliq') or [0])} if x else None,
         'sis': {'cfop': s.get('cfop', ''), 'cst': (s.get('orig', '') + s.get('cst', '')), 'vc': s.get('vc', 0), 'bc': s.get('bc', 0),
                 'icms': s.get('icms', 0), 'st': s.get('st', 0), 'ipi': s.get('ipi', 0), 'pis': s.get('pis', 0), 'cofins': s.get('cofins', 0),
                 'cst_pis': ','.join(sorted(s.get('cst_pis', []))), 'aliq': ','.join('%g' % a for a in sorted(s.get('aliq', []))),
                 'linhas': s.get('linhas', [])} if s else None}
    if extra:
        p.update(extra)
    return p


def regras_nfe(n, d, L, ctx):
    raw = d['raw']
    cols = ctx['cols']
    por_cfop, por_cst = 'cfop' in cols, 'cst_icms' in cols
    csosn = raw['emit'].get('crt') in ('1', '4') and n.mov == 'entrada'
    X = _grupos_xml_nfe(raw, por_cfop, por_cst)
    S = _grupos_sis(L['linhas'], por_cfop, por_cst, csosn)
    entrada = n.mov == 'entrada' and not d['propria']

    # 1º dígito da entrada pela UF do fornecedor × UF da empresa (não pelo CFOP dele: posto, restaurante e loja
    # de outro estado emitem 5xxx na venda presencial, mas para a empresa a compra é interestadual)
    uf_f, uf_e = (d['part_uf'] or '').upper(), (ctx['uf'] or '').upper()
    pri_uf = ('3' if uf_f == 'EX' else ('1' if uf_f == uf_e else '2')) if uf_f and uf_e else ''

    def esperado(x):
        if not entrada:
            return x['cfop']
        e = cfop_entrada_esperado(x['cfop'], x['st'] > 0 or x['cst'] in CST_COM_ST)
        return (pri_uf + e[1:]) if pri_uf and len(e) == 4 else e

    def pri_forn(x):
        return {'5': '1', '6': '2', '7': '3'}.get(x['cfop'][:1], '')

    def cfop_ok(x, s):
        if not entrada:
            return (s['cfop'] == x['cfop'], s['cfop'] == x['cfop'])
        e = esperado(x)
        ok = s['cfop'][:1] in {e[:1], pri_forn(x)} and s['cfop'][1:] in sufixos_entrada_aceitos(x['cfop'])
        return (ok, s['cfop'] == e)

    def aceitos_cst(x):
        if not entrada:
            return {x['cst'][-2:] if x['cst'] not in CSOSN else x['cst']}
        return CST_ENTRADA.get(x['cst'], {x['cst'][-2:]})

    def cst_ok(x, s):
        return s['cst'] in aceitos_cst(x)

    pares, sobra_x, sobra_s = _parear(X, S, cfop_ok, cst_ok)
    for idx, (i, j) in enumerate(pares):
        x, s = X[i], S[j]
        n.pares.append(_par_registro(x, s, {'esperado': esperado(x), 'desc_xml': desc_cfop(x['cfop']), 'desc_sis': desc_cfop(s['cfop']),
                                            'prod': x['prod'], 'itens': x['itens']}))
        suf = s['cfop'][1:]
        # ---- CFOP
        if por_cfop:
            if not s['cfop']:
                n.ap('CFOP_BRANCO', 'Linha lançada sem CFOP.', x['cfop'], '', 'Informar o CFOP %s.' % esperado(x), idx, 'cfop')
            elif not entrada:
                if s['cfop'] != x['cfop']:
                    n.ap('CFOP_SAIDA', 'O CFOP lançado não é o da nota emitida.', '%s %s' % (x['cfop'], desc_cfop(x['cfop'])),
                         '%s %s' % (s['cfop'], desc_cfop(s['cfop'])), 'Trocar %s por %s.' % (s['cfop'], x['cfop']), idx, 'cfop', s['vc'])
            else:
                e = esperado(x)
                ok, _ = cfop_ok(x, s)
                if s['cfop'][:1] != e[:1] and s['cfop'][:1] == pri_forn(x) and s['cfop'][1:] in sufixos_entrada_aceitos(x['cfop']):
                    n.ap('CFOP_UF_PRESENCIAL', 'Fornecedor em %s emitiu %s (venda presencial, interna para ele). Para a empresa em %s a compra vem de outro estado.' % (uf_f, x['cfop'], uf_e),
                         '%s (fornecedor em %s)' % (x['cfop'], uf_f), s['cfop'], 'Usar %s.' % (e[0] + s['cfop'][1:]), idx, 'cfop')
                elif s['cfop'][:1] != e[:1] and s['cfop'][1:] in sufixos_entrada_aceitos(x['cfop']):
                    uf = d['part_uf'] or '?'
                    n.ap('CFOP_UF', 'Fornecedor em %s e empresa em %s: a entrada é %s.' % (uf, ctx['uf'] or '?', 'interestadual (2xxx)' if e[0] == '2' else 'estadual (1xxx)' if e[0] == '1' else 'do exterior (3xxx)'),
                         '%s (fornecedor)' % x['cfop'], s['cfop'], 'Trocar %s por %s.' % (s['cfop'], e[0] + s['cfop'][1:]), idx, 'cfop', s['vc'])
                elif not ok:
                    n.ap('CFOP_INCOMPATIVEL', 'Nota do fornecedor com CFOP %s (%s). O CFOP lançado não corresponde a essa operação.' % (x['cfop'], desc_cfop(x['cfop']) or 'ver tabela'),
                         '%s → esperado %s' % (x['cfop'], e), '%s %s' % (s['cfop'], desc_cfop(s['cfop'])),
                         'Usar %s%s.' % (e, ' (ou o CFOP de compra conforme o destino: %s1 industrialização, %s56 uso e consumo, %s51 ativo)' % (e[:2], e[0] + '5', e[0] + '5') if x['cfop'][1:] in ('101', '102') else ''),
                         idx, 'cfop', s['vc'])
        # ---- CST com ST × CFOP com ST (CST 60 pede 1407/2407, não 1556/2556; sem ST, o contrário)
        if entrada and por_cfop and s['cfop'] and suf in ST_PAR | ST_PAR_INV:
            _regras_st_cfop(n, x, s, idx)
        # ---- destino provável do item (descrição/NCM/valor) × CFOP lançado
        if entrada and por_cfop and s['cfop'] and sufixos_entrada_aceitos(x['cfop']) is COMPRAS:
            _regras_destino(n, x, s, idx, ctx)
        # ---- CST ICMS
        if por_cst and s['cst']:
            ac = aceitos_cst(x)
            if s['cst'] not in ac:
                pref = sorted(ac, key=lambda c: (c != x['cst'][-2:], c))[0] if ac else ''
                if entrada:
                    txt = 'Nota veio com %s %s (%s). O CST lançado não é compatível.' % ('CSOSN' if x['cst'] in CSOSN else 'CST', x['cst'], CST_DESC.get(x['cst'], ''))
                    if s['cst'] in ('40', '41', '50') and x['cst'] not in ('40', '41', '50', '300', '400', '102', '103'):
                        txt = 'Nota tributada (%s %s, %s) lançada como %s (%s). Para não tomar o crédito o CST continua o da tributação; o crédito é que fica zerado.' % (
                            'CSOSN' if x['cst'] in CSOSN else 'CST', x['cst'], CST_DESC.get(x['cst'], ''), s['cst'], CST_DESC.get(s['cst'], ''))
                    n.ap('CST_ICMS', txt,
                         x['cst'], s['orig'] + s['cst'], 'Usar CST %s%s%s.' % (pref, (' com origem %s' % (s['orig'] or x['orig'])) if (s['orig'] or x['orig']) else '', (' (aceitos: ' + ', '.join(sorted(ac)) + ')') if len(ac) > 1 else ''), idx, 'cst')
                else:
                    n.ap('CST_SAIDA', 'CST lançado difere do XML.', x['orig'] + x['cst'], s['orig'] + s['cst'], 'Trocar para %s%s.' % (x['orig'], x['cst'][-2:]), idx, 'cst')
            if s['orig'] and x['orig'] and s['orig'] != x['orig'] and x['cst'] not in CSOSN:
                n.ap('CST_ORIGEM', 'Origem %s no XML (%s), %s no sistema.' % (x['orig'], 'nacional' if x['orig'] == '0' else 'estrangeira/importada' if x['orig'] in '1267' else 'nacional c/ conteúdo importado',
                     s['orig']), x['orig'], s['orig'], 'Ajustar a origem para %s.' % x['orig'], idx, 'cst')
        # ---- ICMS
        if 'valor_icms' in cols:
            if entrada:
                destac = x['credsn'] if x['cst'] in CSOSN else x['icms']
                if suf in USO_CONSUMO and s['icms'] > TOL:
                    n.ap('CREDITO_USO_CONSUMO', 'Entrada em %s (uso e consumo) não pode ter crédito de ICMS (LC 87/96, art. 33, I).' % s['cfop'], 'destacado ' + brl(destac),
                         'crédito ' + brl(s['icms']), 'Zerar o crédito de ICMS desta linha.', idx, 'icms', s['icms'])
                elif suf in COMBUSTIVEL_CONSUMO and s['icms'] > TOL:
                    n.ap('CREDITO_COMBUSTIVEL', 'Combustível para consumo só dá crédito de ICMS se usado na produção ou no transporte da própria mercadoria.', 'destacado ' + brl(destac),
                         'crédito ' + brl(s['icms']), 'Confirmar o uso; se for frota administrativa, zerar o crédito.', idx, 'icms', s['icms'])
                elif suf in ATIVO and s['icms'] > TOL:
                    n.ap('CREDITO_ATIVO', 'Ativo imobilizado: o crédito vai para o CIAP, 1/48 por mês (LC 87/96, art. 20, §5º).', 'destacado ' + brl(destac),
                         'crédito ' + brl(s['icms']), 'Lançar sem crédito na entrada e controlar no CIAP.', idx, 'icms', s['icms'])
                elif s['icms'] > destac + TOL:
                    n.ap('CREDITO_MAIOR', 'O crédito lançado passa o ICMS %s na nota.' % ('permitido (Simples)' if x['cst'] in CSOSN else 'destacado'),
                         destac, s['icms'], 'Reduzir o crédito para %s.' % brl(destac), idx, 'icms', s['icms'] - destac)
                elif destac > TOL and s['icms'] <= TOL and suf in REVENDA_INDUSTRIA:
                    n.ap('CREDITO_NAO_TOMADO', 'Compra para %s com ICMS destacado e sem crédito.' % ('industrialização' if suf in ('101', '401', '116', '120', '122', '124') else 'comercialização'),
                         destac, s['icms'], 'Conferir: se a operação dá crédito, lançar %s.' % brl(destac), idx, 'icms', destac)
                elif TOL < s['icms'] < destac - TOL:
                    n.ap('CREDITO_MENOR', 'Crédito lançado abaixo do destacado.', destac, s['icms'], 'Conferir o crédito (destacado %s).' % brl(destac), idx, 'icms', destac - s['icms'])
                if 'base_icms' in cols and s['icms'] > TOL and x['cst'] not in CSOSN and not igual(s['bc'], x['bc']) and suf not in USO_CONSUMO | ATIVO:
                    n.ap('BASE_ICMS', 'Base de cálculo lançada difere da nota.', x['bc'], s['bc'], 'Ajustar a base para %s.' % brl(x['bc']), idx, 'bc', abs(x['bc'] - s['bc']))
            else:
                nat = d.get('natureza', 'normal')
                if not igual(s['icms'], x['icms']):
                    if n.mov == 'entrada' and nat == 'dev_venda' and s['icms'] < x['icms']:
                        n.ap('CREDITO_DEV_VENDA', 'Na devolução de venda a empresa recupera o ICMS que debitou na venda: o destacado na nota (%s) entra como crédito.' % brl(x['icms']),
                             x['icms'], s['icms'], 'Lançar o crédito de %s.' % brl(x['icms']), idx, 'icms', x['icms'] - s['icms'])
                    elif n.mov == 'saida' and nat == 'dev_compra' and suf == '556' and s['icms'] <= TOL:
                        n.ap('DEV_USO_CONSUMO_ICMS', 'Devolução de material de uso e consumo: se a compra entrou sem crédito, o débito destacado (%s) é anulado por ajuste. Confira.' % brl(x['icms']),
                             x['icms'], s['icms'], 'Confirmar que a compra foi lançada sem crédito.', idx, 'icms')
                    elif n.mov == 'entrada':
                        n.ap('ICMS_ENTRADA_PROPRIA', 'ICMS lançado difere do destacado na nota de entrada própria.', x['icms'], s['icms'], 'Ajustar o ICMS para %s.' % brl(x['icms']), idx, 'icms', abs(x['icms'] - s['icms']))
                    else:
                        n.ap('ICMS_SAIDA', 'ICMS debitado difere do XML.', x['icms'], s['icms'], 'Ajustar o ICMS para %s.' % brl(x['icms']), idx, 'icms', abs(x['icms'] - s['icms']))
                if 'base_icms' in cols and not igual(s['bc'], x['bc']):
                    n.ap('BASE_ICMS', 'Base de cálculo lançada difere do XML.', x['bc'], s['bc'], 'Ajustar a base para %s.' % brl(x['bc']), idx, 'bc', abs(x['bc'] - s['bc']))
        if 'aliq_icms' in cols and s['aliq'] and x['aliq'] and s['aliq'] != x['aliq'] and (not entrada or s['icms'] > TOL):
            n.ap('ALIQ_ICMS', 'Alíquota lançada difere do XML.', ', '.join('%g%%' % a for a in sorted(x['aliq'])), ', '.join('%g%%' % a for a in sorted(s['aliq'])),
                 'Ajustar a alíquota.', idx, 'aliq')
        # ---- ST / IPI
        if 'valor_st' in cols and not igual(s['st'], x['st']):
            n.ap('ST_VALOR', 'ICMS-ST lançado difere do XML.', x['st'], s['st'], 'Ajustar o ICMS-ST para %s.' % brl(x['st']), idx, 'st', abs(x['st'] - s['st']))
        if entrada and suf in USO_CONSUMO and 'valor_ipi' in cols and s['ipi'] > TOL:
            n.ap('IPI_USO_CONSUMO', 'Entrada em %s (uso e consumo) não pode ter crédito de IPI: o IPI vai para o custo.' % s['cfop'], 'IPI da nota ' + brl(x['ipi']),
                 'crédito ' + brl(s['ipi']), 'Zerar o IPI creditado e somá-lo ao valor contábil.', idx, 'ipi', s['ipi'])
        elif 'valor_ipi' in cols and (s['ipi'] > x['ipi'] + TOL or (not entrada and not igual(s['ipi'], x['ipi']))):
            n.ap('IPI_VALOR', 'IPI lançado %s o do XML.' % ('acima do' if s['ipi'] > x['ipi'] else 'diferente do'), x['ipi'], s['ipi'], 'Ajustar o IPI para %s.' % brl(x['ipi']), idx, 'ipi', abs(x['ipi'] - s['ipi']))
        # ---- PIS/COFINS
        if entrada:
            _regras_pc_entrada(n, x, s, idx, cols, raw, suf)
        else:
            if 'valor_pis' in cols and (not igual(s['pis'], x['pis']) or ('valor_cofins' in cols and not igual(s['cofins'], x['cofins']))):
                n.ap('PIS_COFINS_VALOR', 'PIS/COFINS lançados diferem do XML.', 'PIS %s / COFINS %s' % (brl(x['pis']), brl(x['cofins'])),
                     'PIS %s / COFINS %s' % (brl(s['pis']), brl(s['cofins'])), 'Ajustar PIS/COFINS da linha.', idx, 'pis', abs(x['pis'] - s['pis']) + abs(x['cofins'] - s['cofins']))
            if 'cst_pis' in cols and s['cst_pis'] and x['cst_pis'] and s['cst_pis'] != x['cst_pis']:
                n.ap('PC_CST_SAIDA', 'CST de PIS/COFINS difere do XML.', ','.join(sorted(x['cst_pis'])), ','.join(sorted(s['cst_pis'])), 'Usar CST %s.' % ','.join(sorted(x['cst_pis'])), idx, 'cst_pis')
        if len(pares) > 1 and not igual(s['vc'], x['vc'], max(TOL, 0.005 * x['vc'])):
            n.ap('VALOR_LINHA', 'Valor da linha CFOP %s difere da soma dos itens da nota.' % (s['cfop'] or '?'), x['vc'], s['vc'], 'Redistribuir os valores entre as linhas.', idx, 'vc', abs(x['vc'] - s['vc']))
    for i in sobra_x:
        x = X[i]
        n.pares.append(_par_registro(x, None, {'esperado': esperado(x), 'desc_xml': desc_cfop(x['cfop']), 'prod': x['prod'], 'itens': x['itens']}))
        sug = esperado(x)
        misto = entrada and any(X[k].get('com_st') != x.get('com_st') for k in range(len(X)) if k != i)
        if misto and pares:
            irma = S[pares[0][1]]['cfop']  # linha que foi lançada: separa a parte com/sem ST no CFOP par dela
            if len(irma) == 4:
                suf = irma[1:]
                if x.get('com_st') and suf in ST_PAR:
                    sug = irma[0] + ST_PAR[suf]
                elif not x.get('com_st') and suf in ST_PAR_INV:
                    sug = irma[0] + ST_PAR_INV[suf]
        txt = ('A nota tem mercadoria com ST e sem ST. Os itens %s ST (CST %s, %d item(ns): %s) precisam de linha própria.' % ('com' if x.get('com_st') else 'sem', x['cst'], x['itens'], '; '.join(x['prod']))
               if misto else 'Itens da nota com CFOP %s / CST %s não aparecem no lançamento (%d item(ns): %s).' % (x['cfop'], x['cst'], x['itens'], '; '.join(x['prod'])))
        cor = 'Lançar a linha CFOP %s com %s' % (sug, brl(x['vc']))
        if misto and pares:
            xi, sj = X[pares[0][0]], S[pares[0][1]]
            if sj['vc'] > xi['vc'] + TOL:
                cor += ' e deixar %s na linha %s' % (brl(xi['vc']), sj['cfop'])
        n.ap('PARTE_NAO_LANCADA', txt, x['vc'], '', cor + '.', len(n.pares) - 1, 'vc', x['vc'])
    # nota com itens com ST e sem ST lançada numa linha só: um apontamento, com a divisão exata
    if entrada and por_cfop and len({bool(x.get('com_st')) for x in X}) == 2 and len(S) == 1 and len(S[0]['cfop']) == 4:
        s0 = S[0]; suf0, pri0 = s0['cfop'][1:], s0['cfop'][0]
        base = ST_PAR_INV.get(suf0, suf0 if suf0 in ST_PAR else None)
        if base and igual(s0['vc'], sum(x['vc'] for x in X), max(1.0, 0.005 * s0['vc'])):
            v_st = round(sum(x['vc'] for x in X if x.get('com_st')), 2)
            v_no = round(sum(x['vc'] for x in X if not x.get('com_st')), 2)
            p_st = '; '.join(p for x in X if x.get('com_st') for p in x['prod'])[:90]
            p_no = '; '.join(p for x in X if not x.get('com_st') for p in x['prod'])[:90]
            n.ap('ST_MISTO', 'A nota tem itens COM ST (%s: %s) e SEM ST (%s: %s), lançados juntos em %s.' % (brl(v_st), p_st, brl(v_no), p_no, s0['cfop']),
                 'com ST %s | sem ST %s' % (brl(v_st), brl(v_no)), '%s %s' % (s0['cfop'], brl(s0['vc'])),
                 'Separar em duas linhas: %s%s com %s e %s%s com %s.' % (pri0, base, brl(v_no), pri0, ST_PAR[base], brl(v_st)), 0, 'cfop', min(v_st, v_no))
    for j in sobra_s:
        s = S[j]
        n.pares.append(_par_registro(None, s, {'desc_sis': desc_cfop(s['cfop'])}))
        suf_j = s['cfop'][1:]
        tem_st_x = [bool(x.get('com_st')) for x in X]
        if entrada and suf_j in ST_PAR_INV and not any(tem_st_x):
            n.ap('CFOP_ST_SEM_ST', 'Parte da nota (%s) foi lançada em %s, CFOP de ST, mas nenhum item da nota tem ST.' % (brl(s['vc']), s['cfop']), 'sem ST', s['cfop'],
                 'Lançar essa parte em %s%s.' % (s['cfop'][0], ST_PAR_INV[suf_j]), len(n.pares) - 1, 'cfop', s['vc'])
            continue
        if entrada and suf_j in ST_PAR and tem_st_x and all(tem_st_x):
            n.ap('CFOP_ST', 'Parte da nota (%s) foi lançada em %s, sem ST, mas todos os itens têm ST.' % (brl(s['vc']), s['cfop']), 'com ST', s['cfop'],
                 'Lançar essa parte em %s%s.' % (s['cfop'][0], ST_PAR[suf_j]), len(n.pares) - 1, 'cfop', s['vc'])
            continue
        n.ap('LINHA_SOBRANDO', 'Linha CFOP %s / CST %s lançada sem itens equivalentes na nota.' % (s['cfop'] or '?', s['cst'] or '?'), '', s['vc'],
             'Conferir a linha (planilha linha %s).' % ', '.join(map(str, s['linhas'])), len(n.pares) - 1, 'vc', s['vc'])


# CFOP sem ST -> equivalente com ST (e o inverso)
ST_PAR = {'556': '407', '102': '403', '101': '401', '551': '406', '152': '409', '151': '408', '202': '411', '201': '410'}
ST_PAR_INV = {v: k for k, v in ST_PAR.items()}
ST_NOME = {'556': 'uso e consumo', '102': 'comercialização', '101': 'industrialização', '551': 'ativo imobilizado',
           '152': 'transferência p/ comercialização', '151': 'transferência p/ industrialização', '202': 'devolução de venda', '201': 'devolução de venda de produção'}


def _regras_st_cfop(n, x, s, idx):
    suf, pri = s['cfop'][1:], s['cfop'][:1]
    # CST que vale: o lançado na planilha, se houver; senão o da nota
    # mercadoria é com ST se a NOTA tem ST (mesmo que o CST lançado esteja errado) ou se o CST lançado é de ST
    cst_st = bool(x.get('com_st')) or (s.get('cst') or '') in CST_COM_ST
    cst_ref = x['cst'] if x.get('com_st') or not s.get('cst') else s['cst']
    origem = 'da nota' if cst_ref == x['cst'] else 'lançado'
    if cst_st and suf in ST_PAR:
        n.ap('CFOP_ST', 'CST %s %s é de mercadoria com ST: a entrada para %s com ST é %s%s, não %s.' % (cst_ref, origem, ST_NOME[suf], pri, ST_PAR[suf], s['cfop']),
             'CST %s%s' % (x['cst'], (', ST ' + brl(x['st'])) if x['st'] else ''), '%s %s' % (s['cfop'], desc_cfop(s['cfop'])),
             'Trocar %s por %s%s.' % (s['cfop'], pri, ST_PAR[suf]), idx, 'cfop', s['vc'])
    elif not cst_st and cst_ref and suf in ST_PAR_INV:
        base = ST_PAR_INV[suf]
        n.ap('CFOP_ST_SEM_ST', 'CST %s %s é de mercadoria sem ST: a entrada para %s sem ST é %s%s, não %s.' % (cst_ref, origem, ST_NOME[base], pri, base, s['cfop']),
             'CST %s, sem ST' % x['cst'], '%s %s' % (s['cfop'], desc_cfop(s['cfop'])), 'Trocar %s por %s%s.' % (s['cfop'], pri, base), idx, 'cfop', s['vc'])


def _regras_destino(n, x, s, idx, ctx):
    cats, motivo = defaultdict(float), {}
    for it in x.get('det', []):
        c = destino.classificar(it, ctx.get('ncms_vendidos', set()))
        if c:
            cats[c[0]] += it['vProd']
            motivo.setdefault(c[0], (c[1], c[2], it['xProd'].strip()[:50]))
    if not cats:
        return
    total = sum(it['vProd'] for it in x['det']) or 1
    cat, val = max(cats.items(), key=lambda kv: kv[1])
    if val < 0.5 * total:
        return
    rot, mot, item = motivo[cat]
    n.pares[-1]['destino'] = '%s (%s)' % (rot, mot)
    suf, pri = s['cfop'][1:], s['cfop'][:1]
    tem_st = x['st'] > 0 or x['cst'] in CST_COM_ST
    if cat == 'combustivel' and suf not in {'651', '652', '653'} and not ctx.get('vende_combustivel'):
        n.ap('DESTINO_COMBUSTIVEL', '"%s" é combustível (%s). Compra de combustível tem CFOP próprio.' % (item, mot), rot, '%s %s' % (s['cfop'], desc_cfop(s['cfop'])),
             'Usar %s653 (consumo) ou %s652 (industrialização).' % (pri, pri), idx, 'cfop')
    elif cat == 'consumo' and suf in REVENDA_INDUSTRIA:
        n.ap('DESTINO_CONSUMO', '"%s" tem cara de %s (%s) e a empresa não vende esse produto.' % (item, rot, mot), rot, '%s %s' % (s['cfop'], desc_cfop(s['cfop'])),
             'Se não for revenda nem insumo, usar %s%s, sem crédito.' % (pri, '407' if tem_st else '556'), idx, 'cfop')
    elif cat == 'ativo' and suf not in ATIVO:
        n.ap('DESTINO_ATIVO', '"%s" tem cara de %s (%s).' % (item, rot, mot), rot, '%s %s' % (s['cfop'], desc_cfop(s['cfop'])),
             'Se ficar na empresa como bem durável, usar %s%s (crédito pelo CIAP, 1/48).' % (pri, '406' if tem_st else '551'), idx, 'cfop')


def _regras_pc_entrada(n, x, s, idx, cols, raw, suf):
    forn = x['cst_pis']
    sis_cred = bool(s['cst_pis'] & PC_CREDITO) if s['cst_pis'] else ('valor_pis' in cols and s['pis'] > TOL)
    sis_sem = bool(s['cst_pis'] & PC_SEM_CREDITO) if s['cst_pis'] else ('valor_pis' in cols and s['pis'] <= TOL)
    if 'cst_pis' not in cols and 'valor_pis' not in cols:
        return
    revenda = suf in ('102', '403', '117', '113', '121', '118')
    if sis_cred and forn and forn <= PC_FORN_SEM_CREDITO and revenda:
        n.ap('PC_MONOFASICO', 'Fornecedor vendeu com CST PIS %s (monofásico/ST/alíquota zero). Revenda desse produto não dá crédito (Lei 10.833/03, art. 3º, I, b).' % ','.join(sorted(forn)),
             'CST ' + ','.join(sorted(forn)), ('CST ' + ','.join(sorted(s['cst_pis']))) if s['cst_pis'] else 'crédito ' + brl(s['pis'] + s['cofins']),
             'Lançar com CST 70/73 (sem crédito).', idx, 'cst_pis', s['pis'] + s['cofins'])
    elif suf in USO_CONSUMO and (sis_cred or s['pis'] > TOL or s['cofins'] > TOL):
        n.ap('PC_USO_CONSUMO', 'Entrada em %s (uso e consumo) não pode ter crédito de PIS/COFINS.' % s['cfop'], s['cfop'] + ' ' + desc_cfop(s['cfop']),
             ' / '.join(x for x in [('CST ' + ','.join(sorted(s['cst_pis']))) if s['cst_pis'] else '', ('PIS %s, COFINS %s' % (brl(s['pis']), brl(s['cofins']))) if s['pis'] or s['cofins'] else ''] if x),
             'Lançar com CST 70 (sem crédito) e zerar PIS/COFINS.', idx, 'cst_pis', s['pis'] + s['cofins'])
    elif sis_sem and suf in REVENDA_INDUSTRIA and forn & {'01', '02'}:
        base = x['vc'] - x['icms'] - x['ipi'] if raw['emissao'] >= '2023-05' else x['vc'] - x['ipi']
        n.ap('PC_NAO_TOMADO', 'Compra tributada no fornecedor (CST %s), lançada sem crédito de PIS/COFINS.' % ','.join(sorted(forn)),
             'CST ' + ','.join(sorted(forn)), ('CST ' + ','.join(sorted(s['cst_pis']))) if s['cst_pis'] else 'sem crédito',
             'Conferir: crédito estimado %s (base %s).' % (brl(base * (ALIQ_PIS_NC + ALIQ_COFINS_NC) / 100), brl(base)), idx, 'cst_pis', base * 0.0925)
    if suf in USO_CONSUMO:
        return  # já apontado acima: aqui não pode haver crédito nenhum
    if sis_cred and 'base_pis' in cols and s['pis_bc'] > TOL and x['icms'] > TOL and raw['emissao'] >= '2023-05':
        esperado = round(x['vc'] - x['icms'] - x['ipi'], 2)
        com_icms = min(abs(s['pis_bc'] - (x['vc'] - x['ipi'])), abs(s['pis_bc'] - (s['vc'] - s['ipi'])), abs(s['pis_bc'] - x['vc']), abs(s['pis_bc'] - s['vc']))
        if s['pis_bc'] > esperado + max(1.0, 0.002 * esperado) and com_icms <= max(1.0, 0.002 * esperado):
            n.ap('PC_BASE_ICMS', 'Desde 05/2023 o ICMS destacado sai da base do crédito (Lei 14.592/2023).', 'base ' + brl(esperado), 'base ' + brl(s['pis_bc']),
                 'Reduzir a base para %s.' % brl(esperado), idx, 'pis_bc', (s['pis_bc'] - esperado) * 0.0925)
    if sis_cred and 'base_pis' in cols and s['pis_bc'] > TOL:
        ap = round(s['pis'] / s['pis_bc'] * 100, 2) if s['pis'] else 0
        ac = round(s['cofins'] / s['pis_bc'] * 100, 2) if s['cofins'] else 0
        if (ap and abs(ap - ALIQ_PIS_NC) > 0.02) or (ac and abs(ac - ALIQ_COFINS_NC) > 0.02):
            n.ap('PC_ALIQ', 'No não cumulativo: PIS 1,65% e COFINS 7,6%.', 'PIS 1,65% / COFINS 7,6%', 'PIS %g%% / COFINS %g%%' % (ap, ac), 'Ajustar as alíquotas.', idx, 'pis')


def regras_cte(n, d, L, ctx):
    raw = d['raw']
    cols = ctx['cols']
    x = {'cfop': raw['CFOP'], 'cst': (raw['icms'] or {}).get('CST', '') or ('90' if 'SN' in (raw['icms'] or {}).get('grupo', '') else ''),
         'orig': '', 'vc': raw['vTPrest'], 'bc': (raw['icms'] or {}).get('vBC', 0), 'icms': (raw['icms'] or {}).get('vICMS', 0),
         'st': 0, 'ipi': 0, 'pis': 0, 'cofins': 0, 'aliq': {round((raw['icms'] or {}).get('pICMS', 0), 2)} - {0}, 'cst_pis': set()}
    S = _grupos_sis(L['linhas'], False, False, False)
    s = S[0] if S else {'cfop': '', 'cst': '', 'vc': 0, 'icms': 0, 'bc': 0}
    cfops = sorted({r.get('cfop', '') for r in L['linhas'] if r.get('cfop')})
    s['cfop'] = cfops[0] if cfops else ''
    csts = sorted({_cst(r.get('cst_icms'))[1] for r in L['linhas'] if r.get('cst_icms')})
    s['cst'] = csts[0] if csts else ''
    s['cst_pis'] = {_pc(r.get('cst_pis')) for r in L['linhas'] if r.get('cst_pis')}
    uf_ini = raw.get('UFIni') or ''
    pri = '1' if uf_ini and ctx['uf'] and uf_ini == ctx['uf'] else ('2' if uf_ini and ctx['uf'] else {'5': '1', '6': '2', '7': '3'}.get(x['cfop'][:1], '1'))
    esperado = pri + '353'
    tomadora = raiz(raw.get('tomador_cnpj')) == ctx['raiz']
    nfes = [c for c in raw.get('nfes', []) if c]
    ligadas = [ctx['chave_num'][c] for c in nfes if c in ctx['chave_num']]
    n.pares.append(_par_registro(x, s, {'esperado': esperado + ' / ' + pri + '352', 'desc_xml': 'CT-e %s, %s → %s' % (x['cfop'], uf_ini, raw.get('UFFim', '')),
                                        'desc_sis': desc_cfop(s['cfop']), 'prod': ['NF ' + v for v in ligadas][:4], 'itens': len(nfes)}))
    if not tomadora:
        if s['icms'] > TOL:
            n.ap('FRETE_NAO_TOMADOR_CREDITO', 'Tomador do CT-e é %s; quem paga o frete é que tem o crédito.' % (raw.get('tomador_cnpj') or '?'),
                 'tomador ' + (raw.get('tomador_cnpj') or '?'), 'crédito ' + brl(s['icms']), 'Estornar o crédito deste CT-e.', 0, 'icms', s['icms'])
        else:
            n.ap('FRETE_NAO_TOMADOR', 'Tomador do CT-e é %s. Confirme se este frete deve estar na escrita da empresa.' % (raw.get('tomador_cnpj') or '?'),
                 'tomador ' + (raw.get('tomador_cnpj') or '?'), '', 'Conferir com o cliente.', 0, 'cfop')
    if 'cfop' in cols:
        if not s['cfop']:
            n.ap('CFOP_BRANCO', 'Frete lançado sem CFOP.', x['cfop'], '', 'Informar %s (comércio) ou %s352 (indústria).' % (esperado, pri), 0, 'cfop')
        elif s['cfop'][1:] not in FRETE or s['cfop'][:1] != pri:
            n.ap('CFOP_FRETE', 'Aquisição de frete iniciado em %s%s.' % (uf_ini or '?', ' (mesmo estado)' if pri == '1' else ' (outro estado)'),
                 '%s (transportadora)' % x['cfop'], s['cfop'], 'Usar %s (comércio) ou %s352 (indústria).' % (esperado, pri), 0, 'cfop', s['vc'])
    if 'valor_icms' in cols and tomadora:
        if s['icms'] > x['icms'] + TOL:
            n.ap('CREDITO_MAIOR', 'Crédito de ICMS do frete acima do destacado no CT-e.', x['icms'], s['icms'], 'Reduzir o crédito para %s.' % brl(x['icms']), 0, 'icms', s['icms'] - x['icms'])
        elif x['icms'] > TOL and s['icms'] <= TOL:
            n.ap('CREDITO_NAO_TOMADO', 'CT-e com ICMS destacado e sem crédito lançado.', x['icms'], s['icms'], 'Conferir: se o frete é de compra para revenda/industrialização, lançar %s.' % brl(x['icms']), 0, 'icms', x['icms'])
    if 'cst_icms' in cols and s['cst'] and x['cst'] and s['cst'] not in {x['cst'][-2:], '90'}:
        n.ap('CST_ICMS', 'CST lançado difere do CT-e.', x['cst'], s['cst'], 'Usar CST %s.' % x['cst'], 0, 'cst')
    if tomadora and 'cst_pis' in cols and s['cst_pis'] and s['cst_pis'] <= PC_SEM_CREDITO:
        n.ap('PC_FRETE', 'Frete pago pela empresa na compra ou na venda dá crédito de PIS/COFINS no Lucro Real (Lei 10.833/03, art. 3º, IX).',
             '', 'CST ' + ','.join(sorted(s['cst_pis'])), 'Conferir e lançar com CST 50 se for frete de operação de compra/venda.', 0, 'cst_pis', x['vc'] * 0.0925)


def _sim(v):
    if isinstance(v, (int, float)):
        return v > 0
    t = str(v or '').strip().upper()
    if t in ('S', 'SIM', 'X', 'Y', 'YES', 'RETIDO', 'TRUE', '1'):
        return True
    return num(t) > 0


def _item_lc(v):
    """Item da LC 116 como 'NN.NN' a partir de '7.10', '0710', '070101' (cTribNac) ou '17.01.01'."""
    t = str(v or '').strip()
    if not t:
        return ''
    if '.' in t:
        p = [x for x in re.split(r'\D+', t) if x]
        if len(p) >= 2:
            return '%s.%s' % (p[0].zfill(2), p[1].zfill(2))
    d = digitos(t)
    if len(d) in (5, 6):
        d = d.zfill(6)[:4]
    elif len(d) == 3:
        d = d.zfill(4)
    return '%s.%s' % (d[:2], d[2:4]) if len(d) >= 4 else d


def regras_nfse(n, d, L, ctx):
    raw = d['raw']
    cols = ctx['cols']
    linhas = L['linhas']
    sis = {k: round(float(sum(r.get(k, 0) or 0 for r in linhas if isinstance(r.get(k, 0), (int, float)))), 2)
           for k in ('valor_contabil', 'valor_iss', 'irrf', 'csll_ret', 'inss_ret', 'pis_ret', 'cofins_ret', 'base_iss', 'valor_liquido')}
    cod_sis = next((r.get('cod_servico') for r in linhas if r.get('cod_servico')), '')
    aliq = next((r.get('aliq_iss') for r in linhas if r.get('aliq_iss')), 0)
    if aliq and aliq < 1:
        aliq *= 100
    ret_raw = [r.get('iss_retido') for r in linhas if r.get('iss_retido') not in (None, '')]
    n.pares.append({'xml': {'servico': _item_lc(raw.get('cTribNac', '')), 'base_iss': raw.get('vBC') or raw['vServ'], 'vliq': raw.get('vLiq'), 'vc': raw['vServ'], 'iss': raw['vISS'], 'aliq_iss': raw['pAliq'],
                            'iss_retido': 'Sim' if raw['issRetido'] else 'Não', 'irrf': raw['ret']['irrf'], 'csll': raw['ret']['csll'],
                            'inss': raw['ret']['inss'], 'pis_ret': raw['ret']['pis'], 'cofins_ret': raw['ret']['cofins']},
                    'sis': {'servico': _item_lc(cod_sis) if cod_sis else '', 'base_iss': sis['base_iss'] if 'base_iss' in cols else '', 'vliq': sis['valor_liquido'] if 'valor_liquido' in cols else '', 'vc': sis['valor_contabil'], 'iss': sis['valor_iss'], 'aliq_iss': aliq,
                            'iss_retido': ('Sim' if any(_sim(v) for v in ret_raw) else 'Não') if ret_raw else '', 'irrf': sis['irrf'],
                            'csll': sis['csll_ret'], 'inss': sis['inss_ret'], 'pis_ret': sis['pis_ret'], 'cofins_ret': sis['cofins_ret'],
                            'linhas': [r['linha'] for r in linhas]},
                    'servico': True, 'desc_xml': (raw.get('xDescServ') or '')[:90]})
    iss_lanc = sis['valor_iss']
    ret_num = sum(v for v in ret_raw if isinstance(v, (int, float)))
    if not iss_lanc and ret_num:
        iss_lanc = round(ret_num, 2)  # sistema que mostra o ISS só na coluna de retido
    if 'valor_iss' in cols and not igual(iss_lanc, raw['vISS']):
        n.ap('ISS_VALOR', 'ISS lançado difere do ISS da nota.', raw['vISS'], iss_lanc, 'Ajustar o ISS para %s.' % brl(raw['vISS']), 0, 'iss', abs(raw['vISS'] - iss_lanc))
    if 'aliq_iss' in cols and aliq and raw['pAliq'] and abs(aliq - raw['pAliq']) > 0.009:
        n.ap('ISS_ALIQ', 'Alíquota do ISS difere da nota.', '%g%%' % raw['pAliq'], '%g%%' % aliq, 'Usar %g%%.' % raw['pAliq'], 0, 'aliq_iss')
    if 'iss_retido' in cols:
        sis_ret = any(_sim(v) for v in ret_raw)
        if raw['issRetido'] and not sis_ret:
            txt = ('A empresa reteve o ISS do prestador: precisa lançar como retido e recolher a guia de ISS retido.' if n.mov == 'tomado'
                   else 'O tomador reteve o ISS: sem marcar como retido, o ISS sai na apuração e é pago duas vezes.')
            n.ap('ISS_RETIDO_NAO_LANCADO', txt, 'retido ' + brl(raw['vISS']), 'não retido', 'Marcar ISS retido (%s).' % brl(raw['vISS']), 0, 'iss_retido', raw['vISS'])
        elif not raw['issRetido'] and sis_ret:
            n.ap('ISS_RETIDO_INDEVIDO', 'A nota não tem retenção de ISS.', 'não retido', 'retido', 'Desmarcar a retenção.', 0, 'iss_retido', raw['vISS'])
    if 'cod_servico' in cols and cod_sis and raw.get('cTribNac') and _item_lc(cod_sis) != _item_lc(raw['cTribNac']):
        n.ap('SERVICO_CODIGO', 'Item da LC 116 lançado (%s) difere do da nota (%s).' % (_item_lc(cod_sis), _item_lc(raw['cTribNac'])), _item_lc(raw['cTribNac']), _item_lc(cod_sis),
             'Lançar com o item %s.' % _item_lc(raw['cTribNac']), 0, 'servico')
    base_x = raw.get('vBC') or raw['vServ']
    if 'base_iss' in cols and sis['base_iss'] and not igual(sis['base_iss'], base_x):
        n.ap('ISS_BASE', 'Base do ISS lançada difere da nota.', base_x, sis['base_iss'], 'Ajustar a base para %s.' % brl(base_x), 0, 'base_iss', abs(base_x - sis['base_iss']))
    if 'valor_liquido' in cols and sis['valor_liquido'] and raw.get('vLiq') and not igual(sis['valor_liquido'], raw['vLiq']):
        n.ap('SERVICO_LIQUIDO', 'Valor líquido lançado difere do líquido da nota (valor menos retenções).', raw['vLiq'], sis['valor_liquido'],
             'Conferir as retenções; líquido da nota %s.' % brl(raw['vLiq']), 0, 'vliq', abs(raw['vLiq'] - sis['valor_liquido']))
    for campo, chave, nome in (('irrf', 'irrf', 'IRRF'), ('csll_ret', 'csll', 'CSLL'), ('inss_ret', 'inss', 'INSS'), ('pis_ret', 'pis', 'PIS retido'), ('cofins_ret', 'cofins', 'COFINS retida')):
        if campo in cols and not igual(sis[campo], raw['ret'][chave]):
            n.ap('RETENCAO', '%s retido na nota: %s; lançado: %s.' % (nome, brl(raw['ret'][chave]), brl(sis[campo])), raw['ret'][chave], sis[campo],
                 'Ajustar o %s para %s.' % (nome, brl(raw['ret'][chave])), 0, chave if chave != 'csll' else 'csll', abs(raw['ret'][chave] - sis[campo]))


# --------------------------------------------------------------------------------------------- execução

def executar(sessao, cnpj_inf='', competencia_inf='', ajustes=None):
    t0 = datetime.datetime.now()
    docs_raw, eventos, pdfs, erros, inventario = carregar(sessao)
    cancelados = {e['chave'] for e in eventos if e.get('cancelamento')}
    emp = identificar_empresa(docs_raw, cnpj_inf)
    docs_raw = docs_raw + manuais_raw(sessao, emp)   # notas de serviço lançadas à mão entram como NFS-e
    docs = [_unificar(d, emp, cancelados) for d in docs_raw]
    # mesmo XML enviado duas vezes
    vistos, unicos = set(), []
    for d in docs:
        k = d['chave'] or (d['tipo'], d['numero'], d['part_cnpj'], d['valor'])
        if k in vistos:
            continue
        vistos.add(k); unicos.append(d)
    docs = unicos
    descs, regs = ler_planilhas(sessao, ajustes)
    loc_regs = [r for r in regs if r.get('origem') == 'locacoes']
    serv_rel = _servicos_relatorio([r for r in regs if r.get('origem') == 'servicos'])
    regs = [r for r in regs if r.get('origem') != 'locacoes']
    cols = set()
    for dsc in descs:
        if dsc.get('utilizavel'):
            cols |= set(dsc['mapa'].keys())
    lancs = agrupar_lancamentos(regs)
    colmap = {(x['arquivo'], x['aba']): set(x.get('mapa', {})) for x in descs if x.get('utilizavel')}
    for L in lancs:
        L['cols'] = colmap.get((L['arquivo'], L['aba']))
    for L in lancs:  # na chave vem o EMITENTE: em nota própria (saída) ele é a empresa, não o participante
        if L.get('cnpj_da_chave') and raiz(L['cnpj']) == emp['raiz']:
            L['cnpj'] = ''
    competencia = competencia_inf or ''
    if not competencia:
        meses = Counter(_mes(L['data_lanc'] or L['data_emissao']) for L in lancs if (L['data_lanc'] or L['data_emissao']))
        if not meses:
            meses = Counter(_mes(d['emissao']) for d in docs if d['mov'] in ('saida', 'prestado'))
        if not meses:
            meses = Counter(_mes(d['emissao']) for d in docs)
        competencia = meses.most_common(1)[0][0] if meses else ''
    chave_num = {d['chave']: d['numero'] for d in docs if d.get('chave')}
    # o que a empresa vende (NCM das próprias saídas de mercadoria): item com essa NCM na entrada é mercadoria/insumo dela
    venda = {'101', '102', '103', '104', '105', '106', '107', '108', '109', '110', '111', '112', '113', '114', '115', '116', '117',
             '118', '119', '120', '122', '123', '124', '125', '401', '402', '403', '404', '405', '651', '652', '653', '654', '655', '656', '667'}
    ncms_vendidos = {it['NCM'][:4] for d in docs if d['tipo'] == 'nfe' and d['mov'] == 'saida' and d['propria'] and not d['cancelada']
                     for it in d['raw']['itens'] if it['CFOP'][1:] in venda and it.get('NCM')}
    ctx = {'raiz': emp['raiz'], 'uf': emp['uf'], 'cols': cols, 'competencia': competencia, 'chave_num': chave_num,
           'ncms_vendidos': ncms_vendidos, 'vende_combustivel': bool(ncms_vendidos & {'2710', '2207', '2711'})}
    lig = casar(docs, lancs)
    _cancelamentos_nfse(docs, sessao, lig)

    notas = []
    for d in docs:
        Ls = lig.get(id(d), [])
        n = Nota(d['mov'], d, Ls)
        for a in d['avisos']:
            n.aponts.append({'cod': 'AVISO', 'sev': 'info', 'titulo': a, 'detalhe': '', 'xml': '', 'sis': '', 'correcao': '', 'par': None, 'campo': '', 'valor': 0})
        if d.get('sem_protocolo') and d['tipo'] in ('nfe', 'cte'):
            n.ap('XML_SEM_PROTOCOLO', 'O XML não tem o protocolo de autorização da SEFAZ (é o arquivo da nota, não o procNFe autorizado).', 'sem protocolo', '',
                 'Conferir na SEFAZ se a nota foi autorizada e mandar o XML autorizado.')
        if not d['empresa_envolvida']:
            n.ap('OUTRA_EMPRESA', 'Nem emitente nem destinatário/tomador têm a raiz %s.' % emp['raiz'], d['part_cnpj'], '', 'Retirar o XML do lote ou conferir a empresa.')
        if not Ls:
            if d['cancelada'] or d.get('denegada'):
                n.status_extra = 'cancelada'
            elif d['mov'] in ('saida', 'prestado') and competencia and _mes(d['emissao']) != competencia:
                n.ap('FORA_COMPETENCIA', 'Emitida em %s; o fechamento é de %s.' % (_mes(d['emissao']), competencia), d['emissao'], '', 'Conferir se o XML é deste fechamento.')
            elif d['mov'] in ('entrada', 'frete', 'tomado') and competencia and _mes(d['emissao']) < competencia:
                n.ap('NAO_LANCADA_MES_ANT', 'Emitida em %s. Pode ter sido lançada no mês anterior.' % _mes(d['emissao']), d['emissao'], 'não encontrada', 'Conferir se foi lançada em %s; se não, lançar.' % _mes(d['emissao']), valor=d['valor'])
            elif d['empresa_envolvida']:
                n.ap('NAO_LANCADA', 'XML sem lançamento correspondente na planilha.', d['valor'], 'não encontrada', 'Lançar a nota no sistema.', valor=d['valor'])
            notas.append(n)
            continue
        L = Ls[0]
        if len(Ls) > 1 or len(L['blocos']) > 1:
            refs = [x['ref'] for x in Ls] + ['%s › %s › linha %s' % (L['arquivo'], L['aba'], b[0]['linha']) for b in L['blocos'][1:]]
            n.ap('DUPLICIDADE', 'A mesma nota aparece em mais de um lançamento: %s.' % ' | '.join(refs), d['valor'],
                 brl(sum(x['valor'] for x in Ls) + sum(sum(r.get('valor_contabil', 0) for r in b) for b in L['blocos'][1:])), 'Excluir o lançamento repetido.',
                 valor=sum(x['valor'] for x in Ls[1:]) + sum(sum(r.get('valor_contabil', 0) for r in b) for b in L['blocos'][1:]))
        if d.get('denegada'):
            n.ap('DENEGADA', 'A SEFAZ denegou esta nota (protocolo %s): ela não tem validade.' % d['raw'].get('protocolo'), 'denegada', L['valor'],
                 'Excluir o lançamento.', valor=L['valor'])
            n.cancel_fonte = 'denegada'
            notas.append(n)
            continue
        sit_canc = L.get('cancelada_sis') or (L['valor'] <= TOL and not any(any(r.get(c) for c in ('valor_icms', 'base_icms', 'valor_ipi')) for r in L['linhas']))
        if d['cancelada']:
            if L['valor'] > TOL and not L.get('cancelada_sis'):
                n.ap('CANCELADA_LANCADA', 'Existe evento de cancelamento para esta nota.', 'cancelada', L['valor'], 'Excluir o lançamento (ou lançar como cancelada, sem valor).', valor=L['valor'])
            else:
                n.ap('CANCELADA_OK', 'Cancelada no XML e lançada como cancelada/sem valor no sistema.')
                n.status_extra = 'cancelada'
            notas.append(n)
            continue
        if sit_canc and d['valor'] > TOL:
            n.cancel_fonte = 'sistema' if L.get('cancelada_sis') else ''
            n.ap('CANCELADA_SO_NO_SISTEMA', 'O sistema tem a nota %s, mas o XML do lote está autorizado e não veio o evento de cancelamento.%s' % (
                     ('cancelada em %s' % _dbr(L['data_cancelamento'])) if L.get('data_cancelamento') else ('como cancelada' if L.get('cancelada_sis') else 'zerada (sem valor)'),
                     ' Ela fica FORA do faturamento, seguindo o sistema.' if L.get('cancelada_sis') else ' Ela continua no faturamento até confirmar.'),
                 'autorizada ' + brl(d['valor']), 'cancelada' if L.get('cancelada_sis') else 'R$ 0,00', 'Mandar o XML do cancelamento; se a nota vale, lançar o valor.', valor=d['valor'])
            notas.append(n)
            continue
        if L.get('como') == 'numero':
            n.ap('CASOU_SO_NUMERO', 'A planilha não traz CNPJ nem chave; o vínculo foi feito só pelo número.', '', '', 'Conferir se é a mesma nota.')
        if L['chave'] and d['chave'] and L['chave'] != d['chave'] and d['tipo'] != 'nfse':
            n.ap('CHAVE', 'Chave lançada não é a do XML.', d['chave'], L['chave'], 'Corrigir a chave de acesso.')
        if L['cnpj'] and not L.get('cnpj_da_chave') and d['part_cnpj'] and digitos(L['cnpj']) != digitos(d['part_cnpj']):
            if raiz(L['cnpj']) == raiz(d['part_cnpj']):
                n.ap('PARTICIPANTE_FILIAL', 'Mesma empresa, estabelecimento diferente.', d['part_cnpj'], L['cnpj'], 'Trocar o participante para %s.' % d['part_cnpj'])
            else:
                n.ap('PARTICIPANTE', 'O participante lançado não é o do XML (%s).' % (d['part_nome'] or ''), d['part_cnpj'], L['cnpj'], 'Trocar o participante para %s.' % d['part_cnpj'])
        if L['serie'] and d['serie'] and L['serie'].lstrip('0') != d['serie'].lstrip('0'):
            n.ap('SERIE', 'Série lançada difere do XML.', d['serie'], L['serie'], 'Corrigir a série.')
        if d['mov'] in ('entrada', 'frete', 'tomado') and L['data_lanc'] and d['emissao'] and L['data_lanc'] < d['emissao']:
            n.ap('DATA_ANTERIOR', 'A entrada está datada antes da emissão da nota.', d['emissao'], L['data_lanc'], 'Corrigir a data de entrada.')
        if L['data_emissao'] and d['emissao'] and L['data_emissao'] != d['emissao']:
            n.ap('DATA_EMISSAO', 'Data de emissão lançada difere do XML.', d['emissao'], L['data_emissao'], 'Corrigir para %s.' % d['emissao'])
        cl = L.get('cols') or cols   # colunas da aba onde a nota foi lançada
        ctx_l = dict(ctx, cols=cl)
        vliq = (d.get('raw') or {}).get('vLiq') if d['tipo'] == 'nfse' else None
        if 'valor_contabil' in cl and not igual(L['valor'], d['valor']) and vliq and igual(L['valor'], vliq):
            n.ap('VALOR_TOTAL', 'Lançado pelo valor líquido %s (nota menos retenções de %s). O livro registra o valor do serviço, %s.' % (brl(L['valor']), brl(d['valor'] - vliq), brl(d['valor'])),
                 d['valor'], L['valor'], 'Lançar o valor do serviço %s e as retenções em separado.' % brl(d['valor']), valor=abs(L['valor'] - d['valor']))
        elif 'valor_contabil' in cl and not igual(L['valor'], d['valor']):
            n.ap('VALOR_TOTAL', 'Total lançado %s; XML %s. Diferença de %s.' % (brl(L['valor']), brl(d['valor']), brl(L['valor'] - d['valor'])), d['valor'], L['valor'],
                 'Ajustar o valor contábil para %s.' % brl(d['valor']), valor=abs(L['valor'] - d['valor']))
        if d['tipo'] == 'nfe':
            regras_nfe(n, d, L, ctx_l)
        elif d['tipo'] == 'cte':
            regras_cte(n, d, L, ctx_l)
        else:
            regras_nfse(n, d, L, ctx_l)
        notas.append(n)

    # devoluções x nota de origem; inutilizações e numeração própria
    painel_dev = devolucoes.analisar(notas, docs, ctx)
    inuts = [e for e in eventos if e.get('tipo') == 'inut']
    painel_num, novas = devolucoes.numeracao(notas, docs, inuts, lancs, emp, Nota, competencia)
    notas.extend(novas)

    # PDFs
    pdf_docs_ = [p for pdf in pdfs for p in pdf['docs']]
    xml_chaves = {d['chave'] for d in docs if d.get('chave')}
    lanc_sem = [L for L in lancs if not L['usado']]
    # nota de locação no livro de serviços prestados ('NF SEM SERVIÇO'): confere com o relatório de locação
    loc_val = defaultdict(float)
    for r in loc_regs:
        if r.get('numero'):
            loc_val[numero_doc(r['numero'])] += r.get('valor_contabil', 0) or 0
    resto = []
    for L in lanc_sem:
        num_l = numero_doc(L['numero'])
        if loc_val and L['mov_hint'] == 'prestado' and num_l in loc_val:
            if not igual(L['valor'], loc_val[num_l]):
                n = Nota('prestado', None, [L])
                n.ap('LOCACAO_VALOR', 'Nota de locação %s: no livro de serviços %s; no relatório de locação %s. Diferença de %s.' % (L['numero'], brl(L['valor']), brl(loc_val[num_l]), brl(L['valor'] - loc_val[num_l])),
                     loc_val[num_l], L['valor'], 'Conferir qual valor está certo no sistema do cliente.', valor=abs(L['valor'] - loc_val[num_l]))
                notas.append(n)
            continue
        resto.append(L)
    lanc_sem = resto
    pdf_usados = set()
    for L in lanc_sem:
        mov = L['mov_hint'] if L['mov_hint'] in ('entrada', 'saida', 'frete', 'prestado', 'tomado') else ('tomado' if L['mov_hint'] == 'servico' else 'indef')
        n = Nota(mov, None, [L])
        hit = None
        for k, p in enumerate(pdf_docs_):
            if (L['chave'] and p.get('chave') == L['chave']) or (
                    L['numero'] and p.get('numero') == L['numero'] and (not L['cnpj'] or L['cnpj'] in (p.get('cnpj_emit'), *p.get('cnpjs', [])) or raiz(L['cnpj']) == raiz(p.get('cnpj_emit', '')))):
                hit = (k, p); break
        if hit:
            pdf_usados.add(hit[0])
            n.pdf = hit[1]
            n.ap('SEM_XML_PDF', 'Não veio XML; a nota aparece em %s, página %s.' % (hit[1]['arquivo'], hit[1]['pagina']), 'só PDF', L['valor'], 'Pedir o XML ao cliente para conferir CFOP/CST.')
        else:
            p0 = L['linhas'][0]
            cf0, mod0 = (p0.get('cfop') or ''), norm(p0.get('modelo'))
            if mod0.startswith('ct') or mod0.startswith('57') or cf0[1:] in FRETE:
                txt, cor = 'CT-e lançado sem XML: o quadro de Frete de Entrada não tem o XML deste conhecimento.', 'Enviar os XML de CT-e no quadro Frete de Entrada.'
            elif cf0[1:] in ('251', '252', '253', '254', '255', '256', '257', '301', '302', '303', '304', '305', '306', '307'):
                txt, cor = 'Energia elétrica ou comunicação (CFOP %s): esses documentos costumam não ter XML de NF-e (NF3e/NFCom ou modelo 06/21/22).' % cf0, 'Conferir pela fatura (PDF) no quadro Documentos Diversos.'
            elif L['chave'] and L['chave'][2:6] < (competencia[2:4] + competencia[5:7] if competencia else '0000'):
                txt, cor = 'Nota de %s/%s lançada neste mês; o XML não veio no lote.' % (L['chave'][4:6], '20' + L['chave'][2:4]), 'Enviar o XML no quadro de Entradas (pode ser de mês anterior).'
            else:
                txt, cor = 'Lançamento sem XML nem PDF entre os documentos enviados.', 'Pedir o XML ou conferir se a nota existe.'
            n.ap('SEM_XML', txt, 'não encontrado', L['valor'], cor, valor=L['valor'])
        notas.append(n)
    lanc_chaves = {L['chave'] for L in lancs if L['chave']}
    lanc_nums = {(L['numero'], raiz(L['cnpj'])) for L in lancs}
    lanc_nums_sem_cnpj = {L['numero'] for L in lancs}
    for k, p in enumerate(pdf_docs_):
        if k in pdf_usados or p['tipo'] == 'nao_identificado':
            continue
        if p.get('chave') and p['chave'] in xml_chaves:
            continue
        if p.get('chave') and p['chave'] in lanc_chaves:
            continue
        cn = p.get('cnpj_emit') or ''
        if (p['numero'], raiz(cn)) in lanc_nums or (not cn and p['numero'] in lanc_nums_sem_cnpj):
            continue
        if p['tipo'] == 'nfse' and any((p['numero'], raiz(c)) in lanc_nums for c in p.get('cnpjs', [])):
            continue
        if p['tipo'] == 'nfse' and any(d['tipo'] == 'nfse' and d['numero'] == p['numero'] and d['part_cnpj'] in p.get('cnpjs', []) for d in docs):
            continue
        n = Nota('diversos', None, [])
        n.pdf = p
        n.ap('PDF_NAO_LANCADO', 'Documento encontrado em %s, página %s, sem lançamento na planilha.' % (p['arquivo'], p['pagina']),
             'PDF pág. %s' % p['pagina'], 'não encontrado', 'Lançar a nota ou confirmar que não pertence ao período.', valor=p.get('valor', 0))
        notas.append(n)

    for n in notas:
        d = n.doc
        if d and d.get('cancel_fonte') == 'duplicada':
            n.ap('NFSE_REPETIDA', 'Igual à NFS-e %s (mesmo %s, valor, data e descrição). Ficou FORA das contas como provável cancelada; a %s é a que vale.' % (d['dup_de'], 'tomador' if d['mov'] == 'prestado' else 'prestador', d['dup_de']),
                 'repetida da %s' % d['dup_de'], '', 'Confirmar no Portal; se as duas valem, use "Não está cancelada" nesta nota.')
        elif d and d.get('dup_outras'):
            n.ap('NFSE_REPETIDA', 'Igual à(s) NFS-e %s (mesmo %s, valor, data e descrição), que ficaram fora das contas como prováveis canceladas. Esta é a que vale.' % (', '.join(d['dup_outras']), 'tomador' if d['mov'] == 'prestado' else 'prestador'),
                 'vale esta', '', 'Confirmar no Portal qual foi cancelada.')
        elif d and d.get('dup_suspeita') and not d.get('cancelada'):
            n.ap('NFSE_REPETIDA', 'Igual à(s) NFS-e %s (mesmo participante, valor, data e descrição) e o sistema tem todas como válidas. Se uma foi cancelada, marque-a como cancelada.' % ', '.join(d['dup_suspeita']),
                 'repetida', '', 'Confirmar no Portal.')
        elif d and d.get('cancel_fonte') == 'manual':
            n.ap('CANCEL_MANUAL', 'Fora do faturamento, dos créditos e da Estratégia.', 'cancelada', '', 'Guardar o evento de cancelamento.')
    for n in notas:
        if n.doc and (n.doc.get('raw') or {}).get('manual'):
            n.ap('MANUAL', 'Nota de serviço digitada no quadro "Notas lançadas à mão" (sem XML).', 'manual', '', 'Guardar o documento (PDF) da nota.')
    # sem planilha nem relatório do sistema: fechamento só pelas notas (não há lançamento para comparar)
    so_notas = not lancs
    # só compara o lado (entrada/saída/prestado/tomado) que tem referência do sistema: relatório só de tomados não acusa prestados
    movs_ref = set()
    for L in lancs:
        movs_ref |= set(COMPAT.get(L['mov_hint'], COMPAT['']))
    for n in notas:
        if n.doc and (n.doc['mov'] not in movs_ref or (n.doc.get('raw') or {}).get('manual')):
            n.aponts = [x for x in n.aponts if x['cod'] not in ('NAO_LANCADA', 'NAO_LANCADA_MES_ANT')]
    for n in notas:
        consolidar.consolidar(n)
    locs = _locacoes(loc_regs)
    if so_notas:
        regs = (regs or []) + _regs_do_xml(docs, emp, ncms_vendidos)   # apuração de ICMS/IPI a partir das próprias notas
    r = montar_resultado(notas, docs, lancs, descs, emp, competencia, inventario, erros, pdfs, cols, t0, painel_dev, painel_num, inuts, regs, locs, serv_rel)
    r['modo'] = 'so_notas' if so_notas else 'comparativo'
    # ICMS/IPI só existem com NF-e/CT-e (produto e frete); prestadora de serviço só tem NFS-e
    r['tem_icms'] = any(d['tipo'] in ('nfe', 'cte') for d in docs) or any((L.get('mov_hint') in ('entrada', 'saida', 'frete')) for L in lancs)
    return r


def _fiscal(d, L):
    """Resumo fiscal da nota para os relatórios (do XML; sem XML, da planilha)."""
    if d:
        raw = d['raw']
        if d['tipo'] == 'nfe':
            t = raw['totais']
            cfops = sorted({it['CFOP'] for it in raw['itens']})
            return {'fonte': 'xml', 'cfops': cfops, 'vc': t['vNF'], 'bc': t['vBC'], 'icms': t['vICMS'], 'bc_st': t['vBCST'], 'st': t['vST'], 'ipi': t['vIPI'],
                    'pis': t['vPIS'], 'cofins': t['vCOFINS'], 'difal_dest': t.get('vICMSUFDest', 0), 'fcp_dest': t.get('vFCPUFDest', 0), 'difal_remet': t.get('vICMSUFRemet', 0),
                    'idDest': raw.get('idDest'), 'indFinal': raw.get('indFinal'), 'indIEDest': raw['dest'].get('indIEDest'), 'uf_dest': raw['dest'].get('uf'),
                    'uf_emit': raw['emit'].get('uf'), 'crt': raw['emit'].get('crt'), 'natOp': raw.get('natOp', ''),
                    'tem_grupo_difal': any(it.get('ufdest') for it in raw['itens']),
                    'pis_bc': round(sum((it.get('pis') or {}).get('vBC', 0) for it in raw['itens']), 2)}
        if d['tipo'] == 'cte':
            ic = raw.get('icms') or {}
            return {'fonte': 'xml', 'cfops': [raw.get('CFOP', '')], 'vc': raw['vTPrest'], 'bc': ic.get('vBC', 0), 'icms': ic.get('vICMS', 0), 'st': 0, 'ipi': 0,
                    'uf_ini': raw.get('UFIni'), 'uf_fim': raw.get('UFFim'), 'tomador': raw.get('tomador_cnpj')}
        return {'fonte': 'xml', 'cfops': [], 'vc': raw['vServ'], 'iss': raw['vISS'], 'aliq_iss': raw['pAliq'], 'iss_retido': bool(raw['issRetido']),
                'ret': raw['ret'], 'vliq': raw.get('vLiq', 0), 'servico': raw.get('cTribNac', ''), 'desc_serv': (raw.get('xDescServ') or '')[:200]}
    if L:
        ls = L['linhas']
        sm = lambda k: round(sum(r.get(k, 0) or 0 for r in ls if isinstance(r.get(k, 0), (int, float))), 2)
        return {'fonte': 'planilha', 'cfops': sorted({r.get('cfop', '') for r in ls if r.get('cfop')}), 'vc': L['valor'], 'bc': sm('base_icms'), 'icms': sm('valor_icms'),
                'st': sm('valor_st'), 'ipi': sm('valor_ipi'), 'iss': sm('valor_iss'),
                'ret': {'irrf': sm('irrf'), 'csll': sm('csll_ret'), 'inss': sm('inss_ret'), 'pis': sm('pis_ret'), 'cofins': sm('cofins_ret')}}
    return {}


def _nota_json(i, n):
    d, L = n.doc, (n.lancs[0] if n.lancs else None)
    pdf = getattr(n, 'pdf', None)
    erros = [a for a in n.aponts if a['sev'] == 'erro']
    alertas = [a for a in n.aponts if a['sev'] == 'alerta']
    status = 'erro' if erros else ('alerta' if alertas else ('cancelada' if getattr(n, 'status_extra', '') == 'cancelada' else 'ok'))
    if d:
        ident = {'tipo': {'nfe': 'NF-e' if d.get('modelo') != '65' else 'NFC-e', 'cte': 'CT-e', 'nfse': 'NFS-e'}[d['tipo']], 'numero': d['numero'],
                 'serie': d['serie'], 'chave': d['chave'], 'emissao': d['emissao'], 'part_cnpj': d['part_cnpj'], 'part_nome': d['part_nome'],
                 'part_uf': d['part_uf'], 'valor_xml': d['valor'], 'arquivo': d['arquivo'], 'cancelada': d['cancelada']}
    elif pdf:
        ident = {'tipo': {'nfe': 'NF-e', 'nfce': 'NFC-e', 'cte': 'CT-e', 'nfse': 'NFS-e'}.get(pdf['tipo'], 'Doc.'), 'numero': pdf.get('numero', ''),
                 'serie': pdf.get('serie', ''), 'chave': pdf.get('chave', ''), 'emissao': ('20' + pdf['aamm'][:2] + '-' + pdf['aamm'][2:]) if pdf.get('aamm') else '',
                 'part_cnpj': pdf.get('cnpj_emit') or (pdf.get('cnpjs') or [''])[0], 'part_nome': '', 'part_uf': '', 'valor_xml': pdf.get('valor', 0) or None,
                 'arquivo': '%s (pág. %s)' % (pdf['arquivo'], pdf['pagina']), 'cancelada': False}
    elif not L:
        ident = {'tipo': 'Documento', 'numero': '', 'serie': '', 'chave': '', 'emissao': '', 'part_cnpj': '', 'part_nome': '', 'part_uf': '',
                 'valor_xml': None, 'arquivo': '', 'cancelada': False}
    else:
        ident = {'tipo': 'Lançamento', 'numero': L['numero'], 'serie': L['serie'], 'chave': L['chave'], 'emissao': L['data_emissao'],
                 'part_cnpj': L['cnpj'], 'part_nome': L['participante'], 'part_uf': '', 'valor_xml': None, 'arquivo': '', 'cancelada': False}
    if L and not ident['part_nome']:
        ident['part_nome'] = L['participante']
    nat = 'normal'
    fonte = ''
    if d:
        if d.get('cancelada'):
            nat, fonte = 'cancelada', d.get('cancel_fonte') or 'xml'
        elif d.get('denegada'):
            nat, fonte = 'cancelada', 'denegada'
        elif getattr(n, 'cancel_fonte', '') == 'sistema':
            nat, fonte = 'cancelada', 'sistema'
        else:
            nat = d.get('natureza', 'normal')
    elif any(a['cod'] == 'INUTILIZADA_LANCADA' for a in n.aponts):
        nat = 'inutilizada'
    elif getattr(n, 'numeracao', None):
        nat = 'numeracao'
        nm = n.numeracao
        ident.update({'tipo': 'Numeração', 'numero': str(nm['ini']) if nm['ini'] == nm['fim'] else '%s a %s' % (nm['ini'], nm['fim']), 'serie': nm['serie']})
    j = {'id': i, 'mov': n.mov, 'status': status, 'natureza': nat, 'cancel_fonte': fonte,
         'cancelada_efetiva': nat == 'cancelada', 'natureza_rot': devolucoes.ROTULO.get(nat, 'Numeração'), **ident,
         'valor_sis': L['valor'] if L else None, 'data_lanc': L['data_lanc'] if L else '', 'ref_planilha': L['ref'] if L else '',
         'como': L.get('como', '') if L else '', 'pdf': ('%s, pág. %s' % (pdf['arquivo'], pdf['pagina'])) if pdf else '',
         'aponts': n.aponts, 'pares': n.pares, 'n_erros': len(erros), 'n_alertas': len(alertas),
         'impacto': round(sum(a['valor'] for a in n.aponts if a['sev'] in ('erro', 'alerta')), 2)}
    # chave estável da nota e de cada apontamento: o "Corrigido" sobrevive a um novo cálculo
    nk = j['chave'] or '%s|%s|%s|%s|%s' % (n.mov, j['tipo'], j['numero'], j['part_cnpj'], j['ref_planilha'] if not d else '')
    j['k'] = nk
    for a in n.aponts:
        a['k'] = '%s|%s|%s' % (nk, a['cod'], ','.join(str(p) for p in a.get('pares') or ([] if a['par'] is None else [a['par']])))
    j['fiscal'] = _fiscal(d, L)
    tipo_ck = d['tipo'] if d else ''
    j['checklist'] = consolidar.checklist(n, (L.get('cols') if L else None) or [], tipo_ck, bool(L), nat) if d else []
    return j


def chave_doc(d):
    return d.get('chave') or d.get('arquivo') or ''


def _cancelamentos_nfse(docs, sessao, lig=None):
    """A NFS-e nacional não muda quando é cancelada (o cancelamento é um evento à parte). Sem o evento:
    - NFS-e repetida (mesmo participante, valor, data e descrição) = cancelada e reemitida: a de número menor fica fora;
    - o que o usuário marcar na tela (sessao/canceladas.json: {chave: true/false}) vence a regra acima."""
    import json
    p = os.path.join(sessao, 'canceladas.json')
    try:
        marc = json.load(open(p, encoding='utf-8')) if os.path.exists(p) else {}
    except Exception:
        marc = {}
    grupos = defaultdict(list)
    for d in docs:
        if d['tipo'] != 'nfse' or d.get('cancelada'):
            continue
        desc = norm((d.get('raw') or {}).get('xDescServ') or '')[:300]
        grupos[(d['mov'], digitos(d.get('part_cnpj')), round(d.get('valor') or 0, 2), desc, d.get('emissao'))].append(d)
    lig = lig or {}
    canc_sis = lambda d: any(L.get('cancelada_sis') for L in lig.get(id(d), []))
    for g in grupos.values():
        if len(g) < 2:
            continue
        if any(canc_sis(d) for d in g):
            continue  # o sistema/Portal já diz qual das repetidas foi cancelada: vale ele, não a regra
        g.sort(key=lambda x: int(x['numero']) if str(x['numero']).isdigit() else 0)
        if all(lig.get(id(d)) for d in g):
            for d in g:   # o sistema tem todas como válidas: só avisa, quem decide é o usuário
                d['dup_suspeita'] = [x['numero'] for x in g if x is not d]
            continue
        vale = g[-1]
        for d in g[:-1]:
            d['cancelada'] = True; d['cancel_fonte'] = 'duplicada'; d['dup_de'] = vale['numero']
        vale['dup_outras'] = [d['numero'] for d in g[:-1]]
    for d in docs:
        m = marc.get(chave_doc(d))
        if m is True and not d.get('cancelada'):
            d['cancelada'] = True; d['cancel_fonte'] = 'manual'
        elif m is True and d.get('cancel_fonte') == 'duplicada':
            d['cancel_fonte'] = 'manual'
        elif m is False and d.get('cancel_fonte') in ('duplicada', 'manual'):
            d['cancelada'] = False; d['cancel_fonte'] = ''; d.pop('dup_de', None)


def manuais_raw(sessao, emp):
    """Notas de serviço digitadas na tela (sessao/manuais.json) no mesmo formato da NFS-e lida do XML."""
    import json
    p = os.path.join(sessao, 'manuais.json')
    if not os.path.exists(p):
        return []
    try:
        lista = json.load(open(p, encoding='utf-8')) or []
    except Exception:
        return []
    out = []
    for i, m in enumerate(lista, start=1):
        valor = float(m.get('valor') or 0)
        if valor <= 0:
            continue
        aliq = float(m.get('aliq_iss') or 0)
        iss = float(m['iss']) if m.get('iss') not in (None, '') else round(valor * aliq / 100, 2)
        ret = {k: float(m.get(k) or 0) for k in ('irrf', 'csll', 'inss', 'pis', 'cofins')}
        part = {'cnpj': digitos(m.get('cnpj')), 'nome': (m.get('nome') or '').strip(), 'uf': (m.get('uf') or '').upper()}
        empresa = {'cnpj': emp.get('cnpj') or '', 'nome': emp.get('nome') or '', 'uf': emp.get('uf') or ''}
        tomado = m.get('tipo', 'tomado') == 'tomado'
        out.append({'tipo': 'nfse', 'leiaute': 'manual', 'arquivo': 'manual #%d' % i, 'chave': '', 'numero': str(m.get('numero') or '').strip(),
                    'emissao': m.get('data') or '', 'competencia': m.get('data') or '',
                    'prestador': part if tomado else empresa, 'tomador': empresa if tomado else part,
                    'cTribNac': digitos(m.get('servico')), 'xDescServ': m.get('descricao') or '', 'vServ': valor, 'vISS': iss, 'pAliq': aliq,
                    'vLiq': round(valor - sum(ret.values()) - (iss if m.get('iss_retido') else 0), 2), 'vBC': valor, 'issRetido': bool(m.get('iss_retido')),
                    'ret': ret, 'pisCofins': {'CST': '', 'vPis': ret['pis'], 'vCofins': ret['cofins']}, 'cancelada': False,
                    'card': 'tomado' if tomado else 'prestado', 'manual': True, 'id_manual': m.get('id')})
    return out


def _regs_do_xml(docs, emp, ncms_vendidos):
    """Sem planilha do sistema: monta as linhas do livro (CFOP, CST, valores) direto das NF-e e CT-e.
    Nota própria usa o CFOP dela; nota de terceiro vira CFOP de entrada (1º dígito pela UF do emitente × UF da empresa,
    uso e consumo/ativo/combustível pelo item, sem crédito de ICMS em uso e consumo e ativo)."""
    uf_emp = (emp.get('uf') or '').upper()
    # indústria (vende produção própria): compra de mercadoria entra como insumo (x101/x401), não revenda (x102/x403)
    industria = any(d['tipo'] == 'nfe' and d.get('propria') and d['mov'] == 'saida' and not d.get('cancelada')
                    and any((it.get('CFOP') or '')[1:] in ('101', '116', '122', '124', '125') for it in d['raw'].get('itens', [])) for d in docs)
    out, i = [], 0
    def pri(uf):
        uf = (uf or '').upper()
        return '3' if uf == 'EX' else ('1' if (uf == uf_emp or not uf) else '2')
    for d in docs:
        if d.get('cancelada') or d.get('natureza') in ('cancelada', 'inutilizada') or d['tipo'] not in ('nfe', 'cte'):
            continue
        raw = d['raw']
        if d['tipo'] == 'cte':
            if d['mov'] != 'frete':
                continue
            ic = raw.get('icms') or {}
            cf = raw.get('CFOP') or ''
            if len(cf) == 4:
                cf = pri(d.get('part_uf')) + cfop_entrada_esperado(cf)[1:]
            i += 1
            out.append({'arquivo': 'XML', 'aba': 'xml', 'linha': i, 'cfop': cf, 'cst_icms': ic.get('CST', ''), 'valor_contabil': raw.get('vTPrest', 0),
                        'base_icms': ic.get('vBC', 0), 'valor_icms': ic.get('vICMS', 0), 'valor_ipi': 0.0, 'valor_st': 0.0, 'base_ipi': 0.0,
                        'numero': d['numero'], 'chave': d.get('chave', ''), 'cnpj': d.get('part_cnpj', ''), 'origem': 'xml'})
            continue
        for it in raw.get('itens', []):
            ic, ip = it.get('icms') or {}, it.get('ipi') or {}
            cst = (ic.get('orig', '') + ic.get('CST', '')) if len(ic.get('CST', '')) == 2 else ic.get('CST', '')
            st = ic.get('vICMSST', 0) + ic.get('vFCPST', 0)
            vc = it['vProd'] - it['vDesc'] + it['vFrete'] + it['vSeg'] + it['vOutro'] + ip.get('vIPI', 0) + st
            bc, icms, cf = ic.get('vBC', 0), ic.get('vICMS', 0), it.get('CFOP') or ''
            if not d.get('propria') and len(cf) == 4:
                tem_st = st > 0 or ic.get('CST', '')[-2:] in ('10', '30', '60', '70') or ic.get('CST', '') in ('201', '202', '203', '500')
                suf = cfop_entrada_esperado(cf, tem_st)[1:]
                c = destino.classificar(it, ncms_vendidos)
                if c and c[0] == 'consumo':
                    suf = '407' if tem_st else '556'; bc = icms = 0.0
                elif c and c[0] == 'ativo':
                    suf = '406' if tem_st else '551'; bc = icms = 0.0   # crédito do ativo é pelo CIAP (1/48), fora do mês
                elif c and c[0] == 'combustivel':
                    suf = '653'
                elif industria and suf in ('102', '403'):
                    suf = {'102': '101', '403': '401'}[suf]
                cf = pri(d.get('part_uf')) + suf
            i += 1
            out.append({'arquivo': 'XML', 'aba': 'xml', 'linha': i, 'cfop': cf, 'cst_icms': cst, 'valor_contabil': round(vc, 2), 'base_icms': bc, 'valor_icms': icms,
                        'valor_ipi': ip.get('vIPI', 0), 'base_ipi': ip.get('vBC', 0), 'valor_st': st, 'numero': d['numero'], 'chave': d.get('chave', ''),
                        'cnpj': d.get('part_cnpj', ''), 'origem': 'xml'})
    return out


def montar_resultado(notas, docs, lancs, descs, emp, competencia, inventario, erros, pdfs, cols, t0, painel_dev=None, painel_num=None, inuts=None, regs=None, locs=None, serv_rel=None):
    for i, n in enumerate(notas, start=1):
        n._id = i
    cols_aba = {(x['arquivo'], x['aba']): set(x.get('mapa', {})) for x in descs if x.get('utilizavel')}
    cols_aba[('XML', 'xml')] = {'cst_icms', 'base_ipi', 'valor_contabil', 'base_icms', 'valor_icms', 'valor_ipi', 'valor_st'}
    apur = apuracao.montar(notas, [r for r in (regs or []) if r.get('origem') != 'servicos'], emp, cols_aba)
    out_notas = [_nota_json(i, n) for i, n in enumerate(notas, start=1)]
    por_id = {x['id']: x for x in out_notas}
    devs = []
    for reg in painel_dev or []:
        x = por_id[reg['n']._id]
        devs.append({'id': x['id'], 'natureza': reg['natureza'], 'rotulo': reg['rotulo'], 'numero': reg['numero'], 'emissao': reg['emissao'],
                     'part_nome': reg['part_nome'], 'valor': reg['valor'], 'valor_sis': x['valor_sis'], 'originais': reg['originais'], 'pct': reg['pct'],
                     'erros': len([a for a in x['aponts'] if a['cod'].startswith('DEV') and a['sev'] == 'erro']),
                     'alertas': len([a for a in x['aponts'] if a['cod'].startswith('DEV') and a['sev'] == 'alerta']), 'status': x['status']})
    def soma(f, campo):
        return round(sum((x[campo] or 0) for x in out_notas if f(x)), 2)
    vendas = lambda x: x['mov'] == 'saida' and x['natureza'] == 'normal' and x['tipo'] in ('NF-e', 'NFC-e')
    dvv = lambda x: x['natureza'] == 'dev_venda'
    dvc = lambda x: x['natureza'] in ('dev_compra', 'dev_compra_forn')
    compras = lambda x: x['mov'] == 'entrada' and x['natureza'] == 'normal' and x['tipo'] == 'NF-e'
    balanco = {'vendas': {'xml': soma(vendas, 'valor_xml'), 'sis': soma(vendas, 'valor_sis')},
               'dev_venda': {'xml': soma(dvv, 'valor_xml'), 'sis': soma(dvv, 'valor_sis'), 'qtd': len([x for x in out_notas if dvv(x)])},
               'compras': {'xml': soma(compras, 'valor_xml'), 'sis': soma(compras, 'valor_sis')},
               'dev_compra': {'xml': soma(dvc, 'valor_xml'), 'sis': soma(dvc, 'valor_sis'), 'qtd': len([x for x in out_notas if dvc(x)])}}
    canceladas = [{'id': x['id'], 'tipo': x['tipo'], 'numero': x['numero'], 'serie': x['serie'], 'emissao': x['emissao'], 'part_nome': x['part_nome'],
                   'valor': x['valor_xml'], 'valor_sis': x['valor_sis'], 'mov': x['mov'], 'status': x['status'], 'fonte': x['cancel_fonte'] or 'a confirmar',
                   'confirmada': x['cancel_fonte'] in ('evento', 'protocolo', 'xml'), 'servico': x['tipo'] == 'NFS-e'}
                  for x in out_notas if x['cancelada_efetiva'] or any(a['cod'] == 'CANCELADA_SO_NO_SISTEMA' for a in x['aponts'])]
    movs = {}
    for m in ['entrada', 'saida', 'frete', 'prestado', 'tomado', 'diversos', 'indef']:
        ns = [x for x in out_notas if x['mov'] == m]
        if not ns and m in ('diversos', 'indef'):
            continue
        xml_ok = [x for x in ns if x['valor_xml'] is not None and not x['cancelada'] and x['tipo'] != 'Lançamento' and not x['pdf']]
        movs[m] = {'nome': MOV_NOME[m], 'total': len(ns),
                   'xml': len([x for x in ns if x['tipo'] != 'Lançamento' and not x['pdf']]),
                   'lancadas': len([x for x in ns if x['valor_sis'] is not None]),
                   'ok': len([x for x in ns if x['status'] in ('ok', 'cancelada')]),
                   'erro': len([x for x in ns if x['status'] == 'erro']), 'alerta': len([x for x in ns if x['status'] == 'alerta']),
                   'nao_lancadas': len([x for x in ns if any(a['cod'] in ('NAO_LANCADA', 'PDF_NAO_LANCADO') for a in x['aponts'])]),
                   'sem_xml': len([x for x in ns if any(a['cod'] in ('SEM_XML', 'SEM_XML_PDF') for a in x['aponts'])]),
                   'valor_xml': round(sum(x['valor_xml'] or 0 for x in xml_ok), 2),
                   'valor_sis': round(sum(x['valor_sis'] or 0 for x in ns if x['valor_sis'] is not None), 2)}
    tipos = Counter()
    tipos_val = defaultdict(float)
    tipos_mov = defaultdict(Counter)
    for x in out_notas:
        for a in x['aponts']:
            if a['sev'] in ('erro', 'alerta'):
                tipos[a['cod']] += 1
                tipos_val[a['cod']] += a['valor']
                tipos_mov[a['cod']][x['mov']] += 1
    ranking = [{'cod': c, 'sev': CATALOGO[c][0], 'titulo': next((a['titulo'] for x in out_notas for a in x['aponts'] if a['cod'] == c), CATALOGO[c][1]),
                'categoria': consolidar.CATEGORIA.get(c, 'valores'), 'qtd': q, 'valor': round(tipos_val[c], 2),
                'notas': len([x for x in out_notas if any(a['cod'] == c for a in x['aponts'])]), 'movs': dict(tipos_mov[c])}
               for c, q in tipos.items()]
    ranking.sort(key=lambda r: (r['sev'] != 'erro', -r['notas']))
    res = {
        'empresa': emp, 'competencia': competencia, 'gerado_em': datetime.datetime.now().strftime('%d/%m/%Y %H:%M'),
        'tempo': round((datetime.datetime.now() - t0).total_seconds(), 1),
        'resumo': {'notas': len(out_notas), 'com_erro': len([x for x in out_notas if x['status'] == 'erro']),
                   'com_alerta': len([x for x in out_notas if x['status'] == 'alerta']),
                   'ok': len([x for x in out_notas if x['status'] in ('ok', 'cancelada')]),
                   'apontamentos': sum(tipos.values()), 'impacto': round(sum(tipos_val[c] for c in tipos if CATALOGO[c][0] == 'erro'), 2),
                   'xml': len(docs), 'lancamentos': len(lancs)},
        'movimentos': movs, 'ranking': ranking, 'notas': out_notas,
        'devolucoes': devs, 'balanco': balanco, 'canceladas': canceladas, 'categorias': consolidar.CATEGORIAS, 'apuracao': apur, 'painel': painel.montar(notas, emp, locs), 'locacoes': locs, 'servicos_relatorio': serv_rel,
        'inutilizadas': (painel_num or {}).get('inutilizadas', []), 'pulos': (painel_num or {}).get('pulos', []),
        'planilha': [{k: v for k, v in dsc.items() if k != 'amostra'} for dsc in descs],
        'colunas': sorted(cols), 'colunas_rotulo': {c: planilha.CAMPO_ROTULO.get(c, c) for c in cols},
        'colunas_faltando': [planilha.CAMPO_ROTULO[c] for c in ('cnpj', 'cfop', 'cst_icms', 'valor_contabil', 'valor_icms', 'cst_pis') if c not in cols],
        'inventario': inventario, 'erros_leitura': erros,
        'pdf_sem_texto': [{'arquivo': p['arquivo'], 'paginas': p['sem_texto']} for p in pdfs if p['sem_texto']],
    }
    return res
