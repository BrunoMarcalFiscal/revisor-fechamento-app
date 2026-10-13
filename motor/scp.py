# -*- coding: utf-8 -*-
"""Estratégia tributária Matriz (Lucro Real) × SCP (Lucro Presumido).

Lógica (a mesma da planilha de fechamento do escritório, com as correções legais sinalizadas):
- Matriz fica com o faturamento de produtos e com as entradas que geram crédito de PIS/COFINS (não cumulativo).
- Parte das receitas de serviço e locação fica na Matriz só até zerar (ou quase) o PIS/COFINS: entradas = saídas.
- O excedente de serviço e locação vai para a SCP (cumulativo 0,65% e 3%; IRPJ/CSLL pela presunção de 32%).
- Produto nunca vai para a SCP (premissa: a SCP não tem estoque nem opera mercadoria).

extrair(res) monta os dados a partir do fechamento; calcular(dados, cfg) faz a conta com memória de cálculo,
sem alterar os dados originais. Os parâmetros da simulação ficam em cfg (sessao/scp.json)."""
import re
from collections import defaultdict

PIS_NC, COF_NC = 1.65, 7.6
PIS_CU, COF_CU = 0.65, 3.0
VENDA = {'101', '102', '103', '104', '105', '106', '107', '108', '109', '110', '111', '112', '113', '114', '115', '116', '117', '118', '119', '120',
         '122', '123', '124', '125', '401', '402', '403', '404', '405', '651', '652', '653', '654', '655', '656', '667'}
DEV_VENDA = {'201', '202', '410', '411'}
DEV_COMPRA = {'201', '202', '410', '411'}
CRED_PADRAO = {'101', '102', '111', '113', '116', '117', '118', '120', '121', '122', '124', '125', '126', '401', '403',  # insumo e revenda
               '251', '252', '253', '254', '255', '256', '257',  # energia elétrica
               '351', '352', '353', '354', '355', '356', '360'}  # frete na compra
CRED_NOME = {**{s: 'Insumo / industrialização' for s in ('101', '111', '116', '120', '122', '124', '125', '126', '401')},
             **{s: 'Mercadoria para revenda' for s in ('102', '113', '117', '118', '121', '403')},
             **{s: 'Energia elétrica' for s in ('251', '252', '253', '254', '255', '256', '257')},
             **{s: 'Frete' for s in ('351', '352', '353', '354', '355', '356', '360')}}


def _r(v):
    return round(v + 1e-9, 2)


def _b(v):
    s = '{:,.2f}'.format(v or 0).replace(',', 'X').replace('.', ',').replace('X', '.')
    return s


# ------------------------------------------------------------------------------------------------ dados do fechamento

def extrair(res):
    ap = res.get('apuracao') or {}
    def _nfs(l):
        ns = {n for cs in l.get('csts', []) for n in cs.get('notas', [])}
        return sorted(ns, key=lambda x: int(x) if str(x).isdigit() else 0)
    linhas = [{'cfop': l['cfop'], 'desc': l.get('desc', ''), 'vc': l['vc'], 'icms': l['icms'], 'ipi': l['ipi'], 'st': l['st'], 'notas': _nfs(l)} for l in ap.get('linhas', [])]
    sr = res.get('servicos_relatorio') or {}
    notas = [n for n in res.get('notas', []) if n.get('natureza') not in ('cancelada', 'numeracao', 'inutilizada')]
    emp_raiz = (res.get('empresa') or {}).get('raiz', '')
    # serviços prestados: relatório de saída de serviços; sem ele, as NFS-e prestadas do lote
    def _do_xml(n, mov):
        f = n.get('fiscal') or {}; r = f.get('ret') or {}
        manual = (n.get('arquivo') or '').startswith('manual')
        x = {'numero': n['numero'], 'data': n['emissao'], 'participante': n.get('part_nome'), 'cnpj': n.get('part_cnpj') or '', 'valor': n['valor_xml'] or 0,
             'servico': f.get('servico', ''), 'fonte': 'lançada à mão' if manual else ('XML de NFS-e prestada' if mov == 'prestado' else 'XML de NFS-e tomada')}
        if mov == 'prestado':
            x.update({'iss': f.get('iss', 0), 'iss_retido': f.get('iss_retido'), 'irrf': r.get('irrf', 0), 'csll': r.get('csll', 0), 'pis': r.get('pis', 0),
                      'cofins': r.get('cofins', 0), 'inss': r.get('inss', 0), 'descricao': f.get('desc_serv', ''), 'emitente': (n.get('chave') or '')})
        else:
            x['pf'] = len(n.get('part_cnpj') or '') == 11
        return x
    def _juntar(lado, mov, rotulo):
        """Relatório do sistema/Portal é a base; soma as notas lançadas à mão e os XML que não estão no relatório (sem repetir)."""
        base = [dict(x, fonte=rotulo) for x in sr.get(lado) or []]
        xml = [n for n in notas if n['tipo'] == 'NFS-e' and n['mov'] == mov]
        if not base:
            return [_do_xml(n, mov) for n in xml]
        tem = {(str(x.get('numero') or '').lstrip('0'), re.sub(r'\D', '', x.get('cnpj') or '')[:8]) for x in base}
        nums = {k[0] for k in tem}
        for n in xml:
            num, cn = str(n['numero'] or '').lstrip('0'), re.sub(r'\D', '', n.get('part_cnpj') or '')[:8]
            if (num, cn) in tem or (not cn and num in nums):
                continue
            base.append(_do_xml(n, mov))
        return base
    prest = _juntar('saida', 'prestado', 'relatório de saída de serviços')
    toma = _juntar('entrada', 'tomado', 'relatório de entrada de serviços')
    loc = (res.get('locacoes') or {}).get('linhas', [])
    # importação: PIS/COFINS-Importação pago, pelas NF-e de entrada de importação (CFOP 3xxx) do lote
    imp_pis = imp_cof = imp_vc = 0.0
    for n in notas:
        f = n.get('fiscal') or {}
        if n['tipo'] == 'NF-e' and any(c.startswith('3') for c in f.get('cfops', [])):
            imp_pis += f.get('pis', 0) or 0; imp_cof += f.get('cofins', 0) or 0; imp_vc += f.get('vc', 0) or 0
    com = sem = 0; icms_dentro = 0.0
    for n in notas:
        f = n.get('fiscal') or {}
        if n['tipo'] == 'NF-e' and n['mov'] == 'saida' and f.get('pis_bc'):
            vc_sem = (f.get('vc') or 0) - (f.get('ipi') or 0) - (f.get('st') or 0)
            if f.get('icms') and abs(f['pis_bc'] - vc_sem) <= max(1.0, 0.002 * vc_sem):
                com += 1; icms_dentro += f['icms']
            else:
                sem += 1
    return {'icms_pis_saidas': {'com_icms': com, 'sem_icms': sem, 'icms_dentro': round(icms_dentro, 2)},
            'empresa': res.get('empresa', {}), 'competencia': res.get('competencia'), 'cfops': linhas, 'prestados': prest, 'tomados': toma,
            'locacoes': loc, 'importacao_xml': {'pis': _r(imp_pis), 'cofins': _r(imp_cof), 'vc': _r(imp_vc)},
            'credita_ipi': (ap.get('totais') or {}).get('entrada', {}).get('ipi', 0) > 0, 'emp_raiz': emp_raiz}


def cfg_padrao(dados):
    return {'scp_cnpj': '', 'scp_nome': '', 'matriz_nome': '', 'modo': 'auto', 'prioridade': 'locacao', 'pct_serv_matriz': 0.0, 'pct_loc_matriz': 0.0,
            'creditos_cfop': {}, 'excluir_ipi_credito': bool(dados.get('credita_ipi')), 'excluir_ipi_debito': True, 'excluir_st_debito': True,
            'deduzir_devolucoes': True, 'importacao': 'pago' if (dados.get('importacao_xml') or {}).get('pis') else 'base',
            'imp_pis_manual': None, 'imp_cof_manual': None, 'tomados_base_manual': None, 'tomados_excluir': [], 'excluir_pf': True,
            'outros_creditos': [], 'saldo_pis_ant': 0.0, 'saldo_cof_ant': 0.0, 'serv_manual': None, 'loc_manual': None,
            'ret_pis_manual': None, 'ret_cof_manual': None, 'ret_irrf_manual': None, 'ret_csll_manual': None,
            'tri_outros_meses': [], 'retidas_scp': False, 'irpj_pago_tri': 0.0, 'csll_pago_tri': 0.0, 'lucro_matriz': 0.0, 'margem_serv_real': 0.0,
            'saida_cfop_extra': {}, 'ant_faturamento': None, 'ant_pis': None, 'ant_cofins': None, 'ant_irpj': None, 'ant_csll': None, 'ant_ipi': None,
            'ant_iss': None, 'ant_icms': None, 'ant_carga_pct': 21.69, 'compensa': {'pis': True, 'cofins': True, 'icms': True, 'retencao': True}}


# ------------------------------------------------------------------------------------------------ cálculo

def calcular(dados, cfg):
    c = dict(cfg_padrao(dados), **(cfg or {}))
    mem = defaultdict(list)       # memória de cálculo por seção: [rótulo, valor, fórmula]
    alertas = []
    def m(sec, rot, val, form='', tipo=''):
        mem[sec].append({'rot': rot, 'valor': _r(val), 'formula': form, 'tipo': tipo})
        return _r(val)

    # ---- 1. produtos: débito (saídas) e créditos (entradas), linha a linha de CFOP
    deb_lin, cred_lin, fora_lin, deb_fora = [], [], [], []
    deb_ov = c.get('debito_cfop') or {}   # ADM: CFOP de saída/devolução que entra (True) ou sai (False) da base de débito
    for l in dados['cfops']:
        cf, suf = l['cfop'], l['cfop'][1:]
        ent = cf[0] in '123'
        nfs = l.get('notas') or []
        ov = deb_ov.get(cf)
        if not ent and ov is False or (not ent and ov is None and not (suf in VENDA or c['saida_cfop_extra'].get(cf)) and not (suf in DEV_COMPRA and c['deduzir_devolucoes'])):
            deb_fora.append({'cfop': cf, 'desc': l['desc'], 'vc': l['vc'], 'icms': l['icms'], 'ipi': l['ipi'], 'st': l['st'], 'base': _r(l['vc'] - l['icms']), 'sinal': 1, 'notas': nfs, 'padrao': suf in VENDA})
            continue
        if ent and suf in DEV_VENDA and (ov is False or (ov is None and not c['deduzir_devolucoes'])):
            deb_fora.append({'cfop': cf, 'desc': l['desc'] + ' (devolução de venda)', 'vc': l['vc'], 'icms': l['icms'], 'ipi': l['ipi'], 'st': l['st'], 'base': _r(-(l['vc'] - l['icms'])), 'sinal': -1, 'notas': nfs, 'padrao': True})
            continue
        if not ent and (ov or suf in VENDA or c['saida_cfop_extra'].get(cf)):
            exc_ipi = l['ipi'] if c['excluir_ipi_debito'] else 0
            exc_st = l['st'] if c['excluir_st_debito'] else 0
            base = l['vc'] - l['icms'] - exc_ipi - exc_st
            deb_lin.append({'cfop': cf, 'desc': l['desc'], 'vc': l['vc'], 'icms': l['icms'], 'ipi': exc_ipi, 'st': exc_st, 'base': _r(base), 'sinal': 1, 'notas': nfs, 'padrao': suf in VENDA})
        elif ent and suf in DEV_VENDA:
            base = l['vc'] - l['icms'] - (l['ipi'] if c['excluir_ipi_debito'] else 0) - (l['st'] if c['excluir_st_debito'] else 0)
            deb_lin.append({'cfop': cf, 'desc': l['desc'] + ' (devolução, reduz o débito)', 'vc': l['vc'], 'icms': l['icms'], 'ipi': l['ipi'] if c['excluir_ipi_debito'] else 0,
                            'st': l['st'] if c['excluir_st_debito'] else 0, 'base': _r(-base), 'sinal': -1, 'notas': nfs, 'padrao': True})
        elif ent and cf[0] in '12':
            padrao = suf in CRED_PADRAO
            inclui = c['creditos_cfop'].get(cf, padrao)
            exc_ipi = l['ipi'] if c['excluir_ipi_credito'] else 0
            base = l['vc'] - l['icms'] - exc_ipi
            reg = {'cfop': cf, 'desc': l['desc'], 'grupo': CRED_NOME.get(suf, 'Outras entradas'), 'vc': l['vc'], 'icms': l['icms'], 'ipi': exc_ipi, 'base': _r(base),
                   'padrao': padrao, 'inclui': inclui, 'notas': nfs}
            (cred_lin if inclui else fora_lin).append(reg)
        elif not ent and suf in DEV_COMPRA and c['deduzir_devolucoes']:
            base = l['vc'] - l['icms'] - (l['ipi'] if c['excluir_ipi_credito'] else 0)
            cred_lin.append({'cfop': cf, 'desc': l['desc'] + ' (devolução de compra, reduz o crédito)', 'grupo': 'Devolução de compra', 'vc': l['vc'], 'icms': l['icms'],
                             'ipi': l['ipi'] if c['excluir_ipi_credito'] else 0, 'base': _r(-base), 'padrao': True, 'inclui': True, 'notas': nfs})
    base_prod = _r(sum(x['base'] for x in deb_lin))
    ed = []  # valores digitados pelo ADM (sobrepõem o calculado)
    if c.get('base_prod_manual') is not None or c.get('devol_manual') is not None:
        vend = c['base_prod_manual'] if c.get('base_prod_manual') is not None else sum(x['base'] for x in deb_lin if x['sinal'] > 0)
        dvl = c['devol_manual'] if c.get('devol_manual') is not None else -sum(x['base'] for x in deb_lin if x['sinal'] < 0)
        base_prod = _r(vend - dvl)
        ed.append('base de produtos')
    m('produtos', 'Saídas de produtos (vendas, sem 5922)', sum(x['vc'] for x in deb_lin if x['sinal'] > 0), 'soma do valor contábil dos CFOP de venda')
    m('produtos', '(−) ICMS destacado', sum(x['icms'] * x['sinal'] for x in deb_lin), 'exclusão do ICMS da base (STF, Tema 69)')
    if c['excluir_ipi_debito']:
        m('produtos', '(−) IPI destacado', sum(x['ipi'] * x['sinal'] for x in deb_lin), 'IPI não integra a receita bruta (DL 1.598/77, art. 12, §4º)')
    if c['excluir_st_debito']:
        m('produtos', '(−) ICMS-ST', sum(x['st'] * x['sinal'] for x in deb_lin), 'ICMS-ST não integra a receita bruta (art. 12, §4º)')
    dv = sum(x['base'] for x in deb_lin if x['sinal'] < 0)
    if dv:
        m('produtos', '(−) Devoluções de venda (líquidas de ICMS/IPI)', dv, 'Lei 10.833/03, art. 1º, §3º, V, b')
    m('produtos', '= Base de débito de produtos', base_prod, 'valor contábil − ICMS%s%s − devoluções' % (' − IPI' if c['excluir_ipi_debito'] else '', ' − ST' if c['excluir_st_debito'] else ''), 'tot')

    base_cred_prod = _r(sum(x['base'] for x in cred_lin))
    if c.get('cred_prod_manual') is not None:
        base_cred_prod = _r(c['cred_prod_manual']); ed.append('entradas de produtos')
    for g in sorted({x['grupo'] for x in cred_lin}):
        m('creditos', 'Entradas: ' + g, sum(x['base'] for x in cred_lin if x['grupo'] == g), ' + '.join(x['cfop'] for x in cred_lin if x['grupo'] == g) + ' (valor contábil − ICMS%s)' % (' − IPI' if c['excluir_ipi_credito'] else ''))
    # importação
    imp = dados.get('importacao_xml') or {}
    imp_lin = [l for l in dados['cfops'] if l['cfop'][0] == '3' and c['creditos_cfop'].get(l['cfop'], l['cfop'][1:] in CRED_PADRAO)]
    imp_vc = sum(l['vc'] for l in imp_lin)
    if c.get('imp_base_manual') is not None:
        imp_vc = c['imp_base_manual']; c['importacao'] = 'base'; ed.append('importação')
    if c['imp_pis_manual'] is not None or c['imp_cof_manual'] is not None:
        imp_pis, imp_cof, imp_form = c['imp_pis_manual'] or 0, c['imp_cof_manual'] or 0, 'informado: PIS/COFINS-Importação pago (DI/DUIMP)'
    elif c['importacao'] == 'pago' and (imp.get('pis') or imp.get('cofins')):
        imp_pis, imp_cof, imp_form = imp['pis'], imp['cofins'], 'PIS/COFINS-Importação destacado nas NF-e de entrada de importação (Lei 10.865/04, art. 15)'
    else:
        imp_pis, imp_cof = imp_vc * PIS_NC / 100, imp_vc * COF_NC / 100
        imp_form = 'valor contábil %s × 1,65%% / 7,6%% (aproximação)' % _b(imp_vc)
        if imp_vc:
            alertas.append({'sev': 'alerta', 'tipo': 'premissa', 'txt': 'Importação: crédito calculado a 9,25%% sobre o valor contábil (%s). O crédito correto é o PIS/COFINS-Importação efetivamente pago (Lei 10.865/04, art. 15): informe os valores da DI/DUIMP.' % _b(imp_vc)})
    # serviços tomados (insumo)
    tom = [t for t in dados['tomados'] if t.get('cnpj') not in c['tomados_excluir'] and not (c['excluir_pf'] and (t.get('pf') or len(t.get('cnpj') or '') == 11))]
    tom_pf = [t for t in dados['tomados'] if c['excluir_pf'] and (t.get('pf') or len(t.get('cnpj') or '') == 11)]
    base_tom = c['tomados_base_manual'] if c['tomados_base_manual'] is not None else sum(t['valor'] for t in tom)
    m('creditos', 'Serviços tomados de PJ (insumo)', base_tom, 'informado' if c['tomados_base_manual'] is not None else '%d nota(s) de %s; pessoa física fora (não dá crédito)' % (len(tom), (dados['tomados'][0]['fonte'] if dados['tomados'] else 'nenhuma fonte')))
    if tom_pf:
        alertas.append({'sev': 'info', 'tipo': 'dados', 'txt': '%d serviço(s) tomado(s) de pessoa física (%s) ficaram fora do crédito: PF não gera crédito de PIS/COFINS.' % (len(tom_pf), _b(sum(t['valor'] for t in tom_pf)))})
    loc_desp = 0 if c['tomados_base_manual'] is not None else sum(x['valor'] for x in dados['locacoes'] if x['tipo'] == 'despesa' and len(x.get('cnpj') or '') == 14)
    if loc_desp:
        m('creditos', 'Aluguéis pagos a PJ', loc_desp, 'Lei 10.833/03, art. 3º, IV')
    base_cred = _r(base_cred_prod + base_tom + loc_desp)
    m('creditos', '= Base de créditos (9,25%)', base_cred, 'entradas de produtos + serviços tomados + aluguéis PJ', 'tot')
    outros_pis = sum(o.get('pis', 0) or 0 for o in c['outros_creditos'])
    outros_cof = sum(o.get('cofins', 0) or 0 for o in c['outros_creditos'])
    for o in c['outros_creditos']:
        if not (o.get('desc') or '').strip():
            alertas.append({'sev': 'alerta', 'tipo': 'premissa', 'txt': 'Há "outros créditos" (PIS %s / COFINS %s) sem descrição da origem: identifique a base legal (depreciação, ativo, armazenagem...).' % (_b(o.get('pis', 0)), _b(o.get('cofins', 0)))})

    # ---- 2. receitas de serviço e locação a distribuir
    serv_total = c['serv_manual'] if c['serv_manual'] is not None else sum(x['valor'] for x in dados['prestados'])
    loc_total = c['loc_manual'] if c['loc_manual'] is not None else sum(x['valor'] for x in dados['locacoes'] if x['tipo'] == 'receita')
    ret = {'pis': c['ret_pis_manual'] if c['ret_pis_manual'] is not None else sum(x.get('pis', 0) or 0 for x in dados['prestados']),
           'cofins': c['ret_cof_manual'] if c['ret_cof_manual'] is not None else sum(x.get('cofins', 0) or 0 for x in dados['prestados']),
           'irrf': c['ret_irrf_manual'] if c['ret_irrf_manual'] is not None else sum(x.get('irrf', 0) or 0 for x in dados['prestados']),
           'csll': c['ret_csll_manual'] if c['ret_csll_manual'] is not None else sum(x.get('csll', 0) or 0 for x in dados['prestados'])}
    iss_total = sum(x.get('iss', 0) or 0 for x in dados['prestados'])

    # créditos disponíveis da Matriz (valor), para achar a proporção que zera o PIS/COFINS
    cred_pis = base_cred * PIS_NC / 100 + imp_pis + outros_pis
    cred_cof = base_cred * COF_NC / 100 + imp_cof + outros_cof
    disp_cof = cred_cof + c['saldo_cof_ant']
    folga_base = disp_cof / (COF_NC / 100) - base_prod   # quanto de receita a Matriz ainda absorve com PIS/COFINS zero
    # opção: nota com PIS/COFINS retido vai direto para a SCP (no Presumido a retenção de 0,65% + 3% zera o PIS/COFINS da nota)
    retidas_on = bool(c.get('retidas_scp')) and c['modo'] == 'auto' and c['serv_manual'] is None
    tem_ret = lambda x: ((x.get('pis') or 0) + (x.get('cofins') or 0)) > 0.004
    pr_ret = [x for x in dados['prestados'] if tem_ret(x)] if retidas_on else []
    pr_liv = [x for x in dados['prestados'] if not tem_ret(x)] if retidas_on else list(dados['prestados'])
    serv_fixo_scp = sum(x['valor'] for x in pr_ret)
    serv_livre = serv_total - serv_fixo_scp
    if c['modo'] == 'auto':
        absorve = max(folga_base, 0)
        ordem = [('loc', loc_total), ('serv', serv_livre)] if c['prioridade'] == 'locacao' else [('serv', serv_livre), ('loc', loc_total)]
        fica = {}
        for k, tot in ordem:
            q = min(absorve, tot); fica[k] = q; absorve -= q
        serv_m, loc_m = fica.get('serv', 0), fica.get('loc', 0)
        if retidas_on and pr_ret:
            m('distribuicao', 'Notas com PIS/COFINS retido (direto na SCP)', serv_fixo_scp, '%d nota(s): %s' % (len(pr_ret), ', '.join(str(x.get('numero')) for x in pr_ret[:30])))
            m('distribuicao', 'Notas sem retenção (Matriz até igualar os créditos)', serv_livre, 'Matriz fica com %s; o resto vai para a SCP' % _b(serv_m))
    else:
        serv_m, loc_m = serv_total * c['pct_serv_matriz'] / 100, loc_total * c['pct_loc_matriz'] / 100
    serv_s, loc_s = serv_total - serv_m, loc_total - loc_m
    pct_serv_m = 100 * serv_m / serv_total if serv_total else 0
    pct_loc_m = 100 * loc_m / loc_total if loc_total else 0
    # retenções seguem a receita (proporcional); com a opção das retidas, as notas retidas levam a retenção inteira para a SCP
    manual_ret = any(c.get(k) is not None for k in ('ret_pis_manual', 'ret_cof_manual', 'ret_irrf_manual', 'ret_csll_manual'))
    if retidas_on and pr_ret and not manual_ret:
        sh_l = (serv_m / serv_livre) if serv_livre else 0
        soma = lambda xs, k: sum(x.get(k, 0) or 0 for x in xs)
        ret_m = {k: soma(pr_liv, k) * sh_l for k in ret}
        ret_s = {k: soma(pr_ret, k) + soma(pr_liv, k) * (1 - sh_l) for k in ret}
    else:
        sh_m = (serv_m / serv_total) if serv_total else 0
        ret_m = {k: v * sh_m for k, v in ret.items()}
        ret_s = {k: v * (1 - sh_m) for k, v in ret.items()}
    for kk in ('pis', 'cofins', 'irrf', 'csll'):
        if c.get('ret_m_' + kk) is not None:
            ret_m[kk] = c['ret_m_' + kk]; ed.append('retenção Matriz')
        if c.get('ret_s_' + kk) is not None:
            ret_s[kk] = c['ret_s_' + kk]; ed.append('retenção SCP')

    # ---- 3. Matriz: PIS e COFINS não cumulativos (cada tributo separado; créditos, saldo anterior e retenção não se misturam)
    def apura_nc(nome, aliq, cred_periodo, saldo_ant, retido):
        base_deb = base_prod + serv_m + loc_m
        deb = base_deb * aliq / 100
        disp = cred_periodo + saldo_ant
        usado = min(deb, disp)
        devido = deb - usado
        ret_usado = min(devido, retido)
        return {'base_deb': _r(base_deb), 'debito': _r(deb), 'cred_periodo': _r(cred_periodo), 'saldo_ant': _r(saldo_ant), 'disponivel': _r(disp),
                'cred_usado': _r(usado), 'saldo_credor_final': _r(disp - usado), 'devido_apos_creditos': _r(devido), 'retencao': _r(retido),
                'retencao_usada': _r(ret_usado), 'retencao_excedente': _r(retido - ret_usado), 'a_recolher': _r(devido - ret_usado)}
    mp = apura_nc('PIS', PIS_NC, cred_pis, c['saldo_pis_ant'], ret_m['pis'])
    mc = apura_nc('COFINS', COF_NC, cred_cof, c['saldo_cof_ant'], ret_m['cofins'])
    for nome, a, aliq, cp, op, ip in (('PIS', mp, PIS_NC, base_cred * PIS_NC / 100, outros_pis, imp_pis), ('COFINS', mc, COF_NC, base_cred * COF_NC / 100, outros_cof, imp_cof)):
        s = 'matriz_' + nome.lower()
        m(s, 'Receita de produtos (base)', base_prod, 'seção Produtos')
        m(s, 'Receita de serviços na Matriz', serv_m, '%s%% de %s' % (_b(pct_serv_m), _b(serv_total)))
        m(s, 'Receita de locação na Matriz', loc_m, '%s%% de %s' % (_b(pct_loc_m), _b(loc_total)))
        m(s, '= Base de cálculo do débito', a['base_deb'], 'produtos + serviços + locação na Matriz', 'tot')
        m(s, 'Débito %s (%s%%)' % (nome, _b(aliq)), a['debito'], '%s × %s%%' % (_b(a['base_deb']), _b(aliq)), 'deb')
        m(s, 'Créditos das entradas', cp, '%s × %s%%' % (_b(base_cred), _b(aliq)), 'cred')
        if ip:
            m(s, 'Crédito da importação', ip, imp_form, 'cred')
        if op:
            m(s, 'Outros créditos (informados)', op, '; '.join('%s: %s' % (o.get('desc') or 'sem descrição', _b(o.get(nome.lower(), 0))) for o in c['outros_creditos']), 'cred')
        m(s, 'Saldo credor de períodos anteriores', a['saldo_ant'], 'informado/mês anterior', 'cred')
        m(s, 'Créditos utilizados no período', a['cred_usado'], 'menor entre débito (%s) e créditos disponíveis (%s)' % (_b(a['debito']), _b(a['disponivel'])))
        m(s, '= Débito após os créditos', a['devido_apos_creditos'], 'débito − créditos utilizados', 'tot')
        m(s, '(−) %s retido nas notas de serviço (dedução)' % nome, a['retencao_usada'], 'retido na Matriz %s; usa até o devido' % _b(a['retencao']), 'ret')
        m(s, '= %s a recolher' % nome, a['a_recolher'], 'débito após créditos − retenção usada', 'tot')
        m(s, 'Saldo credor a transportar', a['saldo_credor_final'], 'créditos disponíveis − utilizados (não se mistura com retenção)', 'saldo')
        if a['retencao_excedente'] > 0.005:
            m(s, 'Retenção sem débito para deduzir', a['retencao_excedente'], 'não vira crédito do não cumulativo: restituição/compensação (PER/DCOMP)', 'saldo')

    # ---- 4. SCP: PIS/COFINS cumulativo, IRPJ/CSLL presumido (trimestral)
    base_scp = serv_s + loc_s + (c.get('scp_outras') or 0)
    sp = {'base': _r(base_scp), 'pis': _r(base_scp * PIS_CU / 100), 'cofins': _r(base_scp * COF_CU / 100)}
    for k in ('pis', 'cofins'):
        sp[k + '_ret'] = _r(min(sp[k], ret_s[k])); sp[k + '_rec'] = _r(sp[k] - sp[k + '_ret']); sp[k + '_ret_exc'] = _r(ret_s[k] - sp[k + '_ret'])
    m('scp_pc', 'Receita de serviços na SCP', serv_s, '%s%% de %s' % (_b(100 - pct_serv_m), _b(serv_total)))
    m('scp_pc', 'Receita de locação na SCP', loc_s, '%s%% de %s' % (_b(100 - pct_loc_m), _b(loc_total)))
    m('scp_pc', '= Base de cálculo (regime cumulativo)', base_scp, 'sem créditos (Lei 9.718/98)', 'tot')
    m('scp_pc', 'PIS 0,65%', sp['pis'], '%s × 0,65%%' % _b(base_scp), 'deb')
    m('scp_pc', 'COFINS 3%', sp['cofins'], '%s × 3%%' % _b(base_scp), 'deb')
    m('scp_pc', '(−) PIS retido', sp['pis_ret'], 'retido nas notas da SCP', 'ret')
    m('scp_pc', '(−) COFINS retida', sp['cofins_ret'], 'retido nas notas da SCP', 'ret')
    m('scp_pc', '= PIS a recolher', sp['pis_rec'], '', 'tot')
    m('scp_pc', '= COFINS a recolher', sp['cofins_rec'], '', 'tot')
    # IRPJ/CSLL trimestral
    comp = dados.get('competencia') or ''
    outros = [x for x in c['tri_outros_meses'] if x.get('competencia') != comp]
    base_tri_rec = base_scp + sum(x.get('receita', 0) or 0 for x in outros)
    meses_tri = 1 + len(outros)
    pres = base_tri_rec * 0.32
    irpj = pres * 0.15
    adic = max(pres - 20000 * 3, 0) * 0.10
    csll = pres * 0.09
    irrf_tri = ret_s['irrf'] + sum(x.get('irrf', 0) or 0 for x in outros)
    csll_ret_tri = ret_s['csll'] + sum(x.get('csll', 0) or 0 for x in outros)
    irpj_rec = max(irpj + adic - irrf_tri - c['irpj_pago_tri'], 0)
    csll_rec = max(csll - csll_ret_tri - c['csll_pago_tri'], 0)
    # mês isolado (como a aba Economia: adicional sobre R$ 20 mil no mês)
    pres_mes = base_scp * 0.32
    irpj_mes = pres_mes * 0.15 + max(pres_mes - 20000, 0) * 0.10
    csll_mes = pres_mes * 0.09
    m('scp_ir', 'Receita da SCP no trimestre', base_tri_rec, '%d mês(es): este + %s' % (meses_tri, ', '.join('%s %s' % (x.get('competencia'), _b(x.get('receita', 0))) for x in outros) or 'sem os outros meses'))
    m('scp_ir', 'Base presumida (32%)', pres, '%s × 32%% (serviços e locação; RIR/2018, art. 592)' % _b(base_tri_rec), 'tot')
    m('scp_ir', 'IRPJ 15%', irpj, '%s × 15%%' % _b(pres), 'deb')
    m('scp_ir', 'Adicional de IRPJ 10%', adic, '(%s − 60.000,00) × 10%% (R$ 20 mil × 3 meses)' % _b(pres), 'deb')
    m('scp_ir', 'CSLL 9%', csll, '%s × 9%%' % _b(pres), 'deb')
    m('scp_ir', '(−) IRRF retido no trimestre', irrf_tri, 'retido nas notas de serviço da SCP', 'ret')
    m('scp_ir', '(−) CSLL retida no trimestre', csll_ret_tri, '', 'ret')
    m('scp_ir', '(−) IRPJ/CSLL já recolhidos no trimestre', c['irpj_pago_tri'] + c['csll_pago_tri'], 'informado', 'ret')
    m('scp_ir', '= IRPJ a recolher no trimestre', irpj_rec, '', 'tot')
    m('scp_ir', '= CSLL a recolher no trimestre', csll_rec, '', 'tot')
    m('scp_ir', 'Quota mensal (3 quotas)', (irpj_rec + csll_rec) / 3, 'opção por quotas: IRPJ e CSLL ÷ 3', 'saldo')
    if meses_tri < 3:
        alertas.append({'sev': 'alerta', 'tipo': 'dados', 'txt': 'IRPJ/CSLL da SCP é trimestral: informe a receita dos outros %d mês(es) do trimestre (ou feche esses meses no sistema) para o adicional e as quotas ficarem certos.' % (3 - meses_tri)})

    # ---- 5. Matriz IRPJ/CSLL (Lucro Real): depende do resultado contábil
    # lucro da Matriz = resultado informado (padrão zero: sem lucro) + margem sobre serviço/locação que ficam nela
    lucro = (c['lucro_matriz'] or 0) + (serv_m + loc_m) * c['margem_serv_real'] / 100
    irpj_m = max(lucro, 0) * 0.15 + max(lucro - 20000, 0) * 0.10 if lucro > 0 else 0
    csll_m = max(lucro, 0) * 0.09
    m('matriz_ir', 'Resultado do período informado', c['lucro_matriz'] or 0, 'pelo balancete; padrão zero (empresa sem lucro)')
    m('matriz_ir', '(+) Margem sobre serviço/locação mantidos na Matriz', (serv_m + loc_m) * c['margem_serv_real'] / 100, '%s × %s%%' % (_b(serv_m + loc_m), _b(c['margem_serv_real'])))
    m('matriz_ir', '= Lucro real estimado', lucro, 'sem lucro, não há IRPJ/CSLL', 'tot')
    m('matriz_ir', 'IRPJ', irpj_m, '15% + adicional 10% acima de R$ 20.000' if lucro > 0 else 'sem lucro no período', 'deb')
    m('matriz_ir', 'CSLL', csll_m, '9%' if lucro > 0 else 'sem lucro no período', 'deb')

    # ---- 6. comparativo do mês: estratégia × tudo no Lucro Real × tudo no Presumido
    est = {'pis_cofins': mp['a_recolher'] + mc['a_recolher'] + sp['pis_rec'] + sp['cofins_rec'],
           'irpj_csll': irpj_m + csll_m + max(irpj_mes - ret_s['irrf'], 0) + max(csll_mes - ret_s['csll'], 0)}
    # tudo na Matriz
    base_all = base_prod + serv_total + loc_total
    def nc_all(aliq, cred, sant, retido):
        d = base_all * aliq / 100
        dev = max(d - cred - sant, 0)
        return max(dev - retido, 0)
    lucro_serv = (serv_total + loc_total) * c['margem_serv_real'] / 100 + (c['lucro_matriz'] or 0)
    tr = {'pis_cofins': nc_all(PIS_NC, cred_pis, c['saldo_pis_ant'], ret['pis']) + nc_all(COF_NC, cred_cof, c['saldo_cof_ant'], ret['cofins']),
          'irpj_csll': max(max(lucro_serv, 0) * 0.15 + max(lucro_serv - 20000, 0) * 0.10 - ret['irrf'], 0) + max(max(lucro_serv, 0) * 0.09 - ret['csll'], 0)}
    # tudo no Presumido
    pres_p = base_prod * 0.08 + (serv_total + loc_total) * 0.32
    pres_c = base_prod * 0.12 + (serv_total + loc_total) * 0.32
    tp = {'pis_cofins': max(base_all * (PIS_CU + COF_CU) / 100 - ret['pis'] - ret['cofins'], 0),
          'irpj_csll': max(pres_p * 0.15 + max(pres_p - 20000, 0) * 0.10 - ret['irrf'], 0) + max(pres_c * 0.09 - ret['csll'], 0)}
    for x in (est, tr, tp):
        x['total'] = _r(x['pis_cofins'] + x['irpj_csll']); x['pis_cofins'] = _r(x['pis_cofins']); x['irpj_csll'] = _r(x['irpj_csll'])
    receita_total = base_prod + serv_total + loc_total
    comparativo = {'estrategia': est, 'lucro_real': tr, 'presumido': tp, 'iss': _r(iss_total), 'receita': _r(receita_total),
                   'economia_vs_real': _r(tr['total'] - est['total']), 'economia_vs_presumido': _r(tp['total'] - est['total']),
                   'premissas': ['Mês isolado; ICMS, IPI e ISS são iguais nos três cenários e ficam fora da soma.',
                                 'Lucro da Matriz, nos dois cenários: resultado informado (padrão zero, sem lucro) mais a margem de %s%% (editável) sobre o serviço e a locação que ficarem nela.' % _b(c['margem_serv_real']),
                                 'Tudo no Presumido: PIS/COFINS cumulativo sobre toda a receita (sem ICMS/IPI), presunção de 8%/12% nos produtos e 32% em serviços e locação, adicional mensal acima de R$ 20 mil.',
                                 'Estratégia: IRPJ/CSLL da SCP no mês isolado (adicional mensal), para comparar; a guia sai pela apuração trimestral.']}

    # ---- 7. alertas e riscos
    ic = dados.get('icms_pis_saidas') or {}
    if ic.get('com_icms'):
        alertas.append({'sev': 'info', 'tipo': 'icms', 'txt': '%d NF-e de venda saíram com o ICMS dentro da base do PIS/COFINS no XML (%s de ICMS). A exclusão (Tema 69) é feita aqui na apuração; confira se foi ajustada na EFD-Contribuições.' % (ic['com_icms'], _b(ic['icms_dentro']))})
    if not c['scp_cnpj']:
        alertas.append({'sev': 'info', 'tipo': 'dados', 'txt': 'Informe o CNPJ e o nome da SCP para os relatórios.'})
    if loc_s > 0:
        if any((x.get('iss') or 0) > 0 and 'loca' in (x.get('descricao') or '').lower() for x in dados['prestados']):
            alertas.append({'sev': 'alerta', 'tipo': 'premissa', 'txt': 'Há ISS sobre locação: locação de bens móveis não sofre ISS (STF, Súmula Vinculante 31).'})
    if ret_m['pis'] + ret_m['cofins'] > 0 and (mp['retencao_excedente'] + mc['retencao_excedente']) > 0.005:
        alertas.append({'sev': 'alerta', 'tipo': 'premissa', 'txt': 'A Matriz fica com PIS/COFINS retido (%s) sem débito para deduzir: vira restituição/compensação. Serviços com retenção rendem mais na SCP.' % _b(mp['retencao_excedente'] + mc['retencao_excedente'])})
    if folga_base < 0:
        alertas.append({'sev': 'info', 'tipo': 'resultado', 'txt': 'Os créditos da Matriz não cobrem nem a receita de produtos (faltam %s de base): ela paga PIS/COFINS sobre o excedente e não há como passar produto para a SCP.' % _b(-folga_base)})
    if receita_total and (serv_s + loc_s) * 12 > 78000000:
        alertas.append({'sev': 'risco', 'tipo': 'risco', 'txt': 'A receita da SCP anualizada passa de R$ 78 milhões, limite do Lucro Presumido (Lei 9.718/98, art. 13).'})
    if fora_lin:
        alertas.append({'sev': 'info', 'tipo': 'dados', 'txt': 'Entradas fora do crédito por padrão: %s. Marque as que forem insumo.' % ', '.join('%s (%s)' % (x['cfop'], _b(x['base'])) for x in fora_lin[:12])})
    if not dados['prestados'] and c['serv_manual'] is None:
        alertas.append({'sev': 'alerta', 'tipo': 'dados', 'txt': 'Sem relatório de saída de serviços e sem NFS-e prestadas: a receita de serviços está zerada.'})
    if not dados['locacoes'] and c['loc_manual'] is None:
        alertas.append({'sev': 'info', 'tipo': 'dados', 'txt': 'Sem relatório de locações: a receita de locação está zerada (informe o valor ou envie o relatório).'})
    if c['excluir_ipi_credito'] is False and dados.get('credita_ipi'):
        alertas.append({'sev': 'alerta', 'tipo': 'premissa', 'txt': 'O IPI das entradas está entrando na base do crédito, mas a empresa credita IPI: o IPI recuperável não compõe o custo (base do crédito).'})
    if not c['excluir_ipi_debito']:
        alertas.append({'sev': 'alerta', 'tipo': 'premissa', 'txt': 'O IPI destacado nas vendas está na base do débito: ele não compõe a receita bruta (DL 1.598/77, art. 12, §4º).'})
    if not c['deduzir_devolucoes']:
        alertas.append({'sev': 'alerta', 'tipo': 'premissa', 'txt': 'Devoluções fora do cálculo: devolução de venda reduz o débito e devolução de compra reduz o crédito.'})

    resumo = {'matriz': {'receita': _r(base_prod + serv_m + loc_m), 'creditos': _r(mp['disponivel'] + mc['disponivel']), 'debitos': _r(mp['debito'] + mc['debito']),
                         'pis_cofins': _r(mp['a_recolher'] + mc['a_recolher']), 'irpj_csll': _r(irpj_m + csll_m), 'saldo_credor': _r(mp['saldo_credor_final'] + mc['saldo_credor_final'])},
              'scp': {'receita': _r(base_scp), 'pis_cofins': _r(sp['pis_rec'] + sp['cofins_rec']), 'irpj_csll_tri': _r(irpj_rec + csll_rec),
                      'irpj_csll_mes': _r(max(irpj_mes - ret_s['irrf'], 0) + max(csll_mes - ret_s['csll'], 0))},
              'distribuicao': {'serv_total': _r(serv_total), 'loc_total': _r(loc_total), 'serv_matriz': _r(serv_m), 'loc_matriz': _r(loc_m), 'serv_scp': _r(serv_s),
                               'loc_scp': _r(loc_s), 'pct_serv_matriz': _r(pct_serv_m), 'pct_loc_matriz': _r(pct_loc_m), 'folga_base': _r(folga_base), 'modo': c['modo']}}
    if ed:
        alertas.append({'sev': 'info', 'tipo': 'adm', 'txt': 'Valores digitados pelo ADM substituem o calculado em: %s.' % ', '.join(sorted(set(ed)))})
    return {'cfg': c, 'editados': sorted(set(ed)), 'resumo': resumo, 'matriz': {'pis': mp, 'cofins': mc}, 'scp': {'pc': sp, 'irpj': _r(irpj + adic), 'csll': _r(csll), 'irpj_rec': _r(irpj_rec),
            'csll_rec': _r(csll_rec), 'base_tri': _r(base_tri_rec), 'pres_tri': _r(pres)}, 'memoria': dict(mem), 'comparativo': comparativo, 'alertas': alertas,
            'linhas_debito': deb_lin, 'linhas_credito': cred_lin, 'linhas_fora': fora_lin, 'linhas_debito_fora': deb_fora, 'retencoes': {'total': {k: _r(v) for k, v in ret.items()},
            'matriz': {k: _r(v) for k, v in ret_m.items()}, 'scp': {k: _r(v) for k, v in ret_s.items()}}, 'iss': _r(iss_total),
            'prestados': dados['prestados'][:500], 'tomados': dados['tomados'][:500],
            'importacao': {'modo': 'manual' if (c['imp_pis_manual'] is not None or c['imp_cof_manual'] is not None) else ('pago' if (c['importacao'] == 'pago' and (imp.get('pis') or imp.get('cofins'))) else 'base'),
                           'base': _r(imp_vc), 'pis': _r(imp_pis), 'cofins': _r(imp_cof)}}
