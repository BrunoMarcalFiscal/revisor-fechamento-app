# -*- coding: utf-8 -*-
"""Visão estratégica do fechamento: evolução mês a mês, margem bruta estimada, saldo credor e o que fazer com ele,
créditos não aproveitados, projeção do Lucro Real (PIS/COFINS não cumulativo e estimativa de IRPJ/CSLL), base para a
simulação da reforma (CBS/IBS), concentração e risco, e indicadores do fechamento.
Tudo marcado como estimativa quando depende de premissa; as premissas vão junto no resultado."""
from collections import defaultdict
from .painel import MP, FINALIDADE

REVENDA = FINALIDADE[1][1]
PC_NC = 9.25  # PIS 1,65% + COFINS 7,6% (Leis 10.637/02 e 10.833/03)


def _sufs_lancados(n):
    out = []
    for p in n.get('pares') or []:
        s = p.get('sis') or {}
        if s.get('cfop'):
            out.append((s['cfop'][1:], s.get('vc') or 0))
    return out


def _base_compras(res):
    """Compras de terceiros que dão crédito (insumo/revenda), sem ICMS e IPI destacados, e as compras do Simples."""
    insumo = simples = frete = 0.0
    for n in res['notas']:
        if n.get('natureza') != 'normal' or n['mov'] not in ('entrada', 'frete'):
            continue
        f = n.get('fiscal') or {}
        if n['tipo'] == 'CT-e':
            frete += (f.get('vc') or 0) - (f.get('icms') or 0)
            continue
        if n['tipo'] != 'NF-e' or f.get('fonte') != 'xml':
            continue
        ls = _sufs_lancados(n)
        tot = sum(v for _, v in ls) or 0
        share = (sum(v for s, v in ls if s in MP or s in REVENDA) / tot) if tot else 0
        base = ((f.get('vc') or 0) - (f.get('icms') or 0) - (f.get('ipi') or 0)) * share
        if f.get('crt') in ('1', '4'):
            simples += base
        else:
            insumo += base
    return round(insumo, 2), round(simples, 2), round(frete, 2)


def montar(res, apu=None, hist=None, recebido_em=None):
    P = res.get('painel') or {}
    f = P.get('faturamento') or {}
    imp = P.get('impostos') or {}
    calc = (apu or {}).get('calc') or {}
    hist = [h for h in (hist or []) if h.get('cnpj') == res['empresa'].get('cnpj')]
    prod_liq, serv, loc = f.get('liquido', 0), f.get('servicos', 0), f.get('locacoes', 0)
    receita = round(prod_liq + serv + loc, 2)

    # ---- 2. margem bruta estimada (produtos)
    fin = {x['nome']: x['valor'] for x in P.get('finalidade', [])}
    custo = round(fin.get('Matéria-prima / industrialização', 0) + fin.get('Mercadoria para revenda', 0), 2)
    margem = {'vendas_liquidas': prod_liq, 'compras_mp_revenda': custo, 'margem': round(prod_liq - custo, 2),
              'margem_pct': round(100 * (prod_liq - custo) / prod_liq, 1) if prod_liq else None,
              'premissa': 'Vendas líquidas de produtos menos as compras de matéria-prima e mercadoria do mesmo mês. Não considera a variação de estoque: '
                          'mês de compra grande para estoque derruba a margem; mês que vende estoque a eleva.'}

    # ---- 3. saldo credor e o que fazer
    icms_saldo = calc.get('saldo')
    ipi_saldo = calc.get('ipiSaldo')
    if icms_saldo is None:
        t = (res.get('apuracao') or {}).get('totais') or {}
        if t:
            icms_saldo = round(t['saida'].get('icms', 0) - t['entrada'].get('icms', 0), 2)
            ipi_saldo = round(t['saida'].get('ipi', 0) - t['entrada'].get('ipi', 0), 2)
    cfops_saida = {c for n in res['notas'] if n['mov'] == 'saida' and n.get('natureza') == 'normal' for c in (n.get('fiscal') or {}).get('cfops', [])}
    exporta = any(c.startswith('7') for c in cfops_saida)
    deb_med = [h['indicadores'].get('icms_saldo', 0) for h in hist if h.get('indicadores', {}).get('icms_saldo', 0) > 0]
    saldo = {'icms': {'saldo': icms_saldo, 'credor': (icms_saldo or 0) < 0, 'valor': abs(icms_saldo or 0)},
             'ipi': {'saldo': ipi_saldo, 'credor': (ipi_saldo or 0) < 0, 'valor': abs(ipi_saldo or 0)}, 'acoes': []}
    if saldo['ipi']['credor']:
        saldo['acoes'].append({'imposto': 'IPI', 'acao': 'Pedido de ressarcimento ou compensação do saldo credor trimestral (PER/DCOMP)',
                               'base': 'Lei 9.779/99, art. 11; IN RFB 2.055/2021', 'quando': 'saldo credor que se mantém ao fim do trimestre-calendário'})
    if saldo['icms']['credor']:
        if exporta:
            saldo['acoes'].append({'imposto': 'ICMS', 'acao': 'Transferência do saldo credor gerado por exportação a outros estabelecimentos ou a terceiros',
                                   'base': 'LC 87/96, art. 25, §1º, e RICMS da UF', 'quando': 'há saídas para o exterior (CFOP 7xxx)'})
        saldo['acoes'].append({'imposto': 'ICMS', 'acao': 'Conferir no RICMS da UF as hipóteses de transferência ou uso do saldo credor acumulado (fornecedores, aquisição de ativo, ST)',
                               'base': 'RICMS da UF', 'quando': 'saldo credor recorrente'})
        if deb_med:
            saldo['meses_para_consumir'] = round(saldo['icms']['valor'] / (sum(deb_med) / len(deb_med)), 1)

    # ---- 4. créditos não aproveitados
    cred = defaultdict(lambda: {'notas': set(), 'valor': 0.0})
    for n in res['notas']:
        for a in n['aponts']:
            if a['sev'] not in ('erro', 'alerta'):
                continue
            k = {'CREDITO_NAO_TOMADO': 'ICMS destacado sem crédito', 'CREDITO_MENOR': 'ICMS creditado a menor', 'CREDITO_DEV_VENDA': 'ICMS de devolução de venda sem crédito',
                 'PC_NAO_TOMADO': 'PIS/COFINS de compra tributada sem crédito (estimado)', 'PC_FRETE': 'PIS/COFINS de frete sem crédito (estimado)'}.get(a['cod'])
            if k:
                cred[k]['notas'].add(n['id']); cred[k]['valor'] += a['valor'] or 0
    sn = [x for x in (res.get('apuracao') or {}).get('simples', []) if not x.get('sem_direito') and not x.get('ja_no_livro')]
    if sn:
        cred['Crédito do Simples Nacional destacado e fora do livro']['notas'] = set(range(len(sn)))
        cred['Crédito do Simples Nacional destacado e fora do livro']['valor'] = sum(x['credito'] for x in sn)
    creditos = sorted([{'tipo': k, 'notas': len(v['notas']), 'valor': round(v['valor'], 2)} for k, v in cred.items() if v['valor'] > 0.005], key=lambda x: -x['valor'])

    # ---- 5. projeção Lucro Real
    insumo, simples, frete = _base_compras(res)
    loc_desp = f.get('locacoes_despesa', 0)
    pc_deb = round(receita * PC_NC / 100, 2)
    pc_cred = round((insumo + simples + frete + loc_desp) * PC_NC / 100, 2)
    irpj_base = round(prod_liq * 0.08 + (serv + loc) * 0.32, 2)
    csll_base = round(prod_liq * 0.12 + (serv + loc) * 0.32, 2)
    irpj = round(irpj_base * 0.15 + max(irpj_base - 20000, 0) * 0.10, 2)
    csll = round(csll_base * 0.09, 2)
    lr = {'receita_base': receita, 'pc_debito': pc_deb, 'pc_credito': pc_cred, 'pc_saldo': round(pc_deb - pc_cred, 2),
          'pc_destacado_xml': round((imp.get('pis_saidas', 0) or 0) + (imp.get('cofins_saidas', 0) or 0), 2),
          'base_credito': {'insumo_revenda': insumo, 'simples': simples, 'frete': frete, 'alugueis_pj': loc_desp},
          'irpj_base': irpj_base, 'irpj': irpj, 'csll_base': csll_base, 'csll': csll,
          'premissas': ['PIS/COFINS não cumulativo a 9,25% sobre a receita líquida do mês (produtos, serviços e locações); algumas receitas podem ser cumulativas ou ter alíquota própria.',
                        'Créditos sobre compras para industrialização/revenda (sem ICMS e IPI destacados), fretes e aluguéis de PJ. Energia, depreciação e outros insumos não entram aqui.',
                        'IRPJ/CSLL pela estimativa mensal sobre a receita bruta (Lei 9.430/96, art. 2º): presunção de 8%/12% (produtos) e 32% (serviços e locações); adicional de 10% acima de R$ 20.000.',
                        'No Lucro Real a estimativa pode ser suspensa ou reduzida por balancete; o valor definitivo depende do resultado contábil.']}

    # ---- 6. base para a simulação da reforma (a alíquota é editada na tela)
    reforma = {'receita_base': receita, 'icms_saidas': imp.get('icms_saidas', 0), 'compras_credito': round(insumo + frete + loc_desp, 2), 'compras_simples': simples,
               'pis_cofins_hoje': lr['pc_saldo'], 'ref_cbs': 8.8, 'ref_ibs': 17.7,
               'premissas': ['Alíquotas de referência estimadas (as definitivas saem por resolução do Senado); edite na tela.',
                             '2026: CBS 0,9% e IBS 0,1% destacados, compensáveis com PIS/COFINS (LC 214/2025).',
                             '2027: CBS cheia no lugar do PIS/COFINS; IBS 0,1%; ICMS e ISS continuam até a transição de 2029 a 2032.',
                             'Base sem ICMS, ISS e PIS/COFINS. Crédito amplo sobre as compras de contribuintes do regime regular; do Simples, só o que ele recolhe.']}

    # ---- 7. concentração e risco
    cli = [c for c in P.get('clientes', []) if not c.get('demais')]
    forn = [c for c in P.get('fornecedores', []) if not c.get('demais')]
    risco = {'top1_cliente': cli[0] if cli else None, 'top5_clientes_pct': P.get('concentracao_top5', 0),
             'top1_fornecedor': forn[0] if forn else None, 'simples_pct_compras': (P.get('simples') or {}).get('pct_compras', 0), 'alertas': []}
    if cli and cli[0]['pct'] >= 40:
        risco['alertas'].append('Um cliente concentra %s%% da receita (%s): perda dele afeta o caixa e o fluxo de impostos.' % (('%.1f' % cli[0]['pct']).replace('.', ','), cli[0]['nome']))
    if risco['top5_clientes_pct'] >= 80:
        risco['alertas'].append('Os 5 maiores clientes somam %s%% da receita.' % ('%.1f' % risco['top5_clientes_pct']).replace('.', ','))
    if forn and forn[0]['pct'] >= 40:
        risco['alertas'].append('Um fornecedor concentra %s%% das compras (%s).' % (('%.1f' % forn[0]['pct']).replace('.', ','), forn[0]['nome']))
    if risco['simples_pct_compras'] >= 30:
        risco['alertas'].append('%s%% das compras vêm do Simples: crédito de ICMS limitado e, na reforma, crédito de CBS/IBS só do que o fornecedor recolher.' % ('%.1f' % risco['simples_pct_compras']).replace('.', ','))

    # ---- 8. indicadores do fechamento
    tot = sum(1 for n in res['notas'] for a in n['aponts'] if a['sev'] in ('erro', 'alerta'))
    por_cat = defaultdict(int)
    for n in res['notas']:
        for a in n['aponts']:
            if a['sev'] in ('erro', 'alerta'):
                por_cat[a.get('categoria', 'outros')] += 1
    ind = {'documentos_xml': res['resumo']['xml'], 'lancamentos': res['resumo']['lancamentos'], 'notas_com_erro': res['resumo']['com_erro'],
           'apontamentos': tot, 'por_categoria': dict(por_cat), 'recebido_em': recebido_em, 'calculado_em': res.get('gerado_em'),
           'taxa_erro': round(100 * res['resumo']['com_erro'] / res['resumo']['notas'], 1) if res['resumo']['notas'] else 0}

    # ---- 1. evolução
    atual = indicadores(res, calc)
    evol = sorted([h['indicadores'] for h in hist if h.get('indicadores') and h['indicadores'].get('competencia') != atual['competencia']] + [atual], key=lambda x: x['competencia'])
    return {'evolucao': evol, 'margem': margem, 'saldo': saldo, 'creditos': creditos, 'lucro_real': lr, 'reforma': reforma, 'risco': risco, 'operacao': ind}


def indicadores(res, calc=None):
    """Resumo do mês guardado no histórico, para a evolução."""
    P = res.get('painel') or {}
    f = P.get('faturamento') or {}
    fin = {x['nome']: x['valor'] for x in P.get('finalidade', [])}
    custo = fin.get('Matéria-prima / industrialização', 0) + fin.get('Mercadoria para revenda', 0)
    calc = calc or {}
    t = (res.get('apuracao') or {}).get('totais') or {'entrada': {}, 'saida': {}}
    icms = calc.get('saldo', round(t['saida'].get('icms', 0) - t['entrada'].get('icms', 0), 2) if t.get('saida') else 0)
    ipi = calc.get('ipiSaldo', round(t['saida'].get('ipi', 0) - t['entrada'].get('ipi', 0), 2) if t.get('saida') else 0)
    return {'competencia': res.get('competencia'), 'produtos': f.get('liquido', 0), 'servicos': f.get('servicos', 0), 'locacoes': f.get('locacoes', 0),
            'receita': round(f.get('liquido', 0) + f.get('servicos', 0) + f.get('locacoes', 0), 2), 'compras': P.get('compras_total', 0),
            'margem': round(f.get('liquido', 0) - custo, 2), 'icms_saldo': icms, 'ipi_saldo': ipi,
            'notas': res['resumo']['notas'], 'com_erro': res['resumo']['com_erro'], 'carga': P.get('carga_efetiva', 0)}
