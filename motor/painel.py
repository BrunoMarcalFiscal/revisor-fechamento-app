# -*- coding: utf-8 -*-
"""Painel gerencial da empresa a partir dos XML do mês (documentos reais) e do lançamento.

Faturamento bruto − devoluções = líquido, serviços prestados, impostos do período (destacados nas notas e
retenções), clientes que mais compram, fornecedores (geral, matéria-prima e Simples), compras por finalidade,
vendas por UF e produtos. Não há guia de pagamento nos arquivos: "impostos" aqui é o que foi destacado/apurado."""
from collections import defaultdict
from .regras_base import desc_cfop
from .util import raiz

VENDA = {'101', '102', '103', '104', '105', '106', '107', '108', '109', '110', '111', '112', '113', '114', '115', '118', '119', '120',
         '116', '117', '122', '123', '124', '125', '401', '402', '403', '404', '405', '651', '652', '653', '654', '655', '656', '667'}
SERVICO_NFE = {'933'}  # serviço tributado pelo ISS em NF-e conjugada: vai para o faturamento de serviços
ENTREGA_FUTURA = {'922'}  # simples faturamento de venda para entrega futura: a receita entra na remessa (5116/5117)
FINALIDADE = [  # (rótulo, sufixos do CFOP de ENTRADA lançado)
    ('Matéria-prima / industrialização', {'101', '111', '116', '120', '122', '124', '125', '401', '651', '652'}),
    ('Mercadoria para revenda', {'102', '113', '117', '118', '121', '403'}),
    ('Uso e consumo', {'556', '407'}),
    ('Ativo imobilizado', {'551', '406'}),
    ('Combustível', {'653'}),
    ('Frete', {'351', '352', '353', '354', '355', '356', '360'}),
    ('Devoluções de venda', {'201', '202', '410', '411', '553'}),
    ('Remessas e retornos', {'901', '902', '903', '908', '909', '910', '911', '912', '913', '915', '916', '920', '921', '925', '949'}),
]
MP = FINALIDADE[0][1]


def _top(d, n=10):
    tot = sum(v['valor'] for v in d.values()) or 1
    lst = sorted(d.values(), key=lambda v: -v['valor'])
    out = [dict(v, pct=round(100 * v['valor'] / tot, 1)) for v in lst[:n]]
    resto = lst[n:]
    if resto:
        out.append({'nome': 'Demais (%d)' % len(resto), 'cnpj': '', 'uf': '', 'notas': sum(v['notas'] for v in resto),
                    'valor': round(sum(v['valor'] for v in resto), 2), 'pct': round(100 * sum(v['valor'] for v in resto) / tot, 1), 'demais': True})
    return out


def _acc(d, chave, nome, cnpj, uf, valor, extra=None):
    if cnpj and len(cnpj) == 14 and chave == cnpj:
        chave = raiz(cnpj)  # matriz e filiais juntas
    e = d.get(chave)
    if e is None:
        e = d[chave] = {'nome': nome, 'cnpj': cnpj, 'uf': uf, 'notas': 0, 'valor': 0.0, 'docs': set(), 'estab': set()}
    if cnpj:
        e['estab'].add(cnpj)
    e['valor'] = round(e['valor'] + valor, 2)
    if extra and extra not in e['docs']:
        e['docs'].add(extra); e['notas'] += 1


def _limpa(lst):
    for v in lst:
        v.pop('docs', None)
        est = v.pop('estab', None)
        v['estabelecimentos'] = len(est) if est else 0
    return lst


def montar(notas, emp, locs=None):
    uf_emp = (emp.get('uf') or '').upper()
    vendas = {'bruto': 0.0, 'devol': 0.0, 'entrega_futura': 0.0, 'outras_saidas': 0.0, 'servicos': 0.0, 'notas_venda': 0, 'notas_devol': 0, 'notas_serv': 0}
    imp = defaultdict(float)
    clientes, fornecedores, forn_mp, forn_sn, prestadores = {}, {}, {}, {}, {}
    por_uf = defaultdict(float)
    prod_vend, prod_mp = {}, {}
    finalidade = defaultdict(float)
    compras_total, sn_credito = 0.0, 0.0
    sist = {'bruto': 0.0, 'devol': 0.0}
    for n in notas:
        d = n.doc
        if not d or d.get('cancelada') or d.get('denegada') or getattr(n, 'cancel_fonte', '') in ('sistema', 'denegada'):
            continue  # cancelada (evento, protocolo ou no sistema) e denegada ficam fora de tudo
        raw, nat = d['raw'], d.get('natureza', 'normal')
        if d['tipo'] == 'nfe' and d['mov'] == 'saida' and d['propria'] and nat == 'normal':
            v_venda = 0.0
            for it in raw['itens']:
                suf = it['CFOP'][1:]
                ic = it.get('icms') or {}
                vi = it['vProd'] - it['vDesc'] + it['vFrete'] + it['vSeg'] + it['vOutro'] + (it.get('ipi') or {}).get('vIPI', 0) + ic.get('vICMSST', 0)
                if suf in VENDA:
                    v_venda += vi
                    _acc(prod_vend, (it.get('NCM'), it['xProd'][:60]), it['xProd'][:60], it.get('NCM', ''), '', vi, d['chave'])
                    imp['icms_saidas'] += ic.get('vICMS', 0); imp['ipi_saidas'] += (it.get('ipi') or {}).get('vIPI', 0)
                    imp['st_saidas'] += ic.get('vICMSST', 0); imp['pis_saidas'] += (it.get('pis') or {}).get('vPIS', 0)
                    imp['cofins_saidas'] += (it.get('cofins') or {}).get('vCOFINS', 0)
                elif suf in SERVICO_NFE:
                    vendas['servicos_nfe'] = vendas.get('servicos_nfe', 0.0) + vi
                    _acc(clientes, d['part_cnpj'] or d['part_nome'], d['part_nome'], d['part_cnpj'], d['part_uf'], vi, d['chave'])
                elif suf in ENTREGA_FUTURA:
                    vendas['entrega_futura'] += vi
                else:
                    vendas['outras_saidas'] += vi
            if v_venda:
                vendas['bruto'] += v_venda; vendas['notas_venda'] += 1
                _acc(clientes, d['part_cnpj'] or d['part_nome'], d['part_nome'], d['part_cnpj'], d['part_uf'], v_venda, d['chave'])
                por_uf['Dentro do estado (%s)' % uf_emp if (d['part_uf'] or uf_emp) == uf_emp else (d['part_uf'] or 'EX')] += v_venda
                if n.lancs:
                    sist['bruto'] += n.lancs[0]['valor'] * (v_venda / d['valor'] if d['valor'] else 1)
        elif d['tipo'] == 'nfe' and nat == 'dev_venda':
            vendas['devol'] += d['valor']; vendas['notas_devol'] += 1
            if n.lancs:
                sist['devol'] += n.lancs[0]['valor']
        elif d['tipo'] == 'nfe' and d['mov'] == 'entrada' and not d['propria'] and nat == 'normal':
            compras_total += d['valor']
            _acc(fornecedores, d['part_cnpj'], d['part_nome'], d['part_cnpj'], d['part_uf'], d['valor'], d['chave'])
            if raw['emit'].get('crt') in ('1', '4'):
                _acc(forn_sn, d['part_cnpj'], d['part_nome'], d['part_cnpj'], d['part_uf'], d['valor'], d['chave'])
                sn_credito += sum((it.get('icms') or {}).get('vCredICMSSN', 0) for it in raw['itens'])
            # finalidade pelo CFOP lançado no sistema (proporcional às linhas); sem lançamento, pelo CFOP do fornecedor
            linhas = [(p['sis']['cfop'][1:], p['sis']['vc'], p.get('xml')) for p in n.pares if p.get('sis') and p['sis'].get('cfop')]
            if not linhas:
                linhas = [(it['CFOP'][1:], it['vProd'], None) for it in raw['itens']]
            tot_l = sum(v for _, v, _ in linhas) or 1
            for suf, v, _ in linhas:
                rot = next((r for r, sufs in FINALIDADE if suf in sufs), 'Outras entradas')
                finalidade[rot] += d['valor'] * v / tot_l
            if any(suf in MP for suf, _, _ in linhas):
                v_mp = d['valor'] * sum(v for suf, v, _ in linhas if suf in MP) / tot_l
                _acc(forn_mp, d['part_cnpj'], d['part_nome'], d['part_cnpj'], d['part_uf'], v_mp, d['chave'])
                for it in raw['itens']:
                    _acc(prod_mp, (it.get('NCM'), it['xProd'][:60]), it['xProd'][:60], it.get('NCM', ''), '', it['vProd'] * v_mp / (d['valor'] or 1), d['chave'])
        elif d['tipo'] == 'cte' and d['mov'] == 'frete':
            finalidade['Frete'] += d['valor']
        elif d['tipo'] == 'nfse' and d['mov'] == 'prestado':
            vendas['servicos'] += raw['vServ']; vendas['notas_serv'] += 1
            vendas['serv_liquido'] = vendas.get('serv_liquido', 0.0) + (raw.get('vLiq') or raw['vServ'])
            if n.lancs:
                sist['serv'] = sist.get('serv', 0.0) + n.lancs[0]['valor']
            if raw['issRetido']:
                imp['iss_retido_prestados'] += raw['vISS']
            else:
                imp['iss_prestados'] += raw['vISS']
            for k in ('irrf', 'csll', 'inss', 'pis', 'cofins'):
                imp['ret_sofridas'] += raw['ret'][k]
            _acc(clientes, d['part_cnpj'] or d['part_nome'], d['part_nome'], d['part_cnpj'], d['part_uf'], raw['vServ'], d['chave'] or d['numero'])
        elif d['tipo'] == 'nfse' and d['mov'] == 'tomado':
            _acc(prestadores, d['part_cnpj'], d['part_nome'], d['part_cnpj'], '', raw['vServ'], d['chave'] or d['numero'])
            if raw['issRetido']:
                imp['iss_retido_tomados'] += raw['vISS']
            for k in ('irrf', 'csll', 'inss', 'pis', 'cofins'):
                imp['ret_a_recolher'] += raw['ret'][k]
    bruto = round(vendas['bruto'], 2)
    liquido = round(bruto - vendas['devol'], 2)
    serv_nfe = round(vendas.get('servicos_nfe', 0.0), 2)
    serv_total = round(vendas['servicos'] + serv_nfe, 2)
    tot_imp_vendas = imp['icms_saidas'] + imp['ipi_saidas'] + imp['pis_saidas'] + imp['cofins_saidas'] + imp['iss_prestados'] + imp['iss_retido_prestados']
    loc = (locs or {}).get('linhas', [])
    loc_rec = round(sum(x['valor'] for x in loc if x['tipo'] == 'receita'), 2)
    loc_desp = round(sum(x['valor'] for x in loc if x['tipo'] == 'despesa'), 2)
    for x in loc:
        if x['tipo'] == 'receita':
            _acc(clientes, x['cnpj'] or x['participante'], x['participante'], x['cnpj'], '', x['valor'], 'loc%s' % x['linha'])
    receita_total = bruto + serv_total + loc_rec
    top_cli = _top(clientes)
    conc5 = round(sum(c['pct'] for c in top_cli[:5] if not c.get('demais')), 1)
    return {
        'faturamento': {'bruto': bruto, 'devolucoes': round(vendas['devol'], 2), 'liquido': liquido, 'servicos': serv_total,
                        'servicos_nfse': round(vendas['servicos'], 2), 'servicos_nfe': serv_nfe, 'servicos_liquido_receber': round(vendas.get('serv_liquido', 0.0) + serv_nfe, 2),
                        'sistema_servicos': round(sist.get('serv', 0.0), 2),
                        'locacoes': loc_rec, 'locacoes_despesa': loc_desp, 'notas_loc': len([x for x in loc if x['tipo'] == 'receita']),
                        'receita_liquida_total': round(liquido + serv_total + loc_rec, 2), 'entrega_futura': round(vendas['entrega_futura'], 2),
                        'outras_saidas': round(vendas['outras_saidas'], 2), 'notas_venda': vendas['notas_venda'], 'notas_devol': vendas['notas_devol'],
                        'notas_serv': vendas['notas_serv'], 'sistema_bruto': round(sist['bruto'], 2), 'sistema_devol': round(sist['devol'], 2)},
        'impostos': {k: round(v, 2) for k, v in imp.items()},
        'carga_efetiva': round(100 * tot_imp_vendas / receita_total, 2) if receita_total else 0,
        'clientes': _limpa(top_cli), 'concentracao_top5': conc5,
        'fornecedores': _limpa(_top(fornecedores)), 'fornecedores_mp': _limpa(_top(forn_mp)),
        'simples': {'qtd': len(forn_sn), 'valor': round(sum(v['valor'] for v in forn_sn.values()), 2),
                    'pct_compras': round(100 * sum(v['valor'] for v in forn_sn.values()) / compras_total, 1) if compras_total else 0,
                    'credito_destacado': round(sn_credito, 2), 'top': _limpa(_top(forn_sn, 8))},
        'compras_total': round(compras_total, 2),
        'finalidade': sorted([{'nome': k, 'valor': round(v, 2), 'pct': round(100 * v / (sum(finalidade.values()) or 1), 1)} for k, v in finalidade.items() if v > 0.005], key=lambda x: -x['valor']),
        'vendas_uf': sorted([{'nome': k, 'valor': round(v, 2), 'pct': round(100 * v / (bruto or 1), 1)} for k, v in por_uf.items()], key=lambda x: -x['valor']),
        'produtos_vendidos': _limpa(_top(prod_vend)), 'materias_primas': _limpa(_top(prod_mp)),
        'prestadores': _limpa(_top(prestadores, 8)),
    }
