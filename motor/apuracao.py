# -*- coding: utf-8 -*-
"""Apuração de ICMS e IPI a partir da planilha do sistema (o livro como o cliente lançou).

- Quadro por CFOP (Valor contábil, Base ICMS, ICMS, Base IPI, IPI, Isentas, Outras), aberto por CST.
  Quando a planilha não tem CST, Base IPI, Isentas ou Outras, o valor vem do XML casado com a linha
  ou é calculado (Isentas/Outras = Valor contábil − Base ICMS) e fica marcado como calculado.
- Ajustes sugeridos: crédito do Simples Nacional (pCredSN do XML ou texto do rodapé), DIFAL de uso e
  consumo (2556), estornos de crédito e de débito tirados dos apontamentos.
O cálculo final (débito × crédito, saldo) é feito na tela, onde os ajustes são editados, e salvo em sessao/apuracao.json."""
import re
from collections import defaultdict
from .util import num, brl, raiz, digitos, TOL
from .regras_base import desc_cfop, CST_DESC

ALIQ_INTERNA_UF = {  # referência ano-base 2025 (mesma tabela do Revisor Tributário); editável na tela
    'AC': 19.0, 'AL': 20.0, 'AM': 20.0, 'AP': 18.0, 'BA': 20.5, 'CE': 20.0, 'DF': 20.0, 'ES': 17.0, 'GO': 19.0, 'MA': 23.0,
    'MG': 18.0, 'MS': 17.0, 'MT': 17.0, 'PA': 19.0, 'PB': 20.0, 'PE': 20.5, 'PI': 22.5, 'PR': 19.5, 'RJ': 20.0, 'RN': 20.0,
    'RO': 19.5, 'RR': 20.0, 'RS': 17.0, 'SC': 17.0, 'SE': 20.0, 'SP': 18.0, 'TO': 20.0,
}
UF_SUL_SUDESTE = {'PR', 'SC', 'RS', 'SP', 'RJ', 'MG'}
CAMPOS = ('vc', 'bc_icms', 'icms', 'bc_ipi', 'ipi', 'isentas', 'outras', 'st')
ISENTAS_CST = {'40', '41', '30', '102', '103', '300', '400'}
RE_SN = re.compile(r'(?:aproveitamento|aproveitar|apropria[çc][ãa]o)[^.;]{0,120}?cr[ée]dito[^.;]{0,120}?R\$\s*([\d.]+,\d{2})', re.I)
RE_SN_ALIQ = re.compile(r'(\d{1,2}[,.]\d{1,2})\s*%')


def aliq_interestadual(uf_o, uf_d):
    if uf_o in UF_SUL_SUDESTE and uf_d not in UF_SUL_SUDESTE:
        return 7.0
    return 12.0


def _vazio():
    return {k: 0.0 for k in CAMPOS}


def montar(notas, regs, emp, cols_por_aba):
    uf = (emp.get('uf') or '').upper()
    # ---- linha da planilha -> dados do XML casado (CST, base do IPI, ICMS de origem, alíquota)
    da_linha = {}
    for n in notas:
        if not n.lancs or not n.doc:
            continue
        L = n.lancs[0]
        for p in n.pares:
            x, s = p.get('xml'), p.get('sis')
            if not x or not s or not s.get('linhas'):
                continue
            tot = s.get('vc') or 0
            for r in L['linhas']:
                if r['linha'] in s['linhas']:
                    share = (r.get('valor_contabil', 0) / tot) if tot else 1 / len(s['linhas'])
                    cst = x.get('cst_n') or ''
                    da_linha[(r['arquivo'], r['aba'], r['linha'])] = {
                        'cst': (x.get('orig', '') + cst) if cst and len(cst) == 2 else cst,
                        'bc_ipi': round((x.get('bc_ipi') or 0) * share, 2), 'icms_xml': round((x.get('icms') or 0) * share, 2),
                        'paliq': x.get('paliq') or 0, 'nota': n, 'uf_part': (n.doc.get('part_uf') or '').upper()}
    # ---- quadro por CFOP / CST
    por_cfop = defaultdict(lambda: {'tot': _vazio(), 'csts': defaultdict(_vazio), 'docs': set(), 'origem_cst': set(), 'nfs': defaultdict(set)})
    deriv = {'cst': False, 'bc_ipi': False, 'isentas_outras': False}
    for r in regs:
        cf = (r.get('cfop') or '')[:4]
        if len(cf) != 4 or cf[0] not in '123567':
            continue
        cols = cols_por_aba.get((r['arquivo'], r['aba']), set())
        x = da_linha.get((r['arquivo'], r['aba'], r['linha']), {})
        cst = digitos(r.get('cst_icms'))
        if cst:
            cst = cst[-3:] if len(cst) >= 3 else cst.zfill(2)
        elif x.get('cst'):
            cst = x['cst']; deriv['cst'] = True
        v = {'vc': r.get('valor_contabil', 0), 'bc_icms': r.get('base_icms', 0), 'icms': r.get('valor_icms', 0),
             'ipi': r.get('valor_ipi', 0), 'st': r.get('valor_st', 0)}
        if 'base_ipi' in cols:
            v['bc_ipi'] = r.get('base_ipi', 0)
        elif v['ipi'] and x:
            v['bc_ipi'] = x.get('bc_ipi', 0); deriv['bc_ipi'] = True
        else:
            v['bc_ipi'] = 0.0
        if 'isentas' in cols or 'outras' in cols:
            v['isentas'], v['outras'] = r.get('isentas', 0), r.get('outras', 0)
        else:
            resto = round(max(v['vc'] - v['bc_icms'], 0), 2)
            trib = cst[-2:] if cst and cst not in ISENTAS_CST else cst
            if cst in ISENTAS_CST or trib in ISENTAS_CST:
                v['isentas'], v['outras'] = resto, 0.0
            else:
                v['isentas'], v['outras'] = 0.0, resto
            deriv['isentas_outras'] = True
        g = por_cfop[cf]
        for k in CAMPOS:
            g['tot'][k] += v.get(k, 0)
            g['csts'][cst or 'sem CST'][k] += v.get(k, 0)
        g['docs'].add((r.get('chave') or '') or (r.get('numero'), r.get('cnpj')))
        nf = (r.get('numero') or '').lstrip('0') or ((r.get('chave') or '')[25:34].lstrip('0') if len(r.get('chave') or '') == 44 else '')
        if nf:
            g['nfs'][cst or 'sem CST'].add(nf)
    linhas = []
    for cf in sorted(por_cfop, key=lambda c: (c[0] not in '123', c)):
        g = por_cfop[cf]
        linhas.append({'cfop': cf, 'desc': desc_cfop(cf), 'grupo': 'entrada' if cf[0] in '123' else 'saida', 'docs': len(g['docs']),
                       **{k: round(g['tot'][k], 2) for k in CAMPOS},
                       'csts': [{'cst': c, 'desc': CST_DESC.get(c[-2:] if len(c) == 3 and c not in ISENTAS_CST else c, ''),
                                 'notas': sorted(g['nfs'].get(c, set()), key=lambda x: int(x) if x.isdigit() else 0),
                                 **{k: round(vv[k], 2) for k in CAMPOS}} for c, vv in sorted(g['csts'].items())]})
    totais = {}
    for grp in ('entrada', 'saida'):
        totais[grp] = {k: round(sum(l[k] for l in linhas if l['grupo'] == grp), 2) for k in CAMPOS}

    # ---- crédito do Simples Nacional destacado na nota (pCredSN / rodapé)
    simples = []
    for n in notas:
        d = n.doc
        if not d or d['tipo'] != 'nfe' or d['mov'] != 'entrada' or d['propria'] or d['raw']['emit'].get('crt') not in ('1', '4') or d.get('cancelada'):
            continue
        raw = d['raw']
        cred = round(sum((it.get('icms') or {}).get('vCredICMSSN', 0) for it in raw['itens']), 2)
        aliq = max([(it.get('icms') or {}).get('pCredSN', 0) for it in raw['itens']] or [0])
        origem = 'XML (pCredSN)'
        if cred <= TOL:
            m = RE_SN.search(raw.get('infAdic') or '')
            if m:
                cred = num(m.group(1)); origem = 'rodapé da nota'
                ma = RE_SN_ALIQ.search(raw['infAdic'][max(0, m.start() - 80):m.end() + 80])
                aliq = num(ma.group(1)) if ma else 0
        if cred <= TOL:
            continue
        lanc_icms = round(sum(r.get('valor_icms', 0) for r in n.lancs[0]['linhas']), 2) if n.lancs else 0
        simples.append({'id': getattr(n, '_id', None), 'numero': d['numero'], 'fornecedor': d['part_nome'], 'cnpj': d['part_cnpj'], 'emissao': d['emissao'],
                        'valor_nota': d['valor'], 'aliq': aliq, 'credito': cred, 'origem': origem, 'lancada': bool(n.lancs),
                        'ja_no_livro': bool(n.lancs) and lanc_icms >= cred - TOL, 'icms_no_livro': lanc_icms,
                        'cfop': ','.join(sorted({r.get('cfop', '') for r in n.lancs[0]['linhas']})) if n.lancs else '',
                        # só dá crédito a compra para comercialização/industrialização; uso e consumo e ativo (CIAP) não
                        'sem_direito': bool(n.lancs) and all((r.get('cfop') or '')[1:] in ('556', '407', '551', '406', '653', '949', '910', '911', '915', '916', '902', '903', '925')
                                                              for r in n.lancs[0]['linhas'])})

    # ---- DIFAL de uso e consumo (2556)
    difal = []
    por_nota = defaultdict(lambda: {'vc': 0.0, 'icms_xml': 0.0, 'paliq': 0.0, 'nota': None, 'linhas': [], 'uf': '', 'r': None})
    for r in regs:
        if (r.get('cfop') or '')[:4] != '2556':
            continue
        x = da_linha.get((r['arquivo'], r['aba'], r['linha']), {})
        k = (r.get('chave') or '') or (r['arquivo'], r.get('numero'), r.get('cnpj'))
        e = por_nota[k]
        e['vc'] += r.get('valor_contabil', 0); e['linhas'].append(r['linha']); e['r'] = e['r'] or r
        if x:
            e['icms_xml'] += x.get('icms_xml', 0); e['paliq'] = e['paliq'] or x.get('paliq', 0); e['nota'] = x['nota']; e['uf'] = x.get('uf_part', '')
    for k, e in por_nota.items():
        r, n = e['r'], e['nota']
        uf_o = e['uf'] or (r.get('uf') or '').upper()
        aliq_inter = e['paliq'] or (aliq_interestadual(uf_o, uf) if uf_o else 12.0)
        icms_o = round(e['icms_xml'], 2) if n else round(e['vc'] * aliq_inter / 100, 2)
        difal.append({'id': getattr(n, '_id', None) if n else None, 'numero': n.doc['numero'] if n else (r.get('numero') or ''),
                      'fornecedor': (n.doc['part_nome'] if n else r.get('participante', '')), 'uf': uf_o, 'valor': round(e['vc'], 2),
                      'aliq_inter': aliq_inter, 'icms_origem': icms_o, 'fonte': 'XML' if n else 'planilha (alíquota estimada pela UF)',
                      'linhas': e['linhas']})

    # ---- sugestões de estorno
    est_cred, est_deb = [], []
    for r in regs:
        cf = (r.get('cfop') or '')[:4]
        if cf[1:] in ('556', '407') and cf[0] in '123' and r.get('valor_icms', 0) > TOL:
            est_cred.append({'valor': round(r['valor_icms'], 2), 'descricao': 'Estorno de crédito de ICMS de uso e consumo (CFOP %s), NF %s, planilha linha %s' % (cf, r.get('numero') or (r.get('chave') or '')[25:34].lstrip('0'), r['linha']),
                             'motivo': 'Uso e consumo não dá crédito (LC 87/96, art. 33, I). Use se o crédito não for corrigido no documento.'})
        if cf[1:] == '556' and cf[0] in '567' and r.get('valor_icms', 0) > TOL:
            est_deb.append({'valor': round(r['valor_icms'], 2), 'descricao': 'Estorno de débito: devolução de uso e consumo (CFOP %s) cuja compra entrou sem crédito, NF %s' % (cf, r.get('numero') or (r.get('chave') or '')[25:34].lstrip('0')),
                            'motivo': 'O débito da devolução anula um crédito que não foi tomado na compra. Confirme a entrada sem crédito.'})
    for n in notas:
        for a in n.aponts:
            if a['cod'] == 'CREDITO_MAIOR' and n.doc and n.doc['mov'] in ('entrada', 'frete'):
                est_cred.append({'valor': round(a['valor'], 2), 'descricao': 'Estorno do crédito acima do destacado, NF %s (%s)' % (n.doc['numero'], n.doc['part_nome'][:30]),
                                 'motivo': 'Crédito limitado ao imposto destacado (LC 87/96, art. 23).'})
            if a['cod'] == 'CANCELADA_LANCADA' and n.doc and n.doc['mov'] == 'saida':
                est_deb.append({'valor': round(sum(r.get('valor_icms', 0) for r in n.lancs[0]['linhas']), 2) if n.lancs else 0,
                                'descricao': 'Estorno de débito de NF %s cancelada e lançada com valor' % n.doc['numero'], 'motivo': 'Nota cancelada não gera débito.'})
    est_cred = [e for e in est_cred if e['valor'] > TOL]
    est_deb = [e for e in est_deb if e['valor'] > TOL]
    pref = (uf or 'UF') + '0'
    return {'linhas': linhas, 'totais': totais, 'derivados': deriv, 'simples': simples, 'difal': difal,
            'sugestoes': {'estorno_credito': est_cred, 'estorno_debito': est_deb},
            'uf': uf, 'aliq_interna': ALIQ_INTERNA_UF.get(uf, 18.0),
            'prefixos': {'simples': pref + '2', 'difal': pref + '0', 'estorno_credito': pref + '1', 'estorno_debito': pref + '3'},
            'tem_planilha': bool(linhas)}
