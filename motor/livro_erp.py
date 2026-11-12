# -*- coding: utf-8 -*-
"""Livro fiscal do ERP (Registro de Apuração do ICMS e do IPI por CFOP, ex.: Sankhya) como fonte dos totais de produtos.

Quando o escritório fecha pelo livro do ERP, os XML do lote podem não ter todas as notas que o ERP escriturou.
Com o livro guardado na sessão (livro_erp.json), as linhas da apuração (valor contábil, base e valor de ICMS e IPI,
isentas e outras) passam a ser as do livro; as notas e os CST dos XML ficam como detalhe e a diferença vai em `xml`.
A Estratégia SCP lê as mesmas linhas (scp.extrair), então entradas e saídas de produtos seguem o livro também."""
import json
import os
import re

CAMPOS_ICMS = ('vc', 'bc_icms', 'icms', 'isentas', 'outras')
CAMPOS_IPI = ('bc_ipi', 'ipi')


def _num(t):
    return float(t.replace('.', '').replace(',', '.'))


def _linhas_pdf(caminho):
    import fitz
    d = fitz.open(caminho)
    return [x.strip() for p in d for x in p.get_text().split('\n') if x.strip()]


def ler_pdf(caminho):
    """Registro de Apuração do ICMS ou do IPI (layout Sankhya): por CFOP, valor contábil, base, imposto, isentas e outras.
    O CFOP vem depois dos cinco valores da linha. Devolve {'tipo': 'icms'|'ipi', 'cfops': {...}, 'resumo': {...}}."""
    L = _linhas_pdf(caminho)
    txt = ' '.join(L)
    tipo = 'ipi' if re.search(r'Registro de Apura[çc][ãa]o do IPI', txt) else 'icms' if re.search(r'Registro de Apura[çc][ãa]o do ICMS', txt) else None
    if not tipo:
        raise ValueError('não é um Registro de Apuração do ICMS ou do IPI')
    num = re.compile(r'^-?[\d.]+,\d{2}$')
    out, buf = {}, []
    for x in L:
        if num.match(x):
            buf.append(_num(x))
        elif re.fullmatch(r'[1-7]\d{3}', x) and len(buf) >= 5:
            vc, bc, imp, ise, outr = buf[-5:]
            out[x] = {'vc': vc, 'bc': bc, 'imp': imp, 'isentas': ise, 'outras': outr}
            buf = []
        elif not num.match(x) and not x[:1].isdigit():
            buf = []
    if not out:
        raise ValueError('nenhum CFOP encontrado no relatório')
    return {'tipo': tipo, 'cfops': out}


def montar(arquivos, fonte='Sankhya'):
    """Junta o livro de ICMS e o de IPI num livro só por CFOP."""
    icms = ipi = None
    for a in arquivos:
        r = ler_pdf(a)
        if r['tipo'] == 'icms':
            icms = r
        else:
            ipi = r
    if not icms:
        raise ValueError('falta o Registro de Apuração do ICMS')
    cf = {}
    for c, v in icms['cfops'].items():
        cf[c] = {'vc': v['vc'], 'bc_icms': v['bc'], 'icms': v['imp'], 'isentas': v['isentas'], 'outras': v['outras']}
    if ipi:
        for c, v in ipi['cfops'].items():
            x = cf.setdefault(c, {'vc': v['vc'], 'bc_icms': 0.0, 'icms': 0.0, 'isentas': 0.0, 'outras': 0.0})
            x['bc_ipi'], x['ipi'] = v['bc'], v['imp']
    return {'fonte': fonte, 'tem_ipi': bool(ipi), 'cfops': cf, 'arquivos': [os.path.basename(a) for a in arquivos]}


def aplicar(res, livro):
    """Troca as linhas da apuração pelas do livro (idempotente: guarda o que veio dos XML em `xml`)."""
    ap = (res or {}).get('apuracao')
    if not ap or not livro or not livro.get('cfops'):
        return res
    tem_ipi = livro.get('tem_ipi')
    por = {}
    for l in ap.get('linhas', []):
        if 'xml' not in l:
            l['xml'] = {k: l.get(k, 0) for k in CAMPOS_ICMS + CAMPOS_IPI}
        por[l['cfop']] = l
    difs = []
    for c, v in sorted(livro['cfops'].items()):
        l = por.get(c)
        if not l:
            l = {'cfop': c, 'desc': '', 'grupo': 'entrada' if c[0] in '123' else 'saida', 'docs': 0, 'st': 0.0, 'csts': [],
                 'xml': {k: 0.0 for k in CAMPOS_ICMS + CAMPOS_IPI}}
            ap['linhas'].append(l); por[c] = l
        for k in CAMPOS_ICMS + (CAMPOS_IPI if tem_ipi else ()):
            if k in v:
                l[k] = round(v[k], 2)
        l['fonte'] = livro.get('fonte') or 'ERP'
    for c, l in por.items():
        if c not in livro['cfops']:
            for k in CAMPOS_ICMS + (CAMPOS_IPI if tem_ipi else ()):
                l[k] = 0.0  # o livro não tem este CFOP: os XML ficam só como detalhe
            l['fonte'] = livro.get('fonte') or 'ERP'
        dv, di = round(l['vc'] - l['xml']['vc'], 2), round(l['icms'] - l['xml']['icms'], 2)
        dp = round(l.get('ipi', 0) - l['xml']['ipi'], 2) if tem_ipi else 0
        if abs(dv) > 0.05 or abs(di) > 0.05 or abs(dp) > 0.05:
            difs.append({'cfop': c, 'vc': dv, 'icms': di, 'ipi': dp})
    ap['linhas'].sort(key=lambda l: (l['grupo'] != 'entrada', l['cfop']))
    tot = {}
    for g in ('entrada', 'saida'):
        ls = [l for l in ap['linhas'] if l['grupo'] == g]
        tot[g] = {k: round(sum(l.get(k, 0) or 0 for l in ls), 2) for k in ('vc', 'bc_icms', 'icms', 'bc_ipi', 'ipi', 'isentas', 'outras', 'st')}
    ap['totais'] = tot
    ap['livro_erp'] = {'fonte': livro.get('fonte'), 'arquivos': livro.get('arquivos', []), 'tem_ipi': tem_ipi, 'diferencas': difs}
    return res


def carregar(sessao):
    p = os.path.join(sessao, 'livro_erp.json')
    try:
        return json.load(open(p, encoding='utf-8')) if os.path.exists(p) else None
    except Exception:
        return None
