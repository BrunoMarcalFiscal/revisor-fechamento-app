# -*- coding: utf-8 -*-
"""Devoluções, canceladas e inutilizadas.

- natureza(): classifica a NF-e (normal, devolução de venda, devolução de compra, complementar, ajuste).
- analisar(): confronta cada devolução com a nota original referenciada (refNFe): item, quantidade
  (acumulada entre devoluções), preço, ICMS/IPI/ST proporcionais, CFOP correspondente e valor total.
- numeracao(): inutilizações lançadas ou com nota autorizada e números pulados na série da empresa."""
from collections import defaultdict
from .util import norm, brl, TOL, raiz, numero_doc

ENT_DEV = {'201', '202', '203', '204', '208', '209', '210', '410', '411', '503', '553', '555', '660', '661', '662', '918', '919'}
SAI_DEV = {'201', '202', '208', '209', '210', '410', '411', '412', '413', '503', '553', '555', '556', '660', '661', '662'}
# CFOP da venda/compra original -> sufixo da devolução correspondente
DEV_DE = {'101': '201', '103': '201', '105': '201', '109': '201', '111': '201', '116': '201', '122': '201', '124': '201', '125': '201',
          '102': '202', '104': '202', '106': '202', '108': '202', '110': '202', '113': '202', '115': '202', '117': '202', '118': '202', '119': '202',
          '120': '201', '121': '202', '123': '202', '929': '202',
          '401': '410', '402': '410', '403': '411', '404': '411', '405': '411', '551': '553', '556': '556', '407': '556',
          '651': '660', '652': '661', '653': '662', '654': '661', '655': '661', '656': '661'}

ROTULO = {'normal': 'Normal', 'dev_venda': 'Devolução de venda', 'dev_compra': 'Devolução de compra',
          'dev_compra_forn': 'Devolução de compra (nota do fornecedor)', 'complementar': 'Complementar', 'ajuste': 'Ajuste',
          'cancelada': 'Cancelada', 'inutilizada': 'Inutilizada'}


def natureza(d, propria, raiz_emp):
    fin = d.get('finNFe')
    if fin == '2':
        return 'complementar'
    if fin == '3':
        return 'ajuste'
    cfops = [i['CFOP'] for i in d.get('itens', []) if i.get('CFOP')]
    ent = bool(cfops) and cfops[0][:1] in '123'
    por_cfop = bool(cfops) and all(c[1:] in (ENT_DEV if ent else SAI_DEV) for c in cfops)
    if ent and not propria:
        por_cfop = False  # 1556 de terceiro não é devolução
    if fin != '4' and not (por_cfop and (propria or ent)):
        return 'normal'
    if propria:
        return 'dev_venda' if d.get('tpNF') == '0' else 'dev_compra'
    return 'dev_venda' if d.get('tpNF') == '1' else 'dev_compra_forn'


def _vun(it):
    return (it['vProd'] / it['qCom']) if it.get('qCom') else it.get('vUnCom', 0)


def _imp(it):
    ic = it.get('icms') or {}
    ipi = (it.get('ipi') or {}).get('vIPI', 0) or it.get('vIPIDevol', 0)
    return ic.get('vICMS', 0), ipi, ic.get('vICMSST', 0), ic.get('pICMS', 0)


def _casar_item(it, itens_o, mesmo_emitente):
    if mesmo_emitente:
        for k, o in enumerate(itens_o):
            if o['cProd'] == it['cProd']:
                return k
    cands = [k for k, o in enumerate(itens_o) if o.get('NCM') and o['NCM'] == it.get('NCM')]
    if not cands:
        dx = norm(it['xProd'])[:25]
        cands = [k for k, o in enumerate(itens_o) if norm(o['xProd'])[:25] == dx]
    if not cands:
        return None
    return min(cands, key=lambda k: abs(_vun(itens_o[k]) - _vun(it)))


def analisar(notas, docs, ctx):
    """Acrescenta apontamentos às notas de devolução e devolve o painel de devoluções."""
    por_chave = {n.doc['chave']: n for n in notas if n.doc and n.doc.get('chave')}
    devs = [n for n in notas if n.doc and n.doc.get('natureza', 'normal').startswith('dev')]
    acum = defaultdict(float)       # (chave original, índice do item) -> quantidade devolvida
    total_dev = defaultdict(float)  # chave original -> valor devolvido
    painel = []
    for n in devs:
        d = n.doc
        raw = d['raw']
        nat = d['natureza']
        refs = [r for r in raw.get('refs', []) if r]
        orig = [por_chave[r] for r in refs if r in por_chave]
        reg = {'nota': None, 'natureza': nat, 'rotulo': ROTULO[nat], 'numero': d['numero'], 'emissao': d['emissao'],
               'part_nome': d['part_nome'], 'valor': d['valor'], 'originais': [], 'pct': None, 'problemas': 0, 'n': n}
        if not refs:
            n.ap('DEV_SEM_REF', 'A devolução não referencia a nota de origem (grupo NFref vazio).', '', '', 'Conferir de qual nota é a devolução.')
        elif not orig:
            n.ap('DEV_ORIGINAL_FORA', 'Nota de origem %s não está entre os XML enviados (provavelmente de mês anterior).' % ', '.join(r[25:34].lstrip('0') for r in refs),
                 '', '', 'Comparação item a item não feita.')
        for no in orig:
            o = no.doc
            reg['originais'].append({'numero': o['numero'], 'emissao': o['emissao'], 'valor': o['valor'], 'chave': o['chave']})
            total_dev[o['chave']] += d['valor']
        if orig:
            itens_o = [(no, k, it) for no in orig for k, it in enumerate(no.doc['raw']['itens'])]
            base = [it for _, _, it in itens_o]
            mesmo = all(no.doc['raw']['emit']['cnpj'] == raw['emit']['cnpj'] for no in orig) or \
                    (nat == 'dev_venda' and d['propria'] and all(no.doc['propria'] for no in orig))
            for it in raw['itens']:
                k = _casar_item(it, base, mesmo)
                if k is None:
                    n.ap('DEV_ITEM', '"%s" (NCM %s) não está na nota de origem %s.' % (it['xProd'][:50], it.get('NCM', ''), ', '.join(no.doc['numero'] for no in orig)),
                         'não consta', '%s × %s' % (_q(it['qCom']), brl(_vun(it))), 'Conferir o item devolvido.', valor=it['vProd'])
                    continue
                no, ko, o = itens_o[k]
                chave_o = no.doc['chave']
                acum[(chave_o, ko)] += it['qCom']
                if o['qCom'] and acum[(chave_o, ko)] > o['qCom'] + 1e-6:
                    n.ap('DEV_QTD', '"%s": devolvido %s no total, a nota %s tem %s.' % (it['xProd'][:40], _q(acum[(chave_o, ko)]), no.doc['numero'], _q(o['qCom'])),
                         _q(o['qCom']), _q(acum[(chave_o, ko)]), 'Conferir a quantidade devolvida.', valor=(acum[(chave_o, ko)] - o['qCom']) * _vun(o))
                vo, vd = _vun(o), _vun(it)
                if abs(vo - vd) > max(0.01, 0.001 * vo):
                    n.ap('DEV_PRECO', '"%s": preço unitário %s na devolução e %s na nota %s.' % (it['xProd'][:40], brl(vd), brl(vo), no.doc['numero']),
                         brl(vo), brl(vd), 'A devolução sai pelo mesmo valor da operação original.', valor=abs(vo - vd) * it['qCom'])
                prop = (it['qCom'] / o['qCom']) if o['qCom'] else 1
                io_, ipo, sto, po = _imp(o)
                id_, ipd, std, pd = _imp(it)
                for cod, nome, esp, real in (('DEV_ICMS', 'ICMS', io_ * prop, id_), ('DEV_IPI', 'IPI', ipo * prop, ipd), ('DEV_ST', 'ICMS-ST', sto * prop, std)):
                    if abs(esp - real) > max(TOL, 0.005 * esp):
                        extra = ' (alíquota %g%% × %g%%)' % (po, pd) if cod == 'DEV_ICMS' and po != pd else ''
                        n.ap(cod, '"%s": %s proporcional da nota %s é %s; na devolução %s%s.' % (it['xProd'][:40], nome, no.doc['numero'], brl(esp), brl(real), extra),
                             brl(esp), brl(real), 'Destacar %s de %s, proporcional ao devolvido.' % (nome, brl(esp)), valor=abs(esp - real))
                # CFOP da devolução própria x CFOP da operação original
                if d['propria'] and len(it['CFOP']) == 4:
                    suf_o = o['CFOP'][1:]
                    if nat == 'dev_compra':
                        suf_o = _cfop_lancado(no, ko) or suf_o
                    esp = DEV_DE.get(suf_o)
                    if esp and it['CFOP'][1:] != esp:
                        n.ap('DEV_CFOP', 'Operação original %s%s; a devolução correspondente é %s%s.' % (o['CFOP'][0] if nat == 'dev_venda' else _pri_ent(no), suf_o, it['CFOP'][0], esp),
                             '%s%s' % (it['CFOP'][0], esp), it['CFOP'], 'Emitir/lançar com %s%s.' % (it['CFOP'][0], esp))
            # devolução de venda emitida pelo cliente: o CFOP certo do lançamento sai da venda original
            if nat == 'dev_venda' and not d['propria']:
                suf_o = orig[0].doc['raw']['itens'][0]['CFOP'][1:]
                esp = DEV_DE.get(suf_o)
                for idx, p in enumerate(n.pares):
                    sc = (p.get('sis') or {}).get('cfop', '')
                    ja = any(a['par'] == idx and a['cod'].startswith('CFOP') for a in n.aponts)
                    if esp and len(sc) == 4 and sc[1:] != esp and not ja:
                        n.ap('DEV_CFOP', 'A venda original foi %s; a devolução se lança em %s%s.' % (orig[0].doc['raw']['itens'][0]['CFOP'], sc[0], esp),
                             sc[0] + esp, sc, 'Lançar com %s%s.' % (sc[0], esp), idx, 'cfop')
        painel.append(reg)
    # devolução maior que a original (somando todas as devoluções da mesma nota)
    for reg in painel:
        n = reg['n']
        for o in reg['originais']:
            if total_dev[o['chave']] > o['valor'] + TOL:
                n.ap('DEV_VALOR_MAIOR', 'Devoluções da nota %s somam %s; a nota é de %s.' % (o['numero'], brl(total_dev[o['chave']]), brl(o['valor'])),
                     o['valor'], total_dev[o['chave']], 'Conferir as devoluções desta nota.', valor=total_dev[o['chave']] - o['valor'])
        if reg['originais']:
            tot = sum(o['valor'] for o in reg['originais'])
            reg['pct'] = round(100 * reg['valor'] / tot, 1) if tot else None
    _integral(notas, por_chave, painel, total_dev)
    # duas devoluções iguais (mesmo participante, mesma natureza, mesmo valor)
    grupos = defaultdict(list)
    for reg in painel:
        if getattr(reg['n'], 'cancel_fonte', '') in ('sistema', 'denegada') or reg['n'].doc.get('cancelada'):
            continue  # a cancelada não conta como repetição
        grupos[(reg['natureza'], reg['n'].doc['part_cnpj'], round(reg['valor'], 2))].append(reg)
    for regs in grupos.values():
        if len(regs) > 1:
            nums = ', '.join(r['numero'] for r in regs)
            for r in regs:
                r['n'].ap('DEV_REPETIDA', 'Devoluções %s têm o mesmo participante e o mesmo valor (%s). Confira se uma foi cancelada ou substituiu a outra.' % (nums, brl(r['valor'])),
                          brl(r['valor']), nums, 'Manter só a devolução válida; mandar o cancelamento da outra.')
    # devolução feita pelo fornecedor (nota de entrada dele) e não lançada: conferir, não é erro certo
    for reg in painel:
        n = reg['n']
        if reg['natureza'] == 'dev_compra_forn' and not n.lancs and any(a['cod'] == 'NAO_LANCADA' for a in n.aponts):
            n.aponts = [a for a in n.aponts if a['cod'] != 'NAO_LANCADA']
            n.ap('DEV_FORN_NAO_LANCADA', 'O próprio fornecedor emitiu nota de entrada para receber a mercadoria de volta. Confira se a compra de origem foi lançada e se a devolução precisa entrar no livro.',
                 brl(n.doc['valor']), 'não lançada', 'Se a compra está no livro, lançar a devolução.', valor=n.doc['valor'])
    return painel


def _cfop_lancado(no, k):
    """CFOP com que a compra original foi lançada no sistema (pelo par da linha)."""
    for p in no.pares:
        if p.get('sis') and p['sis'].get('cfop'):
            return p['sis']['cfop'][1:]
    return ''


def _pri_ent(no):
    for p in no.pares:
        if p.get('sis') and p['sis'].get('cfop'):
            return p['sis']['cfop'][0]
    return '1'


def _q(v):
    return ('%g' % v).replace('.', ',')


def _integral(notas, por_chave, painel, total_dev):
    """Compra devolvida por inteiro (pelo fornecedor ou pela empresa): se nada foi lançado, não é erro; se só uma ponta foi, é."""
    for reg in painel:
        n = reg['n']
        if reg['natureza'] not in ('dev_compra', 'dev_compra_forn') or len(reg['originais']) != 1:
            continue
        o = reg['originais'][0]
        if abs(total_dev[o['chave']] - o['valor']) > TOL:
            continue
        no = por_chave[o['chave']]
        o_lanc, d_lanc = bool(no.lancs), bool(n.lancs)
        if not o_lanc and not d_lanc:
            for x, outro in ((no, n), (n, no)):
                x.aponts = [a for a in x.aponts if a['cod'] != 'NAO_LANCADA']
                x.ap('DEVOLVIDA_INTEGRAL', 'Compra %s devolvida por inteiro na nota %s. Nenhuma das duas está lançada: efeito zero no livro.' % (no.doc['numero'], n.doc['numero']),
                     brl(o['valor']), 'nenhuma lançada', 'Se for lançar, lance as duas (compra e devolução).')
        elif o_lanc and not d_lanc:
            n.aponts = [a for a in n.aponts if a['cod'] != 'NAO_LANCADA']
            n.ap('DEV_NAO_LANCADA', 'A compra %s foi lançada, mas a devolução integral não.' % no.doc['numero'], brl(o['valor']), 'não lançada',
                 'Lançar a devolução %s.' % n.doc['numero'], valor=o['valor'])


def numeracao(notas, docs, inuts, lancs, emp, Nota, competencia=''):
    """Inutilizações e buracos na numeração própria. Devolve (painel, notas novas)."""
    novas, faixas = [], []
    proprias = defaultdict(set)
    for d in docs:
        if d['tipo'] == 'nfe' and d['propria'] and d.get('modelo') in ('55', '65') and (not competencia or d['emissao'][:7] == competencia):
            proprias[(d['modelo'], str(int(d['serie'] or 0)))].add(int(d['numero'] or 0))
    nota_por_num = {(n.doc['modelo'], str(int(n.doc['serie'] or 0)), int(n.doc['numero'] or 0)): n for n in notas
                    if n.doc and n.doc['tipo'] == 'nfe' and n.doc['propria']}
    inut_set = defaultdict(set)
    for u in inuts:
        if raiz(u['cnpj']) != emp['raiz']:
            continue
        serie = str(int(u['serie'] or 0))
        nums = set(range(u['ini'], u['fim'] + 1))
        inut_set[(u['modelo'], serie)] |= nums
        reg = {'modelo': u['modelo'], 'serie': serie, 'ini': u['ini'], 'fim': u['fim'], 'just': u['just'], 'dh': u['dh'], 'problemas': []}
        for num in sorted(nums):
            if num in proprias.get((u['modelo'], serie), set()):
                n = nota_por_num.get((u['modelo'], serie, num))
                if n:
                    n.ap('INUTILIZADA_COM_NOTA', 'O número %s da série %s foi inutilizado, mas há NF-e autorizada com ele.' % (num, serie), 'inutilizado', 'autorizada',
                         'Conferir na SEFAZ: inutilização ou nota, só pode existir um.')
                reg['problemas'].append('%s com nota autorizada' % num)
            for L in lancs:
                if L['numero'] == str(num) and (not L['serie'] or L['serie'].lstrip('0') == serie) and L.get('mov_hint') in ('saida', '') and \
                        (not L['chave'] or raiz(L['chave'][6:20]) == emp['raiz']) and not L.get('usado'):
                    n = Nota('saida', None, [L])
                    n.ap('INUTILIZADA_LANCADA', 'O número %s da série %s foi inutilizado e está lançado na planilha.' % (num, serie), 'inutilizado', L['valor'],
                         'Excluir o lançamento.', valor=L['valor'])
                    L['usado'] = True
                    novas.append(n)
                    reg['problemas'].append('%s lançado' % num)
        faixas.append(reg)
    pulos = []
    for (mod, serie), nums in proprias.items():
        if len(nums) < 2:
            continue
        cobertos = nums | inut_set[(mod, serie)]
        faltam = [x for x in range(min(nums), max(nums) + 1) if x not in cobertos]
        for ini, fim in _faixas(faltam)[:50]:
            n = Nota('saida', None, [])
            n.numeracao = {'modelo': mod, 'serie': serie, 'ini': ini, 'fim': fim}
            txt = ('%s' % ini) if ini == fim else ('%s a %s' % (ini, fim))
            n.ap('NUMERO_PULADO', 'Série %s (modelo %s): número %s não veio em XML, cancelamento nem inutilização.' % (serie, mod, txt), 'sem documento', '',
                 'Mandar o XML (ou o cancelamento/inutilização) desse número.')
            novas.append(n)
            pulos.append({'modelo': mod, 'serie': serie, 'ini': ini, 'fim': fim})
    return {'inutilizadas': faixas, 'pulos': pulos}, novas


def _faixas(nums):
    out = []
    for x in sorted(nums):
        if out and x == out[-1][1] + 1:
            out[-1][1] = x
        else:
            out.append([x, x])
    return [tuple(f) for f in out]
