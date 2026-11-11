# -*- coding: utf-8 -*-
"""Apuração de PIS/COFINS não cumulativo (Lucro Real) a partir dos XML do mês: débitos por natureza, créditos por natureza,
estornos, saldo credor e valor a recolher.

Regras:
- Base da venda: vProd - vDesc + frete + seguro + outras - ICMS próprio (IPI e ICMS-ST fora).
- Monofásico (Lei 10.147/00: 33.03 a 33.07, 3401.11.90, 3401.20.10, 9603.21.00): venda de produção a 2,20% / 10,30%;
  revenda a alíquota zero; compra para revenda sem crédito (art. 3º, I, b; STJ Tema 1093).
- Industrialização por encomenda de 33.03-33.07: alíquota zero para a executora (Lei 11.051/04, art. 10, §3º), sem crédito
  para a encomendante.
- Livro (49.01): alíquota zero (Lei 10.865/04, art. 28, VI).
- Base do crédito: vProd + frete + seguro + outras - desconto - ICMS + ICMS-ST, mais o IPI quando ele não é recuperado.
- CT-e: (vTPrest - ICMS) x fração creditável da carga (na venda, a parte que gerou débito; na compra, a parte com crédito).
- Classificação das entradas: primeiro a decisão já tomada para o mesmo fornecedor e produto (memória), depois as regras
  pelo CFOP/CST/NCM. O que não se encaixa fica "A classificar", sem crédito, até alguém decidir na tela.
"""
import json, os, re
from collections import defaultdict

PIS, COF = 1.65, 7.6
PIS_MONO, COF_MONO = 2.2, 10.3
NCM_MONO = ('3303', '3304', '3305', '3306', '3307', '34011190', '34012010', '96032100')
NCM_LIVRO = ('4901',)

VENDA_PROD = {'101', '103', '105', '107', '109', '111', '113', '116', '118', '122', '401', '402'}
REVENDA = {'102', '104', '106', '108', '110', '112', '114', '115', '117', '119', '120', '123', '403', '404', '405',
           '651', '652', '653', '654', '655', '656', '667'}
INDUSTRIALIZACAO = {'124', '125'}
BONIFICACAO = {'910', '911'}
DEV_COMPRA = {'201', '202', '208', '209', '210', '410', '411', '412', '413', '503', '553', '555', '556', '660', '661', '662'}
DEV_VENDA = {'201', '202', '203', '204', '208', '209', '410', '411', '660', '661', '662'}   # vistos da entrada (1xxx/2xxx)
ATIVO = {'551', '552'}
SERVICO_NFE = {'933', '932'}
COMPRA = VENDA_PROD | REVENDA | {'126', '127', '401', '403', '405', '651', '652', '653', '654', '655', '656', '667'}
CST_FORN_SEM = {'04', '05', '06', '07', '08', '09'}
# entradas que não são aquisição, pelo CFOP do emitente (vêm antes da memória: o mesmo produto comprado de verdade dá crédito)
NAO_AQUISICAO = {'901': 'Remessa p/ industrialização recebida', '924': 'Remessa p/ industrialização por conta e ordem',
                 '902': 'Retorno de industrialização', '903': 'Retorno de industrialização não aplicada', '925': 'Retorno de industrialização',
                 '910': 'Bonificação/doação/brinde', '911': 'Amostra grátis', '912': 'Demonstração', '913': 'Retorno de demonstração',
                 '915': 'Remessa p/ conserto', '916': 'Retorno de conserto', '908': 'Comodato', '909': 'Comodato',
                 '920': 'Vasilhame/sacaria', '921': 'Retorno de vasilhame', '922': 'Simples faturamento (entrega futura)',
                 '923': 'Remessa por conta e ordem de terceiros', '929': 'Lançamento vinculado a cupom fiscal',
                 '905': 'Remessa p/ depósito', '906': 'Retorno de depósito', '907': 'Retorno de depósito', '914': 'Remessa p/ feira/exposição',
                 '917': 'Consignação', '918': 'Devolução de consignação', '919': 'Devolução de consignação', '934': 'Remessa simbólica p/ depósito'}

# naturezas que a pessoa escolhe na tela para uma entrada "A classificar": (credito, motivo, ipi_recuperavel)
OPCOES_ENTRADA = {
    'Insumo - matéria-prima': ('sim', 'Matéria-prima (art. 3º, II)', True),
    'Insumo - embalagem': ('sim', 'Embalagem/rótulo do produto: insumo (art. 3º, II)', True),
    'Insumo - manutenção de máquinas': ('sim', 'Peças/lubrificantes de máquinas da produção: insumo (PN 5/2018)', True),
    'Insumo - EPI da produção': ('sim', 'EPI exigido na produção: insumo (PN 5/2018)', True),
    'Revenda - tributada': ('sim', 'Mercadoria para revenda fora do monofásico (art. 3º, I)', False),
    'Revenda - monofásico': ('nao', 'Revenda de monofásico (Lei 10.147/00): vedado (art. 3º, I, b; STJ Tema 1093)', False),
    'Ativo imobilizado - produção': ('sim', 'Máquina da produção: crédito imediato (Lei 12.546/11, art. 4º)', True),
    'Uso e consumo': ('nao', 'Material de uso e consumo: sem crédito', False),
    'Uso e consumo - refeição': ('nao', 'Alimentação de empregados: não é insumo (PN 5/2018)', False),
    'Remessa/retorno/outras': ('nao', 'Remessa/retorno/outra entrada: não é aquisição', False),
}
OPCOES_SERVICO = {
    'Insumo - serviço': ('Crédito', 'Serviço aplicado na produção: insumo (art. 3º, II; PN 5/2018)', '03'),
    'Aluguel de máquinas e equipamentos (NFS-e)': ('Crédito', 'Aluguel de máquinas e equipamentos pago a PJ (art. 3º, IV)', '06'),
    'Aluguel de prédio (NFS-e)': ('Crédito', 'Aluguel de prédio pago a PJ (art. 3º, IV)', '05'),
    'Frete/armazenagem na venda (NFS-e)': ('Crédito', 'Frete/armazenagem na venda (art. 3º, IX), na parte tributada das vendas do mês', '07'),
    'Serviço administrativo/predial': ('Sem crédito', 'Não é insumo da produção', ''),
}
IPI_RECUPERAVEL = {k for k, v in OPCOES_ENTRADA.items() if v[2]} | {'Industrialização por encomenda (insumo)'}


def _r(v):
    return round(v + 0.0, 2)


def mono(ncm):
    n = re.sub(r'\D', '', ncm or '')
    return n.startswith(NCM_MONO)


def livro(ncm):
    return re.sub(r'\D', '', ncm or '').startswith(NCM_LIVRO)


def _raiz(c):
    return re.sub(r'\D', '', c or '')[:8]


# ------------------------------------------------------------------ memória de classificação
def ler_memoria(arq):
    try:
        return json.load(open(arq, encoding='utf-8'))
    except Exception:
        return {}


def gravar_decisao(arq, raiz, tipo, chave, valor):
    mem = ler_memoria(arq)
    mem.setdefault(raiz, {}).setdefault(tipo, {})[chave] = valor
    os.makedirs(os.path.dirname(arq), exist_ok=True)
    json.dump(mem, open(arq, 'w', encoding='utf-8'), ensure_ascii=False)


# ------------------------------------------------------------------ bases
def base_venda(it):
    return it['vProd'] - it['vDesc'] + it['vFrete'] + it['vSeg'] + it['vOutro'] - (it.get('icms') or {}).get('vICMS', 0)


def base_credito(it, com_ipi=True):
    ic = it.get('icms') or {}
    b = it['vProd'] + it['vFrete'] + it['vSeg'] + it['vOutro'] - it['vDesc'] - ic.get('vICMS', 0) + ic.get('vICMSST', 0) + ic.get('vFCPST', 0)
    if com_ipi:
        b += (it.get('ipi') or {}).get('vIPI', 0)
    return b


# ------------------------------------------------------------------ classificação das saídas (nota própria)
def nome_produto(x):
    return re.split(r'\s+(?:Pedido|Lote)\s*:', x or '')[0].strip().upper()


def classificar_saida(it, fabricante, mem=None):
    c3 = (it.get('CFOP') or '')[1:]
    ncm = it.get('NCM')
    if (it.get('CFOP') or '').startswith('7'):
        return 'Exportação', 'export', 'Exportação: não incide (crédito mantido)'
    if c3 in SERVICO_NFE:
        return 'Prestação de serviço', 'trib', '1,65% / 7,6%'
    if c3 in INDUSTRIALIZACAO:
        if mono(ncm):
            return 'Industrialização p/ terceiros - monofásico', 'zero', 'Alíquota zero (Lei 11.051/04, art. 10, §3º)'
        return 'Industrialização p/ terceiros', 'trib', '1,65% / 7,6%'
    if c3 in VENDA_PROD or c3 in REVENDA:
        if livro(ncm):
            return 'Venda de livro', 'zero', 'Alíquota zero (Lei 10.865/04, art. 28, VI)'
        if mono(ncm):
            ms = ((mem or {}).get('saidas') or {}).get(nome_produto(it.get('xProd')))
            if ms:   # produto comprado pronto e vendido com CFOP de produção: decisão já tomada
                return ms[0], ms[1], ms[2]
            if c3 in VENDA_PROD and fabricante:
                return 'Venda de monofásico (fabricante)', 'mono', 'Monofásico fabricante: 2,20% / 10,30% (Lei 10.147/00)'
            return 'Revenda de monofásico', 'zero', 'Alíquota zero'
        return 'Venda tributada', 'trib', '1,65% / 7,6%'
    if c3 in BONIFICACAO:
        return 'Bonificação/amostra', 'nada', 'Sem receita'
    if c3 in DEV_COMPRA:
        return 'Devolução de compra', 'estorno', ''
    if c3 in ATIVO:
        return 'Venda de ativo imobilizado', 'nada', 'Fora da base (art. 1º, §3º, VI)'
    if c3 == '603':
        return 'Ressarcimento de ICMS-ST', 'nada', ''
    return 'Remessa/retorno/outras saídas', 'nada', ''


def _aliq_venda(tipo):
    return {'mono': (PIS_MONO, COF_MONO), 'trib': (PIS, COF)}.get(tipo, (0.0, 0.0))


# ------------------------------------------------------------------ classificação das entradas de terceiros
def classificar_entrada(d, it, ctx):
    """Devolve (natureza, credito, motivo, origem). credito: sim / nao / verificar / devol_m (devolução de monofásico)."""
    c = it.get('CFOP') or ''
    c3 = c[1:]
    ncm = it.get('NCM') or ''
    emit = d['emit']['cnpj']
    mem = ctx['mem']
    if d.get('tpNF') == '0':
        # nota de entrada emitida pela outra parte (ex.: o fornecedor registra a devolução que a empresa fez)
        if c3 in DEV_VENDA:
            m = _memoria_item(mem, emit, it)
            if m and m[1] == 'sim':
                return 'Devolução/estorno de compra (nota da contraparte)', 'estorno', 'Estorna o crédito da compra original (%s)' % m[0], 'memória'
            return 'Devolução/estorno de compra (nota da contraparte)', 'nao', 'Compra devolvida não gerava crédito', 'regra'
        return 'Remessa/retorno (nota da contraparte)', 'nao', 'Anulação/retorno de remessa: sem efeito no PIS/COFINS', 'regra'
    # devolução feita pelo cliente: segue a alíquota da venda
    if c3 in {'201', '202', '410', '411', '553', '660', '661', '662', '208', '209'}:
        if livro(ncm):
            return 'Devolução de venda de livro', 'nao', 'Livro: alíquota zero', 'regra'
        if mono(ncm):
            if ctx['fabricante'] and ncm[:4] in ctx['ncm_produzidos']:
                return 'Devolução de venda - monofásico', 'devol_m', 'Devolução de venda: crédito/dedução à alíquota da venda (2,20% / 10,30%)', 'regra'
            return 'Devolução de revenda monofásica', 'nao', 'Devolução de revenda monofásica (alíquota zero)', 'regra'
        return 'Devolução de venda', 'sim', 'Devolução de venda: crédito/dedução à alíquota da venda', 'regra'
    if c3 in NAO_AQUISICAO:
        return NAO_AQUISICAO[c3], 'nao', 'Remessa/retorno/outra entrada: não é aquisição', 'regra'
    m = _memoria_item(mem, emit, it)
    if m:
        return m[0], m[1], m[2], 'memória'
    if c3 in INDUSTRIALIZACAO:
        if mono(ncm):
            return 'Industrialização por encomenda - monofásico', 'nao', 'Executora com alíquota zero em 33.03-33.07 (Lei 11.051/04, art. 10, §3º): sem crédito', 'regra'
        return 'Industrialização por encomenda (insumo)', 'sim', 'Serviço de industrialização contratado: insumo (art. 3º, II)', 'regra'
    if c3 not in COMPRA:
        return 'Remessa/retorno/outras', 'nao', 'Remessa/retorno/outra entrada: não é aquisição', 'regra'
    if livro(ncm):
        return 'Revenda - livro (alíquota zero)', 'nao', 'Livro: alíquota zero (Lei 10.865/04, art. 28, VI); sem crédito (art. 3º, §2º, II)', 'regra'
    if mono(ncm) and (not ctx['fabricante'] or c3 in REVENDA):
        return 'Revenda - monofásico', 'nao', 'Revenda de monofásico (Lei 10.147/00): vedado (art. 3º, I, b; STJ Tema 1093)', 'regra'
    cst = ((it.get('pis') or {}).get('CST') or '').zfill(2)
    if cst in CST_FORN_SEM and d['emit'].get('crt') == '3':
        return 'A classificar', 'verificar', 'Fornecedor do regime normal emitiu com CST %s (sem PIS/COFINS): confirmar antes de tomar o crédito' % cst, 'regra'
    if not ctx['fabricante'] and ncm[:4] in ctx['ncm_vendidos']:
        return 'Revenda - tributada', 'sim', 'Mercadoria para revenda fora do monofásico (art. 3º, I)', 'regra'
    return 'A classificar', 'verificar', 'Item novo: escolha a natureza (insumo, revenda, uso e consumo...)', 'regra'


def _memoria_item(mem, emit, it):
    return (mem.get('itens') or {}).get(emit + '|' + (it.get('cProd') or '').strip()) or (mem.get('ncm') or {}).get(emit + '|' + (it.get('NCM') or ''))


# ------------------------------------------------------------------ apuração
# Cada item vira uma linha (grupo 'deb' ou 'cred'); débitos e créditos por natureza saem da soma das linhas. Assim a escolha
# feita na tela para um item "A classificar" muda só aquelas linhas e os totais são refeitos sem recalcular o mês.
def apurar(docs, emp, competencia, serv_rel=None, memoria=None, cfg=None):
    rz = emp.get('raiz') or _raiz(emp.get('cnpj'))
    mem = (memoria or {}).get(rz) or {}
    nfe = [d for d in docs if d['tipo'] == 'nfe' and not d['cancelada'] and not d.get('denegada') and d['empresa_envolvida']]
    saidas = [d for d in nfe if d['propria'] and d['raw']['tpNF'] == '1']
    # fabricante = vende produção própria (CFOP x101/x401...) de produto monofásico
    ncm_prod = {it['NCM'][:4] for d in saidas for it in d['raw']['itens'] if (it['CFOP'] or '')[1:] in VENDA_PROD}
    fabricante = any(mono(it['NCM']) for d in saidas for it in d['raw']['itens'] if (it['CFOP'] or '')[1:] in VENDA_PROD)
    ncm_vend = {it['NCM'][:4] for d in saidas for it in d['raw']['itens'] if (it['CFOP'] or '')[1:] in VENDA_PROD | REVENDA}
    ctx = {'mem': mem, 'fabricante': fabricante, 'ncm_produzidos': ncm_prod, 'ncm_vendidos': ncm_vend}
    linhas = []
    frac_venda, frac_compra = {}, {}

    # ---- saídas próprias
    for d in saidas:
        trib = tot = 0.0
        for it in d['raw']['itens']:
            nat, tipo, motivo = classificar_saida(it, fabricante, mem)
            pxml = (it.get('pis') or {}).get('vPIS', 0); cxml = (it.get('cofins') or {}).get('vCOFINS', 0)
            tot += it['vProd']
            if tipo == 'estorno':
                # devolução de compra: estorna o crédito que a compra deu (o fornecedor é o destinatário; a memória acha a compra pelo NCM)
                n2, cr, _, _ = classificar_entrada({'emit': {'cnpj': d['raw']['dest']['cnpj'], 'crt': ''}}, dict(it, CFOP='1102', cProd=''), ctx)
                if cr == 'sim':
                    b = base_credito(it, com_ipi=False)
                    linhas.append(_lin(d, it, 'cred', 'Devolução de compra (estorno)', 'Estorno', b, -b * PIS / 100, -b * COF / 100,
                                       'Estorna o crédito da compra (%s)' % n2))
                else:
                    linhas.append(_lin(d, it, 'deb', 'Devolução de compra', 'nada', 0, 0, 0, 'Compra não gerava crédito (%s)' % n2))
                continue
            p, c = _aliq_venda(tipo)
            b = base_venda(it) if tipo in ('mono', 'trib', 'zero', 'export') else 0.0
            ln = _lin(d, it, 'deb', nat, tipo, b, b * p / 100, b * c / 100, motivo)
            ln['pis_xml'], ln['cofins_xml'] = _r(pxml), _r(cxml)
            linhas.append(ln)
            if tipo in ('mono', 'trib', 'export'):
                trib += it['vProd']
        frac_venda[d['chave']] = (trib / tot) if tot else 0.0

    # ---- entradas: devolução de venda emitida pela empresa e notas de terceiros
    for d in nfe:
        if d['propria'] and d['raw']['tpNF'] == '1':
            continue
        cb = ct = 0.0
        for it in d['raw']['itens']:
            if d['propria']:
                c3 = (it['CFOP'] or '')[1:]
                if c3 in DEV_VENDA:
                    nat_v, tipo_v, _ = classificar_saida(dict(it, CFOP='5' + ('101' if c3 in ('201', '410') else '102')), fabricante, mem)
                    p, c = _aliq_venda(tipo_v)
                    b = base_venda(it) if p else 0.0
                    linhas.append(_lin(d, it, 'cred', 'Devolução de venda (nota própria)', 'Crédito' if p else 'Sem crédito', b, b * p / 100, b * c / 100,
                                       'Devolução de venda: mesma alíquota da venda original (%s)' % nat_v))
                else:
                    linhas.append(_lin(d, it, 'cred', 'Entrada própria (retorno/outros)', 'Sem crédito', 0, 0, 0, 'Remessa/retorno/outra entrada: não é aquisição'))
                continue
            nat, cr, mot, origem = classificar_entrada(d['raw'], it, ctx)
            ct += it['vProd']
            b_ipi, b_sem = base_credito(it, True), base_credito(it, False)
            if cr == 'sim':
                b = b_sem if (fabricante and nat in IPI_RECUPERAVEL) else b_ipi
                vp, vc, sit = b * PIS / 100, b * COF / 100, 'Crédito'
                cb += it['vProd']
            elif cr == 'devol_m':
                b = b_sem; vp, vc, sit = b * PIS_MONO / 100, b * COF_MONO / 100, 'Crédito'
            elif cr == 'estorno':
                b = b_sem; vp, vc, sit = -b * PIS / 100, -b * COF / 100, 'Estorno'
            else:
                b, vp, vc, sit = 0.0, 0.0, 0.0, ('Verificar' if cr == 'verificar' else 'Sem crédito')
            ln = _lin(d, it, 'cred', nat, sit, b, vp, vc, mot)
            ln.update(origem=origem, emit=d['raw']['emit']['cnpj'], cProd=(it.get('cProd') or '').strip(), b_ipi=_r(b_ipi), b_sem=_r(b_sem))
            linhas.append(ln)
        frac_compra[d['chave']] = (cb / ct) if ct else 0.0

    # ---- CT-e (a empresa é a tomadora)
    fv, fc = mem.get('frac_venda') or {}, mem.get('frac_compra') or {}
    for d in docs:
        if d['tipo'] != 'cte' or d['cancelada'] or _raiz(d['raw']['tomador_cnpj']) != rz or d['raw'].get('tpCTe') == '2':
            continue
        x = d['raw']
        venda = _raiz(x['rem']['cnpj']) == rz
        nat = 'Frete na venda' if venda else 'Frete na compra'
        # NF-e de outro mês: fração guardada na memória quando aquele mês foi apurado
        fr = [frac_venda.get(k, fv.get(k)) if venda else frac_compra.get(k, fc.get(k)) for k in x['nfes']]
        fr = [v for v in fr if v is not None]
        liq = x['vTPrest'] - (x.get('icms') or {}).get('vICMS', 0)
        if not fr:
            linhas.append(_lin_cte(d, nat, 'Verificar', 0, 0, 0, 'NF-e transportada não está nas %s enviadas: confirmar' % ('saídas' if venda else 'entradas')))
            continue
        f = sum(fr) / len(fr)
        b = liq * f
        mot = ('Frete na venda (art. 3º, IX): crédito na parte que gerou débito' if venda else 'Frete na compra: segue a mercadoria transportada') + \
              ' (fração %d%%)' % round(f * 100)
        linhas.append(_lin_cte(d, nat, 'Crédito' if b > 0.004 else 'Sem crédito', b, b * PIS / 100, b * COF / 100, mot))

    # ---- serviços prestados (NFS-e XML; sem XML, o relatório do sistema)
    ret = {'pis': 0.0, 'cofins': 0.0}
    prest = [d for d in docs if d['tipo'] == 'nfse' and d['mov'] == 'prestado' and not d['cancelada']]
    for d in prest:
        x = d['raw']; b = x['vServ'] - (x.get('vISS') or 0)
        ret['pis'] += x.get('vPis') or 0; ret['cofins'] += x.get('vCofins') or 0
        linhas.append({'grupo': 'deb', 'lado': 'prestado', 'doc': 'NFS-e %s' % d['numero'], 'data': d['emissao'], 'part': d['part_nome'],
                       'produto': (x.get('xDescServ') or '')[:80], 'valor': _r(x['vServ']), 'natureza': 'Prestação de serviço', 'situacao': 'trib',
                       'base': _r(b), 'pis': _r(b * PIS / 100), 'cofins': _r(b * COF / 100), 'motivo': 'Serviço prestado (NFS-e): 1,65% / 7,6% sobre o valor sem ISS (STF Tema 118)'})
    if not prest and serv_rel and serv_rel.get('saida'):
        for x in serv_rel['saida']:
            b = x['valor'] - (x.get('iss') or 0)
            ret['pis'] += x.get('pis') or 0; ret['cofins'] += x.get('cofins') or 0
            linhas.append({'grupo': 'deb', 'lado': 'prestado', 'doc': 'Serviço %s' % x['numero'], 'data': x['data'], 'part': x['participante'], 'produto': x['descricao'][:80],
                           'valor': x['valor'], 'natureza': 'Prestação de serviço', 'situacao': 'trib', 'base': _r(b), 'pis': _r(b * PIS / 100),
                           'cofins': _r(b * COF / 100), 'motivo': 'Serviço prestado (relatório do sistema): 1,65% / 7,6% sobre o valor sem ISS'})

    # ---- serviços tomados (frete/armazenagem na venda: só a parte tributada das vendas do mês)
    tom = [d for d in docs if d['tipo'] == 'nfse' and d['mov'] == 'tomado' and not d['cancelada']]
    for d in tom:
        x = d['raw']
        m = (mem.get('servicos') or {}).get(x['prestador']['cnpj'])
        nat, sit, mot, cod, origem = (m[0], m[1], m[2], (m[3] if len(m) > 3 else ''), 'memória') if m else \
            ('A classificar', 'Verificar', 'Prestador novo: escolha se o serviço dá crédito', '', 'regra')
        linhas.append({'grupo': 'cred', 'lado': 'tomado', 'doc': 'NFS-e %s' % d['numero'], 'data': d['emissao'], 'part': d['part_nome'],
                       'prestador': x['prestador']['cnpj'], 'produto': (x.get('xDescServ') or '')[:80], 'valor': _r(x['vServ']), 'natureza': nat,
                       'situacao': sit, 'codigo_bc': cod, 'motivo': mot, 'origem': origem, 'competencia': (x.get('competencia') or '')[:7],
                       'base': 0.0, 'pis': 0.0, 'cofins': 0.0})
    pc = {'competencia': competencia, 'fabricante': fabricante, 'linhas': linhas, 'retencoes': {k: _r(v) for k, v in ret.items()},
          'opcoes_entrada': list(OPCOES_ENTRADA), 'opcoes_servico': list(OPCOES_SERVICO),
          '_frac': {'frac_venda': {k: round(v, 6) for k, v in frac_venda.items()}, 'frac_compra': {k: round(v, 6) for k, v in frac_compra.items()}}}
    return totalizar(pc, cfg)


def rsum(ls):
    return 'R$ ' + '{:,.2f}'.format(sum(l.get('valor') or 0 for l in ls)).replace(',', 'X').replace('.', ',').replace('X', '.')


def _servico_base(ln, frac_mes):
    if ln['situacao'] != 'Crédito':
        return 0.0
    return ln['valor'] * (frac_mes if 'venda' in ln['natureza'].lower() else 1.0)


def totalizar(pc, cfg=None):
    """Soma as linhas por natureza e fecha o mês: débito, créditos, saldo credor anterior, retenções, ajustes e a recolher."""
    cfg = cfg or {}
    L = pc['linhas']
    rec = {'tributada': 0.0, 'aliq_zero': 0.0, 'exportacao': 0.0}
    for l in L:
        if l['grupo'] == 'deb':
            t = l['situacao']
            rec['tributada' if t in ('mono', 'trib') else 'aliq_zero' if t == 'zero' else 'exportacao' if t == 'export' else '_'] = \
                rec.get('tributada' if t in ('mono', 'trib') else 'aliq_zero' if t == 'zero' else 'exportacao' if t == 'export' else '_', 0) + l['base']
    rec.pop('_', None)
    frac_mes = rec['tributada'] / (rec['tributada'] + rec['aliq_zero']) if (rec['tributada'] + rec['aliq_zero']) else 0.0
    for l in L:
        if l['lado'] == 'tomado':
            b = _servico_base(l, frac_mes)
            l['base'], l['pis'], l['cofins'] = _r(b), _r(b * PIS / 100), _r(b * COF / 100)
    deb, cred = {}, {}
    for l in L:
        if l['grupo'] == 'deb':
            g = deb.setdefault(l['natureza'], {'natureza': l['natureza'], 'tipo': l['situacao'], 'itens': 0, 'valor': 0.0, 'base': 0.0, 'pis': 0.0,
                                                'cofins': 0.0, 'pis_xml': 0.0, 'cofins_xml': 0.0, 'motivo': l.get('motivo') or ''})
        else:
            g = cred.setdefault((l['natureza'], l['situacao']), {'natureza': l['natureza'], 'situacao': l['situacao'], 'itens': 0, 'valor': 0.0,
                                                                 'base': 0.0, 'pis': 0.0, 'cofins': 0.0, 'motivo': l.get('motivo') or ''})
        g['itens'] += 1
        for k in ('valor', 'base', 'pis', 'cofins', 'pis_xml', 'cofins_xml'):
            if k in g:
                g[k] += l.get(k) or 0
    aj = [a for a in (cfg.get('ajustes') or []) if a.get('pis') or a.get('cofins')]
    aj_deb = {'pis': sum(float(a.get('pis') or 0) for a in aj if a.get('tipo') == 'debito'), 'cofins': sum(float(a.get('cofins') or 0) for a in aj if a.get('tipo') == 'debito')}
    aj_cred = {'pis': sum(float(a.get('pis') or 0) for a in aj if a.get('tipo') != 'debito'), 'cofins': sum(float(a.get('cofins') or 0) for a in aj if a.get('tipo') != 'debito')}
    t = {}
    for imp in ('pis', 'cofins'):
        d_ = sum(g[imp] for g in deb.values()) + aj_deb[imp]
        c_ = sum(g[imp] for g in cred.values()) + aj_cred[imp]
        ant = float(cfg.get('saldo_%s_ant' % imp) or 0)
        ret = float((pc.get('retencoes') or {}).get(imp) or 0)
        liq = d_ - c_ - ant
        t.update({imp + '_debito': _r(d_), imp + '_credito': _r(c_), imp + '_saldo_anterior': _r(ant), imp + '_retido': _r(ret),
                  imp + '_ajuste_debito': _r(aj_deb[imp]), imp + '_ajuste_credito': _r(aj_cred[imp]),
                  imp + '_recolher': _r(max(0.0, liq - ret)), imp + '_saldo_credor': _r(max(0.0, -liq))})
    t['total_recolher'] = _r(t['pis_recolher'] + t['cofins_recolher'])
    alertas = []
    fora = [l for l in L if l.get('competencia') and pc['competencia'] and l['competencia'] != pc['competencia']]
    if fora:
        alertas.append('%d NFS-e tomada(s) com competência diferente de %s (%s): o crédito entrou neste mês; confirme se é extemporâneo.'
                       % (len(fora), pc['competencia'], ', '.join(sorted({l['doc'] for l in fora})[:6])))
    pend = [l for l in L if l['situacao'] == 'Verificar' and l['lado'] in ('entrada', 'tomado')]
    cte_sem = [l for l in L if l['situacao'] == 'Verificar' and l['lado'] == 'frete']
    if cte_sem:
        alertas.append('%d CT-e (%s) sem a NF-e transportada nos arquivos do mês nem em meses já apurados: ficaram sem crédito. Envie a NF-e para entrar.'
                       % (len(cte_sem), rsum(cte_sem)))
    pc.update(debitos=sorted([{k: (_r(v) if isinstance(v, float) else v) for k, v in g.items()} for g in deb.values()], key=lambda g: -g['valor']),
              creditos=sorted([{k: (_r(v) if isinstance(v, float) else v) for k, v in g.items()} for g in cred.values()],
                              key=lambda g: (g['situacao'] not in ('Crédito', 'Estorno'), -abs(g['pis']), -g['valor'])),
              receita={k: _r(v) for k, v in rec.items()}, frac_tributada=round(frac_mes, 4), totais=t, alertas=alertas,
              a_classificar=len(pend), cfg={'saldo_pis_ant': cfg.get('saldo_pis_ant') or 0, 'saldo_cofins_ant': cfg.get('saldo_cofins_ant') or 0,
                                            'ajustes': aj, 'saldo_origem': cfg.get('saldo_origem') or ''})
    return pc


def aplicar_decisao(pc, idx, natureza, fabricante=None):
    """Aplica a escolha da tela à linha idx e a todas do mesmo fornecedor+produto (ou do mesmo prestador). Devolve a linha modelo."""
    ref = pc['linhas'][idx]
    if ref['lado'] == 'tomado':
        sit, mot, cod = OPCOES_SERVICO[natureza]
        alvo = [l for l in pc['linhas'] if l['lado'] == 'tomado' and l.get('prestador') == ref.get('prestador')]
        for l in alvo:
            l.update(natureza=natureza, situacao=sit, motivo=mot + ' (decisão na tela)', codigo_bc=cod, origem='decisão')
    else:
        cr, mot, ipi_rec = OPCOES_ENTRADA[natureza]
        fab = pc.get('fabricante') if fabricante is None else fabricante
        alvo = [l for l in pc['linhas'] if l['lado'] == 'entrada' and l.get('emit') == ref.get('emit') and l.get('cProd') == ref.get('cProd')]
        for l in alvo:
            b = (l['b_sem'] if (fab and ipi_rec) else l['b_ipi']) if cr == 'sim' else 0.0
            l.update(natureza=natureza, situacao='Crédito' if cr == 'sim' else 'Sem crédito', motivo=mot + ' (decisão na tela)', origem='decisão',
                     base=_r(b), pis=_r(b * PIS / 100), cofins=_r(b * COF / 100))
    return ref


def _lin(d, it, grupo, nat, sit, b, p, c, mot):
    return {'grupo': grupo, 'lado': 'saida' if grupo == 'deb' or (d['propria'] and d['raw']['tpNF'] == '1') else 'entrada', 'chave': d['chave'],
            'doc': 'NF-e %s' % d['numero'], 'data': d['emissao'], 'part': d['part_nome'], 'item': it.get('nItem'),
            'produto': (it.get('xProd') or '')[:80], 'NCM': it.get('NCM'), 'CFOP': it.get('CFOP'), 'valor': _r(it['vProd']),
            'natureza': nat, 'situacao': sit, 'base': _r(b), 'pis': _r(p), 'cofins': _r(c), 'motivo': mot}


def _lin_cte(d, nat, sit, b, p, c, mot):
    return {'grupo': 'cred', 'lado': 'frete', 'chave': d['chave'], 'doc': 'CT-e %s' % d['numero'], 'data': d['emissao'], 'part': d['part_nome'],
            'valor': _r(d['raw']['vTPrest']), 'natureza': nat, 'situacao': sit, 'base': _r(b), 'pis': _r(p), 'cofins': _r(c), 'motivo': mot}


def guardar_fracoes(arq, raiz, fracs):
    """Guarda a fração tributada de cada NF-e apurada: o CT-e do mês seguinte que levar essa nota usa a fração dela."""
    mem = ler_memoria(arq)
    m = mem.setdefault(raiz, {})
    for k, v in (fracs or {}).items():
        m.setdefault(k, {}).update(v)
    os.makedirs(os.path.dirname(arq), exist_ok=True)
    json.dump(mem, open(arq, 'w', encoding='utf-8'), ensure_ascii=False)


def decidir(arq, raiz, linha, natureza):
    """Grava a escolha da tela na memória (vale para os próximos meses)."""
    if linha.get('lado') == 'tomado':
        sit, mot, cod = OPCOES_SERVICO[natureza]
        gravar_decisao(arq, raiz, 'servicos', linha['prestador'], [natureza, sit, mot + ' (decisão na tela)', cod])
    else:
        cr, mot, _ = OPCOES_ENTRADA[natureza]
        gravar_decisao(arq, raiz, 'itens', '%s|%s' % (linha['emit'], linha.get('cProd') or ''), [natureza, cr, mot + ' (decisão na tela)'])
