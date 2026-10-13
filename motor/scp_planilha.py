# -*- coding: utf-8 -*-
"""Modelo da planilha de fechamento do escritório (abas Pis.Cofins, Irpj.Csll, Economia e Icms), preenchido com o
cálculo da Estratégia SCP. O mesmo modelo vira a tela (HTML) e o Excel (openpyxl, com fórmulas nas mesmas células).

Cada célula: {'v': valor calculado, 'f': fórmula do Excel (opcional), 'st': estilo, 'fmt': formato}."""
import io
import re
from .regras_base import desc_cfop

MESES = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
# cores do tema da planilha (tx2 0E2841 com tons): título, seção, linha de crédito, cinza
COR = {'tit': ('0A1E31', 'FFFFFF'), 'sub': ('CFD4D9', '0E2841'), 'clr': ('C3C9CF', '000000'), 'pag': ('CFD4D9', '000000'), 'cinza': ('A6A6A6', '000000'),
       'totg': ('0A1E31', 'FFFFFF'), 'cab': ('0A1E31', 'FFFFFF')}


class Aba:
    def __init__(self, nome, larg):
        self.nome, self.larg, self.cel, self.mesc = nome, larg, {}, []

    def c(self, ref, v=None, f=None, st='lin', fmt=None, b=False, ed=None, dt=None):
        if fmt is None:
            fmt = 'num' if isinstance(v, (int, float)) and not isinstance(v, bool) else 'txt'
        self.cel[ref] = {'v': v, 'f': f, 'st': st, 'fmt': fmt, 'b': b or st in ('tit', 'sub', 'totg', 'pag', 'cab', 'tot', 'cinza')}
        if ed:
            self.cel[ref]['ed'] = ed  # campo do cfg que o ADM edita nesta célula

    def m(self, rng):
        self.mesc.append(rng)


def _r(v):
    return round(v or 0, 2)


def montar(res, k, apu=None):
    c, d, mp, mc, sp = k['cfg'], k['resumo']['distribuicao'], k['matriz']['pis'], k['matriz']['cofins'], k['scp']['pc']
    comp = res.get('competencia') or ''
    ano, mes = (comp[:4], int(comp[5:7])) if len(comp) >= 7 else ('', 1)
    nome_mes = MESES[mes - 1]
    emp = (res.get('empresa') or {}).get('nome') or ''
    curto = c.get('matriz_nome') or (' '.join(emp.split()[1:3]) if len(emp.split()) > 2 else emp)
    scp_nome = c.get('scp_nome') or ('SCP ' + curto)
    ret_m, ret_s = k['retencoes']['matriz'], k['retencoes']['scp']
    linhas_deb = k['linhas_debito']
    prod_vc = _r(c['prod_vc_manual']) if c.get('prod_vc_manual') is not None else _r(sum(x['vc'] for x in linhas_deb if x['sinal'] > 0))
    venda_base = _r(sum(x['base'] for x in linhas_deb if x['sinal'] > 0))
    devol = _r(-sum(x['base'] for x in linhas_deb if x['sinal'] < 0))
    imp = k.get('importacao') or {}
    abas = []

    # ============================================================ Pis.Cofins
    a = Aba('Pis.Cofins', {'A': 3, 'B': 30, 'C': 15, 'D': 7, 'E': 30, 'F': 15, 'J': 22, 'K': 18, 'L': 18, 'M': 18})
    t = 'PIS/COFINS - %s %s' % (nome_mes, ano)
    a.c('B2', t, st='tit'); a.m('B2:C2'); a.c('E2', t, st='tit'); a.m('E2:F2')
    a.c('B3', 'APURAÇÃO LUCRO PRESUMIDO - %s' % scp_nome.upper(), st='sub'); a.m('B3:C3')
    a.c('E3', 'APURAÇÃO LUCRO REAL %s' % curto.upper(), st='sub'); a.m('E3:F3')
    base_cred_prod = sum(x['base'] for x in k['linhas_credito'])
    base_tom = next((l['valor'] for l in k['memoria'].get('creditos', []) if l['rot'].startswith('Serviços tomados')), 0)
    alug = next((l['valor'] for l in k['memoria'].get('creditos', []) if l['rot'].startswith('Aluguéis')), 0)
    imp_base = imp.get('base', 0) if imp.get('modo') == 'base' else 0
    if c.get('cred_prod_manual') is not None:
        base_cred_prod = c['cred_prod_manual']
    for r, rot, vs, vm, edc in ((4, 'Entrada Produtos', 0, base_cred_prod, 'cred_prod_manual'), (5, 'Entrada Serviços Tomados', 0, base_tom + alug, 'tomados_base_manual'),
                                (6, 'Importação ', None, imp_base, 'imp_base_manual')):
        a.c('B%d' % r, rot); a.c('C%d' % r, vs if vs is not None else None, fmt='num')
        a.c('E%d' % r, rot); a.c('F%d' % r, _r(vm), fmt='num', ed=edc)
    a.c('B7', 'Total créditos', st='tot'); a.c('C7', 0, '=SUM(C4:C5)', st='tot')
    a.c('E7', 'Total créditos', st='tot'); a.c('F7', _r(base_cred_prod + base_tom + alug + imp_base), '=SUM(F4:F6)', st='tot')
    a.c('B9', 'Saldo Credor PIS 0,65%'); a.c('C9', 0)
    a.c('E9', 'Saldo Credor PIS 1,65%'); a.c('F9', _r(mp['saldo_ant']), ed='saldo_pis_ant')
    a.c('B10', 'Saldo Credor COFINS 3%'); a.c('C10', 0)
    a.c('E10', 'Saldo Credor COFINS 7,60%'); a.c('F10', _r(mc['saldo_ant']), ed='saldo_cof_ant')
    extra_pis = _r(mp['cred_periodo'] - (base_cred_prod + base_tom + alug + imp_base) * 0.0165)
    extra_cof = _r(mc['cred_periodo'] - (base_cred_prod + base_tom + alug + imp_base) * 0.076)
    a.c('B11', 'PIS 0,65%', st='clr'); a.c('C11', 0, '=C7*1.65%', st='clr')
    a.c('E11', 'PIS 0 1,65%', st='clr'); a.c('F11', _r(mp['cred_periodo']), '=F7*1.65%' + ('+%s' % extra_pis if abs(extra_pis) > 0.004 else ''), st='clr')
    a.c('B12', 'COFINS 3%', st='clr'); a.c('C12', 0, '=C7*7.6%', st='clr')
    a.c('E12', 'COFINS 7,60%', st='clr'); a.c('F12', _r(mc['cred_periodo']), '=F7*7.6%' + ('+%s' % extra_cof if abs(extra_cof) > 0.004 else ''), st='clr')
    a.c('B14', 'Prestação de serviços'); a.c('C14', d['serv_scp'], ed='@serv_scp')
    a.c('E14', 'Prestação de serviços'); a.c('F14', d['serv_matriz'], ed='@serv_matriz')
    a.c('B15', 'Locação '); a.c('C15', d['loc_scp'], ed='@loc_scp')
    a.c('E15', 'Locação '); a.c('F15', d['loc_matriz'], ed='@loc_matriz')
    a.c('B16', 'Saídas Produtos '); a.c('C16', None, fmt='num')
    if c.get('base_prod_manual') is not None:
        venda_base = c['base_prod_manual']
    if c.get('devol_manual') is not None:
        devol = c['devol_manual']
    a.c('E16', 'Saídas Produtos '); a.c('F16', venda_base, ed='base_prod_manual')
    a.c('B17', 'Outras receitas'); a.c('C17', c.get('scp_outras') or 0, ed='scp_outras')
    a.c('E17', '(−) Devoluções de venda'); a.c('F17', -devol if devol else 0, ed='@devol')
    a.c('B18', 'Total Débitos', st='pag'); a.c('C18', sp['base'], '=SUM(C14:C17)', st='pag')
    a.c('E18', 'Total Débitos', st='pag'); a.c('F18', mp['base_deb'], '=SUM(F14:F17)', st='pag')
    a.c('B20', 'PIS 0,65%'); a.c('C20', sp['pis'], '=C18*0.65%')
    a.c('E20', 'PIS 1,65%'); a.c('F20', mp['debito'], '=F18*1.65%')
    a.c('B21', 'COFINS 3%'); a.c('C21', sp['cofins'], '=C18*3%')
    a.c('E21', 'COFINS 7,60%'); a.c('F21', mc['debito'], '=F18*7.6%')
    a.c('B23', 'PIS A PAGAR', st='pag'); a.c('C23', sp['pis'], '=C20', st='pag'); a.c('D23', 8109, st='cod', fmt='txt')
    a.c('E23', 'PIS A PAGAR', st='cinza'); a.c('F23', _r(mp['debito'] - mp['disponivel']), '=F20-(F11+F9)', st='cinza')
    a.c('B24', 'COFINS A PAGAR ', st='pag'); a.c('C24', sp['cofins'], '=C21', st='pag'); a.c('D24', 2172, st='cod', fmt='txt')
    a.c('E24', 'COFINS A PAGAR ', st='cinza'); a.c('F24', _r(mc['debito'] - mc['disponivel']), '=F21-(F12+F10)', st='cinza')
    a.c('B25', 'TOTAL PIS/COFINS', st='totg'); a.c('C25', _r(sp['pis'] + sp['cofins']), '=SUM(C23:C24)', st='totg')
    a.c('E25', 'TOTAL PIS/COFINS', st='totg'); a.c('F25', _r(mp['debito'] - mp['disponivel'] + mc['debito'] - mc['disponivel']), '=SUM(F23:F24)', st='totg')
    # complemento: retenções, valor a recolher e saldo credor a transportar
    a.c('B27', 'RETENÇÕES E SALDOS', st='sub'); a.m('B27:C27'); a.c('E27', 'RETENÇÕES E SALDOS', st='sub'); a.m('E27:F27')
    a.c('B28', 'PIS retido nas notas'); a.c('C28', _r(ret_s['pis']), ed='ret_s_pis')
    a.c('E28', 'PIS retido nas notas'); a.c('F28', _r(ret_m['pis']), ed='ret_m_pis')
    a.c('B29', 'COFINS retida nas notas'); a.c('C29', _r(ret_s['cofins']), ed='ret_s_cofins')
    a.c('E29', 'COFINS retida nas notas'); a.c('F29', _r(ret_m['cofins']), ed='ret_m_cofins')
    a.c('B30', 'PIS a recolher', st='pag'); a.c('C30', sp['pis_rec'], '=MAX(C23-C28,0)', st='pag')
    a.c('E30', 'PIS a recolher', st='pag'); a.c('F30', mp['a_recolher'], '=MAX(F23-F28,0)', st='pag')
    a.c('B31', 'COFINS a recolher', st='pag'); a.c('C31', sp['cofins_rec'], '=MAX(C24-C29,0)', st='pag')
    a.c('E31', 'COFINS a recolher', st='pag'); a.c('F31', mc['a_recolher'], '=MAX(F24-F29,0)', st='pag')
    a.c('E32', 'Saldo credor PIS a transportar'); a.c('F32', mp['saldo_credor_final'], '=MAX(-F23,0)')
    a.c('E33', 'Saldo credor COFINS a transportar'); a.c('F33', mc['saldo_credor_final'], '=MAX(-F24,0)')
    # faturamento e proporção
    a.c('J2', 'Locação'); a.c('K2', d['loc_total'], fmt='brl', ed='loc_manual')
    a.c('J3', 'Prestação '); a.c('K3', d['serv_total'], fmt='brl', ed='serv_manual')
    a.c('J4', 'Produtos'); a.c('K4', prod_vc, fmt='brl', ed='prod_vc_manual')
    a.c('J5', 'Faturamento', b=True); a.c('K5', _r(d['loc_total'] + d['serv_total'] + prod_vc), '=SUM(K2:K4)', fmt='brl', b=True)
    a.c('J7', 'DISTRIBUIÇÃO DO FATURAMENTO', st='tit'); a.m('J7:M7')
    a.c('J8', 'Receita', st='sub'); a.c('K8', 'Lucro Real (Matriz)', st='sub'); a.c('L8', 'Presumido (SCP)', st='sub'); a.c('M8', '% na SCP', st='sub')
    a.c('J9', 'Prestação '); a.c('K9', d['serv_matriz'], '=F14', fmt='brl'); a.c('L9', d['serv_scp'], '=C14', fmt='brl'); a.c('M9', (d['serv_scp'] / d['serv_total']) if d['serv_total'] else 0, '=IF(K3=0,0,L9/K3)', fmt='pct')
    a.c('J10', 'Locação'); a.c('K10', d['loc_matriz'], '=F15', fmt='brl'); a.c('L10', d['loc_scp'], '=C15', fmt='brl'); a.c('M10', (d['loc_scp'] / d['loc_total']) if d['loc_total'] else 0, '=IF(K2=0,0,L10/K2)', fmt='pct')
    a.c('J11', 'Produtos'); a.c('K11', prod_vc, '=K4', fmt='brl'); a.c('L11', 0, fmt='brl'); a.c('M11', 0, fmt='pct')
    tot_m, tot_s = _r(d['serv_matriz'] + d['loc_matriz'] + prod_vc), _r(d['serv_scp'] + d['loc_scp'])
    a.c('J12', 'Total', st='pag'); a.c('K12', tot_m, '=SUM(K9:K11)', st='pag', fmt='brl'); a.c('L12', tot_s, '=SUM(L9:L11)', st='pag', fmt='brl')
    a.c('M12', tot_s / (tot_m + tot_s) if (tot_m + tot_s) else 0, '=IF(K12+L12=0,0,L12/(K12+L12))', st='pag', fmt='pct')
    a.c('J13', '% do faturamento'); a.c('K13', tot_m / (tot_m + tot_s) if (tot_m + tot_s) else 0, '=IF(K12+L12=0,0,K12/(K12+L12))', fmt='pct')
    a.c('L13', tot_s / (tot_m + tot_s) if (tot_m + tot_s) else 0, '=IF(K12+L12=0,0,L12/(K12+L12))', fmt='pct')
    a.c('J14', 'Matriz fica com serviço/locação até igualar os créditos (entrada = saída); o excedente vai para a SCP. Produto fica na Matriz.' if d['modo'] == 'auto'
        else 'Distribuição informada manualmente. Produto fica na Matriz.', st='nota')
    abas.append(a)

    # ============================================================ Irpj.Csll
    # dois quadros lado a lado (Lucro Presumido da SCP × Lucro Real da Matriz), cada um com o mês e o fechamento do trimestre;
    # o faturamento de cada regime e o % do total mostram a divisão
    a = Aba('Irpj.Csll', {'A': 34, 'B': 17, 'C': 17, 'D': 7, 'E': 3, 'F': 34, 'G': 17, 'H': 17})
    tri = (mes - 1) // 3 + 1
    ks = k['scp']
    ini = (tri - 1) * 3 + 1
    outros = [x for x in (c.get('tri_outros_meses') or (k.get('auto') or {}).get('tri_outros_meses') or []) if x.get('competencia', '') < comp]
    irrf_tri = next((l['valor'] for l in k['memoria']['scp_ir'] if l['rot'].startswith('(−) IRRF')), 0)
    csll_tri = next((l['valor'] for l in k['memoria']['scp_ir'] if l['rot'].startswith('(−) CSLL')), 0)
    pago = next((l['valor'] for l in k['memoria']['scp_ir'] if l['rot'].startswith('(−) IRPJ/CSLL já')), 0)
    lucro = c.get('lucro_matriz') or 0
    fat_m_tri = _r(tot_m + sum((x.get('receita_matriz') or 0) for x in outros))
    lucro_tri = c['lucro_matriz_tri'] if c.get('lucro_matriz_tri') is not None else _r(lucro + sum((x.get('lucro') or 0) for x in outros))
    irrf_m_tri = _r(ret_m.get('irrf', 0) + sum((x.get('irrf_m') or 0) for x in outros))
    csll_m_tri = _r(ret_m.get('csll', 0) + sum((x.get('csll_m') or 0) for x in outros))
    t = 'IRPJ E CSLL - %d° TRIMESTRE %s' % (tri, ano)
    a.c('A2', t, st='tit'); a.m('A2:H2')
    a.c('A3', 'APURAÇÃO LUCRO PRESUMIDO - %s' % scp_nome.upper(), st='sub'); a.m('A3:D3')
    a.c('F3', 'APURAÇÃO LUCRO REAL - %s' % curto.upper(), st='sub'); a.m('F3:H3')
    rot_tri = '%d° TRIMESTRE' % tri
    for col, rot in (('A', ''), ('B', nome_mes.upper()), ('C', rot_tri), ('F', ''), ('G', nome_mes.upper()), ('H', rot_tri)):
        a.c('%s4' % col, rot, st='cab')
    base = sp['base']; pres = _r(base * 0.32)
    # Lucro Presumido (SCP)
    a.c('A5', 'FATURAMENTO', b=True); a.c('B5', base, "='Pis.Cofins'!C18", b=True); a.c('C5', ks['base_tri'], b=True)
    a.c('A6', '% do faturamento total'); a.c('B6', base / (base + tot_m) if (base + tot_m) else 0, '=IF(B5+G5=0,0,B5/(B5+G5))', fmt='pct')
    a.c('C6', ks['base_tri'] / (ks['base_tri'] + fat_m_tri) if (ks['base_tri'] + fat_m_tri) else 0, '=IF(C5+H5=0,0,C5/(C5+H5))', fmt='pct')
    a.c('A7', 'PRESUNÇÃO 32%', b=True); a.c('B7', pres, '=B5*32%', b=True); a.c('C7', ks['pres_tri'], '=C5*32%', b=True)
    a.c('A8', 'IRPJ 15%'); a.c('B8', _r(pres * 0.15), '=B7*15%'); a.c('C8', _r(ks['pres_tri'] * 0.15), '=C7*15%')
    a.c('A9', 'ADICIONAL IRPJ 10%'); a.c('B9', _r(max(pres - 20000, 0) * 0.10), '=MAX(B7-20000,0)*10%'); a.c('C9', _r(max(ks['pres_tri'] - 60000, 0) * 0.10), '=MAX(C7-60000,0)*10%')
    a.c('A10', 'IRPJ DEVIDO', st='pag'); a.c('B10', _r(pres * 0.15 + max(pres - 20000, 0) * 0.10), '=B8+B9', st='pag'); a.c('C10', ks['irpj'], '=C8+C9', st='pag')
    a.c('A11', '(−) IRRF retido'); a.c('B11', _r(ret_s['irrf']), ed='ret_s_irrf'); a.c('C11', _r(irrf_tri))
    a.c('A12', 'IRPJ A RECOLHER', st='pag'); a.c('B12', _r(max(a.cel['B10']['v'] - ret_s['irrf'], 0)), '=MAX(B10-B11,0)', st='pag'); a.c('C12', _r(max(ks['irpj'] - irrf_tri, 0)), '=MAX(C10-C11,0)', st='pag')
    a.c('D12', 2089, st='cod', fmt='txt')
    a.c('A13', 'CSLL 9%'); a.c('B13', _r(pres * 0.09), '=B7*9%'); a.c('C13', ks['csll'], '=C7*9%')
    a.c('A14', '(−) CSLL retida'); a.c('B14', _r(ret_s['csll']), ed='@ret_s_csll'); a.c('C14', _r(csll_tri))
    a.c('A15', 'CSLL A RECOLHER', st='pag'); a.c('B15', _r(max(pres * 0.09 - ret_s['csll'], 0)), '=MAX(B13-B14,0)', st='pag'); a.c('C15', _r(max(ks['csll'] - csll_tri, 0)), '=MAX(C13-C14,0)', st='pag')
    a.c('D15', 2372, st='cod', fmt='txt')
    a.c('A16', '(−) Já recolhido no trimestre'); a.c('B16', None, fmt='num'); a.c('C16', _r(pago))
    a.c('A17', 'TOTAL IRPJ/CSLL', st='totg'); a.c('B17', _r(a.cel['B12']['v'] + a.cel['B15']['v']), '=B12+B15', st='totg')
    a.c('C17', _r(max(a.cel['C12']['v'] + a.cel['C15']['v'] - pago, 0)), '=MAX(C12+C15-C16,0)', st='totg')
    a.c('A18', 'QUOTA IRPJ (1/3)'); a.c('B18', None, fmt='num'); a.c('C18', _r(a.cel['C12']['v'] / 3), '=C12/3')
    a.c('A19', 'QUOTA CSLL (1/3)'); a.c('B19', None, fmt='num'); a.c('C19', _r(a.cel['C15']['v'] / 3), '=C15/3')
    # Lucro Real (Matriz)
    a.c('F5', 'FATURAMENTO', b=True); a.c('G5', tot_m, "='Pis.Cofins'!K12", b=True); a.c('H5', fat_m_tri, b=True)
    a.c('F6', '% do faturamento total'); a.c('G6', tot_m / (base + tot_m) if (base + tot_m) else 0, '=IF(B5+G5=0,0,G5/(B5+G5))', fmt='pct')
    a.c('H6', fat_m_tri / (ks['base_tri'] + fat_m_tri) if (ks['base_tri'] + fat_m_tri) else 0, '=IF(C5+H5=0,0,H5/(C5+H5))', fmt='pct')
    a.c('F7', 'RESULTADO DO PERÍODO (LUCRO REAL)', b=True); a.c('G7', lucro, b=True, ed='lucro_matriz'); a.c('H7', lucro_tri, b=True, ed='lucro_matriz_tri')
    a.c('F8', 'IRPJ 15%'); a.c('G8', _r(max(lucro, 0) * 0.15), '=MAX(G7,0)*15%'); a.c('H8', _r(max(lucro_tri, 0) * 0.15), '=MAX(H7,0)*15%')
    a.c('F9', 'ADICIONAL IRPJ 10%'); a.c('G9', _r(max(lucro - 20000, 0) * 0.10), '=MAX(G7-20000,0)*10%'); a.c('H9', _r(max(lucro_tri - 60000, 0) * 0.10), '=MAX(H7-60000,0)*10%')
    a.c('F10', 'IRPJ DEVIDO', st='pag'); a.c('G10', _r(a.cel['G8']['v'] + a.cel['G9']['v']), '=G8+G9', st='pag'); a.c('H10', _r(a.cel['H8']['v'] + a.cel['H9']['v']), '=H8+H9', st='pag')
    a.c('F11', '(−) IRRF retido'); a.c('G11', _r(ret_m.get('irrf', 0)), ed='ret_m_irrf'); a.c('H11', irrf_m_tri)
    a.c('F12', 'IRPJ A RECOLHER', st='pag'); a.c('G12', _r(max(a.cel['G10']['v'] - ret_m.get('irrf', 0), 0)), '=MAX(G10-G11,0)', st='pag'); a.c('H12', _r(max(a.cel['H10']['v'] - irrf_m_tri, 0)), '=MAX(H10-H11,0)', st='pag')
    a.c('F13', 'CSLL 9%'); a.c('G13', _r(max(lucro, 0) * 0.09), '=MAX(G7,0)*9%'); a.c('H13', _r(max(lucro_tri, 0) * 0.09), '=MAX(H7,0)*9%')
    a.c('F14', '(−) CSLL retida'); a.c('G14', _r(ret_m.get('csll', 0)), ed='ret_m_csll'); a.c('H14', csll_m_tri)
    a.c('F15', 'CSLL A RECOLHER', st='pag'); a.c('G15', _r(max(a.cel['G13']['v'] - ret_m.get('csll', 0), 0)), '=MAX(G13-G14,0)', st='pag'); a.c('H15', _r(max(a.cel['H13']['v'] - csll_m_tri, 0)), '=MAX(H13-H14,0)', st='pag')
    a.c('F16', ''); a.c('G16', None, fmt='num'); a.c('H16', None, fmt='num')
    a.c('F17', 'TOTAL IRPJ/CSLL', st='totg'); a.c('G17', _r(a.cel['G12']['v'] + a.cel['G15']['v']), '=G12+G15', st='totg'); a.c('H17', _r(a.cel['H12']['v'] + a.cel['H15']['v']), '=H12+H15', st='totg')
    abas.append(a)

    # ============================================================ Economia
    # quatro quadros nas mesmas linhas: faturamento (3), impostos (5 a 12), carga (14) e % (15); fechamento embaixo
    a = Aba('Economia', {'A': 26, 'B': 17, 'C': 3, 'D': 26, 'E': 17, 'F': 3, 'G': 26, 'H': 17, 'I': 3, 'J': 24, 'K': 17, 'L': 17})
    ant = {kk: c.get('ant_' + kk) for kk in ('faturamento', 'pis', 'cofins', 'irpj', 'csll', 'ipi', 'iss', 'icms')}
    linhas_imp = [('PIS', 'pis'), ('COFINS', 'cofins'), ('IRPJ', 'irpj'), ('IRPJ ADICIONAL', 'adic'), ('CSLL', 'csll'), ('ISS', 'iss'), ('IPI', 'ipi'), ('ICMS', 'icms')]
    # regime anterior
    a.c('A2', 'REGIME ANTERIOR - LUCRO PRESUMIDO', st='tit'); a.m('A2:B2')
    a.c('A3', 'FATURAMENTO', b=True); a.c('B3', ant['faturamento'] or 0, fmt='brl', b=True, ed='ant_faturamento')
    for i, (rot, kk) in enumerate(linhas_imp):
        a.c('A%d' % (5 + i), rot)
        a.c('B%d' % (5 + i), 0 if kk == 'adic' else (ant[kk] or 0), fmt='brl', ed=None if kk == 'adic' else 'ant_' + kk)
    carga_ant = sum(ant[kk] or 0 for kk in ('pis', 'cofins', 'irpj', 'csll', 'ipi', 'iss', 'icms'))
    pct_ant = (carga_ant / ant['faturamento']) if ant['faturamento'] else (c.get('ant_carga_pct') or 21.69) / 100
    a.c('A14', 'CARGA TRIBUTÁRIA', st='pag'); a.c('B14', _r(carga_ant), '=SUM(B5:B12)', st='pag', fmt='brl')
    a.c('A15', '%', st='pag'); a.c('B15', pct_ant, '=IF(B3=0,%s,B14/B3)' % ((c.get('ant_carga_pct') or 21.69) / 100), st='pag', fmt='pct')
    # SCP (Lucro Presumido)
    m3 = sp['base']; n8 = _r(m3 * 0.32)
    pis_s, cof_s = sp['pis_rec'], sp['cofins_rec']
    irpj_s = _r(n8 * 0.15); adic_s = _r(max(n8 - 20000, 0) * 0.10); csll_s = _r(n8 * 0.09)
    a.c('D2', 'LUCRO PRESUMIDO - SCP', st='tit'); a.m('D2:E2')
    a.c('D3', 'FATURAMENTO', b=True); a.c('E3', m3, "='Pis.Cofins'!C18", fmt='brl', b=True)
    vs_s = {'pis': (pis_s, "='Pis.Cofins'!C30"), 'cofins': (cof_s, "='Pis.Cofins'!C31"), 'irpj': (irpj_s, '=E3*32%*15%'), 'adic': (adic_s, '=MAX(E3*32%-20000,0)*10%'),
            'csll': (csll_s, '=E3*32%*9%'), 'iss': (0, None), 'ipi': (0, None), 'icms': (0, None)}
    for i, (rot, kk) in enumerate(linhas_imp):
        a.c('D%d' % (5 + i), rot); a.c('E%d' % (5 + i), vs_s[kk][0], vs_s[kk][1], fmt='brl')
    carga_s = _r(pis_s + cof_s + irpj_s + adic_s + csll_s)
    # Matriz (Lucro Real)
    apc = (apu or {}).get('calc') or {}
    icms_rec = _r(max(apc.get('saldo', 0) or 0, 0)) if apc else _r(max((k.get('icms_saldo') or 0), 0))
    ipi_rec = _r(max(apc.get('ipiSaldo', 0) or 0, 0)) if apc else 0
    iss_m = k.get('iss', 0)
    if c.get('icms_manual') is not None:
        icms_rec = c['icms_manual']
    if c.get('ipi_manual') is not None:
        ipi_rec = c['ipi_manual']
    if c.get('iss_manual') is not None:
        iss_m = c['iss_manual']
    pis_m, cof_m = mp['a_recolher'], mc['a_recolher']
    q3 = tot_m
    a.c('G2', 'LUCRO REAL - MATRIZ', st='tit'); a.m('G2:H2')
    a.c('G3', 'FATURAMENTO', b=True); a.c('H3', q3, "='Pis.Cofins'!K12", fmt='brl', b=True)
    vs_m = {'pis': (pis_m, "='Pis.Cofins'!F30", None), 'cofins': (cof_m, "='Pis.Cofins'!F31", None), 'irpj': (0, "='Irpj.Csll'!G12", None), 'adic': (0, None, None),
            'csll': (0, "='Irpj.Csll'!G15", None), 'iss': (iss_m, None, 'iss_manual'), 'ipi': (ipi_rec, None, 'ipi_manual'), 'icms': (icms_rec, None, 'icms_manual')}
    vs_m['irpj'] = (_r(max(lucro, 0) * 0.15 + max(lucro - 20000, 0) * 0.10 - ret_m.get('irrf', 0)) if lucro > 0 else 0, "='Irpj.Csll'!G12", None)
    vs_m['csll'] = (_r(max(max(lucro, 0) * 0.09 - ret_m.get('csll', 0), 0)), "='Irpj.Csll'!G15", None)
    for i, (rot, kk) in enumerate(linhas_imp):
        a.c('G%d' % (5 + i), rot); a.c('H%d' % (5 + i), vs_m[kk][0], vs_m[kk][1], fmt='brl', ed=vs_m[kk][2])
    carga_m = _r(sum(vs_m[kk][0] for _, kk in linhas_imp))
    fat = _r(m3 + q3)
    for col, vcol, cg in (('D', 'E', carga_s), ('G', 'H', carga_m)):
        a.c('%s14' % col, 'CARGA TRIBUTÁRIA', st='pag'); a.c('%s14' % vcol, cg, '=SUM(%s5:%s12)' % (vcol, vcol), st='pag', fmt='brl')
        a.c('%s15' % col, '%', st='pag'); a.c('%s15' % vcol, cg / fat if fat else 0, '=IF(E3+H3=0,0,%s14/(E3+H3))' % vcol, st='pag', fmt='pct')
    # imposto × empresa (mesmas linhas dos quadros)
    ret_tom = c['ret_tomados_manual'] if c.get('ret_tomados_manual') is not None else k.get('ret_a_recolher', 0)
    comp_flags = c.get('compensa') or {'pis': True, 'cofins': True, 'icms': True, 'retencao': True}
    a.c('J2', 'IMPOSTO POR EMPRESA', st='tit'); a.m('J2:L2')
    a.c('J3', 'IMPOSTO', st='cab'); a.c('K3', 'MATRIZ', st='cab'); a.c('L3', 'SCP', st='cab')
    tab = [(rot, kk, vs_m[kk][0], '=H%d' % (5 + i), vs_s[kk][0], '=E%d' % (5 + i)) for i, (rot, kk) in enumerate(linhas_imp)] + [('RETENÇÃO', 'retencao', ret_tom, None, 0, None)]
    pag_m = 0.0
    refs_pag = []
    for i, (rot, kk, vm, fm, vs_, fs) in enumerate(tab):
        r = 5 + i
        cf = comp_flags.get('irpj' if kk == 'adic' else kk)
        a.c('J%d' % r, rot + (' (compensa)' if cf else '')); a.c('K%d' % r, vm, fm, fmt='brl', ed='ret_tomados_manual' if kk == 'retencao' else None); a.c('L%d' % r, vs_, fs, fmt='brl')
        if not cf:
            pag_m += vm; refs_pag.append('K%d' % r)
    a.c('J14', 'PAGAMENTO DE FATO', st='pag'); a.c('K14', _r(pag_m), ('=' + '+'.join(refs_pag)) if refs_pag else None, st='pag', fmt='brl')
    a.c('L14', _r(pis_s + cof_s + irpj_s + adic_s + csll_s), '=SUM(L5:L12)', st='pag', fmt='brl')
    # fechamento
    carga_comp = _r(fat * pct_ant)
    carga_est = _r(carga_s + carga_m)
    a.c('A18', 'FECHAMENTO %s - %s' % (nome_mes.upper(), ano), st='tit'); a.m('A18:B18')
    a.c('A19', 'Faturamento', b=True); a.c('B19', fat, '=E3+H3', fmt='brl', b=True)
    a.c('A20', '% carga do regime anterior'); a.c('B20', pct_ant, '=B15', fmt='pct', ed='@ant_carga_pct')
    a.c('A21', 'Carga comparativa', b=True); a.c('B21', carga_comp, '=B19*B20', fmt='brl', b=True)
    a.c('A22', 'Carga com estratégia', b=True); a.c('B22', carga_est, '=E14+H14', fmt='brl', b=True)
    a.c('A23', '% carga com estratégia'); a.c('B23', carga_est / fat if fat else 0, '=IF(B19=0,0,B22/B19)', fmt='pct')
    a.c('A24', 'ECONOMIA TRIBUTÁRIA', st='pag'); a.c('B24', _r(carga_comp - carga_est), '=B21-B22', st='pag', fmt='brl')
    pct_ex = (c.get('exito_pct') if c.get('exito_pct') is not None else 15.0) / 100
    a.c('A25', '% honorário de êxito'); a.c('B25', pct_ex, fmt='pct', ed='@exito_pct')
    a.c('A26', 'HONORÁRIO DE ÊXITO', st='totg'); a.c('B26', _r(max(carga_comp - carga_est, 0) * pct_ex), '=MAX(B24,0)*B25', st='totg', fmt='brl')
    abas.append(a)

    # ============================================================ Fechamento Trimestral
    # economia mês a mês do trimestre e o somatório (como a aba "Fechamento Geral" da planilha do escritório)
    a = Aba('Fechamento Trimestral', {'A': 24, 'B': 17, 'C': 17, 'D': 17, 'E': 17, 'F': 17, 'G': 17, 'H': 17})
    tri_n = (mes - 1) // 3 + 1
    ini_m = (tri_n - 1) * 3 + 1
    outros_t = {x.get('competencia'): x for x in (c.get('tri_outros_meses') or (k.get('auto') or {}).get('tri_outros_meses') or [])}
    a.c('A2', 'ECONOMIA DO %d° TRIMESTRE %s' % (tri_n, ano), st='tit'); a.m('A2:H2')
    for col, h in zip('ABCDEFGH', ('MÊS', 'FAT. SCP', 'FAT. MATRIZ', 'FATURAMENTO', 'CARGA COMPARATIVA', 'CARGA C/ ESTRATÉGIA', 'ECONOMIA', 'HONORÁRIO DE ÊXITO')):
        a.c('%s3' % col, h, st='cab')
    faltam = []
    r0 = 4
    for i in range(3):
        mm = ini_m + i
        r = r0 + i
        cp = '%s-%02d' % (ano, mm)
        a.c('A%d' % r, MESES[mm - 1].upper(), b=True)
        if mm == mes:
            fs, fm, ce = m3, q3, carga_est
            a.c('B%d' % r, fs, "='Economia'!E3", fmt='brl'); a.c('C%d' % r, fm, "='Economia'!H3", fmt='brl'); a.c('F%d' % r, ce, "='Economia'!B22", fmt='brl')
        elif mm < mes:
            x = outros_t.get(cp) or {}
            fs, fm = x.get('receita') or 0, x.get('receita_matriz') or 0
            ce = x.get('carga_est')
            if not x or ce is None:
                faltam.append(MESES[mm - 1])
            ce = ce or 0
            a.c('B%d' % r, fs, fmt='brl', ed='#trim:%s:receita' % cp); a.c('C%d' % r, fm, fmt='brl', ed='#trim:%s:receita_matriz' % cp); a.c('F%d' % r, ce, fmt='brl', ed='#trim:%s:carga_est' % cp)
        else:
            fs = fm = ce = 0.0
            a.c('B%d' % r, None, fmt='brl'); a.c('C%d' % r, None, fmt='brl'); a.c('F%d' % r, None, fmt='brl')
        fat_i = _r(fs + fm)
        comp_i = _r(fat_i * pct_ant)
        a.c('D%d' % r, fat_i, '=B%d+C%d' % (r, r), fmt='brl')
        a.c('E%d' % r, comp_i, '=D%d*$B$16' % r, fmt='brl')
        a.c('G%d' % r, _r(comp_i - ce), '=E%d-F%d' % (r, r), fmt='brl')
        a.c('H%d' % r, _r(max(comp_i - ce, 0) * pct_ex), '=MAX(G%d,0)*$B$21' % r, fmt='brl')
    rt = r0 + 3
    a.c('A%d' % rt, 'TOTAL DO TRIMESTRE', st='pag')
    for col in 'BCDEFGH':
        tot = sum((a.cel.get('%s%d' % (col, r0 + i)) or {}).get('v') or 0 for i in range(3))
        a.c('%s%d' % (col, rt), _r(tot), '=SUM(%s%d:%s%d)' % (col, r0, col, rt - 1), st='pag', fmt='brl')
    fat_t = a.cel['D%d' % rt]['v']; est_t = a.cel['F%d' % rt]['v']
    comp_t = _r(fat_t * pct_ant)
    a.c('A13', 'FECHAMENTO %d° TRIMESTRE %s' % (tri_n, ano), st='tit'); a.m('A13:B13')
    a.c('A14', 'Faturamento do trimestre', b=True); a.c('B14', fat_t, '=D%d' % rt, fmt='brl', b=True)
    a.c('A16', '% carga comparativa'); a.c('B16', pct_ant, "='Economia'!B20", fmt='pct')
    a.c('A17', 'Carga comparativa', b=True); a.c('B17', comp_t, '=B14*B16', fmt='brl', b=True)
    a.c('A18', 'Carga com estratégia', b=True); a.c('B18', est_t, '=F%d' % rt, fmt='brl', b=True)
    a.c('A19', '% carga com estratégia'); a.c('B19', est_t / fat_t if fat_t else 0, '=IF(B14=0,0,B18/B14)', fmt='pct')
    a.c('A20', 'ECONOMIA TRIBUTÁRIA', st='pag'); a.c('B20', _r(comp_t - est_t), '=B17-B18', st='pag', fmt='brl')
    a.c('A21', '% honorário de êxito'); a.c('B21', pct_ex, "='Economia'!B25", fmt='pct')
    a.c('A22', 'HONORÁRIO DE ÊXITO', st='totg'); a.c('B22', _r(max(comp_t - est_t, 0) * pct_ex), '=MAX(B20,0)*B21', st='totg', fmt='brl')
    if faltam:
        a.c('D14', 'Sem a carga com estratégia de: %s. Use "Importar meses anteriores" ou digite na linha do mês.' % ', '.join(faltam), st='nota')
    abas.append(a)

    # ============================================================ Icms (só quando há NF-e/CT-e)
    if res.get('tem_icms') is False:
        for nome, mapa in DT.items():
            aba = next((x for x in abas if x.nome == nome), None)
            for ref, key in (mapa.items() if aba else []):
                if ref in aba.cel:
                    aba.cel[ref]['dt'] = key
        mover_colunas(abas, 'Pis.Cofins', {'B': 'A', 'C': 'B', 'D': 'C', 'E': 'D', 'F': 'E', 'J': 'G', 'K': 'H', 'L': 'I', 'M': 'J'}, {'F': 3})
        recalcular(abas, (c.get('cel_override') or {}))
        return abas
    a = Aba('Icms', {'A': 40, 'B': 17, 'C': 16, 'D': 17, 'E': 16, 'F': 16, 'G': 18})
    ap = res.get('apuracao') or {}
    a.c('A1', 'Registro de Apuração do ICMS', b=True)
    a.c('A2', 'FIRMA'); a.c('B2', emp)
    a.c('A3', 'CPF/CNPJ'); a.c('B3', (res.get('empresa') or {}).get('cnpj', ''))
    a.c('A5', 'MÊS OU PERÍODO/ANO'); a.c('B5', '%s/%s' % (comp[5:7], ano))
    cab = ['Codificação Contábil-Fiscal', 'Valores Contábeis', 'Base de Cálculo', 'Imposto Creditado/Debitado', 'Isentas / Não Trib.', 'Outras']
    cred_cf = {x['cfop'] for x in k['linhas_credito'] if x['inclui']}
    r = 8
    tot_rows = {}
    for grp, rot in (('entrada', 'ENTRADAS'), ('saida', 'SAÍDAS')):
        a.c('A%d' % r, rot, b=True); r += 1
        for j, h in enumerate(cab):
            a.c('%s%d' % ('ABCDEF'[j], r), h, st='cab')
        if grp == 'entrada':
            a.c('G%d' % r, 'VC − ICMS (crédito PIS/COFINS)', st='cab')
        r += 1
        ini_r = r
        for l in [x for x in ap.get('linhas', []) if x['grupo'] == grp]:
            a.c('A%d' % r, '%s %s' % (l['cfop'], l.get('desc') or desc_cfop(l['cfop'])))
            for col, kk in zip('BCDEF', ('vc', 'bc_icms', 'icms', 'isentas', 'outras')):
                a.c('%s%d' % (col, r), _r(l[kk])); a.cel['%s%d' % (col, r)]['dt'] = 'apu:' + l['cfop']
            if grp == 'entrada' and l['cfop'] in cred_cf:
                a.c('G%d' % r, _r(l['vc'] - l['icms']), '=B%d-D%d' % (r, r))
            r += 1
        fim_r = r - 1
        a.c('A%d' % r, 'Total', st='pag')
        for col, kk in zip('BCDEF', ('vc', 'bc_icms', 'icms', 'isentas', 'outras')):
            a.c('%s%d' % (col, r), _r((ap.get('totais') or {}).get(grp, {}).get(kk, 0)), '=SUM(%s%d:%s%d)' % (col, ini_r, col, fim_r) if fim_r >= ini_r else None, st='pag')
        if grp == 'entrada':
            gv = _r(sum(l['vc'] - l['icms'] for l in ap.get('linhas', []) if l['grupo'] == 'entrada' and l['cfop'] in cred_cf))
            a.c('G%d' % r, gv, '=SUM(G%d:G%d)' % (ini_r, fim_r) if fim_r >= ini_r else None, st='pag')
        tot_rows[grp] = r
        r += 2
    estado = (apu or {}).get('estado') or {}
    saldo_ant = c['icms_saldo_manual'] if c.get('icms_saldo_manual') is not None else _num(estado.get('saldo_icms'))
    te, ts = tot_rows['entrada'], tot_rows['saida']
    cred_tot = (ap.get('totais') or {}).get('entrada', {}).get('icms', 0)
    deb_tot = (ap.get('totais') or {}).get('saida', {}).get('icms', 0)
    a.c('A%d' % r, 'Total Crédito', b=True); a.c('B%d' % r, _r(cred_tot), '=D%d' % te, b=True); rc = r; r += 1
    a.c('A%d' % r, 'Saldo Credor ', b=True); a.c('B%d' % r, _r(saldo_ant), b=True, ed='icms_saldo_manual'); rs_ = r; r += 1
    a.c('A%d' % r, 'Total Débitos', b=True); a.c('B%d' % r, _r(deb_tot), '=D%d' % ts, b=True); rd = r; r += 1
    aj = apc.get('ajustes') or []
    aj_cred = sum(x['valor'] for x in aj if x['tipo'].startswith(('Outros créditos', 'Estorno de débito')))
    aj_deb = sum(x['valor'] for x in aj if x['tipo'].startswith(('Outros débitos', 'Estorno de crédito')))
    raj = None
    if aj_cred or aj_deb:
        a.c('A%d' % r, 'Ajustes (aba Apuração): créditos − débitos', b=True); a.c('B%d' % r, _r(aj_cred - aj_deb), b=True); raj = r; r += 1
    a.c('A%d' % r, 'Total a Pagar ', st='pag')
    a.c('B%d' % r, _r(cred_tot + saldo_ant - deb_tot + (aj_cred - aj_deb)), '=B%d+B%d-B%d%s' % (rc, rs_, rd, ('+B%d' % raj) if raj else ''), st='pag'); a.c('C%d' % r, 'ICMS ', st='nota')
    r += 1
    vendas = [x for x in linhas_deb if x['sinal'] > 0]
    ib = r
    for x in vendas:
        a.c('A%d' % r, int(x['cfop'])); a.c('B%d' % r, _r(x['vc'] - x['icms'])); a.cel['B%d' % r]['dt'] = 'apu:' + x['cfop']
        r += 1
    a.c('A%d' % r, 'Total Base Saída Produtos', b=True); a.c('B%d' % r, _r(sum(x['vc'] - x['icms'] for x in vendas)), '=SUM(B%d:B%d)' % (ib, r - 1) if vendas else None, b=True); a.c('C%d' % r, 'Redução ICMS ', st='nota')
    r += 1
    a.c('A%d' % r, 'Faturamento', b=True); a.c('B%d' % r, prod_vc, b=True); a.cel['B%d' % r]['dt'] = 'prod_fat'
    a.cel['A%d' % rs_]['dt'] = a.cel['B%d' % rs_]['dt'] = 'saldo_icms'
    abas.append(a)
    for nome, mapa in DT.items():
        aba = next((x for x in abas if x.nome == nome), None)
        for ref, key in (mapa.items() if aba else []):
            if ref in aba.cel:
                aba.cel[ref]['dt'] = key
    mover_colunas(abas, 'Pis.Cofins', {'B': 'A', 'C': 'B', 'D': 'C', 'E': 'D', 'F': 'E', 'J': 'G', 'K': 'H', 'L': 'I', 'M': 'J'}, {'F': 3})
    recalcular(abas, (c.get('cel_override') or {}))
    return abas


def mover_colunas(abas, nome, mapa, larg_extra=None):
    """Muda as colunas de uma aba (células, mesclagens, larguras e todas as fórmulas que apontam para ela, inclusive de outras abas)."""
    a = next((x for x in abas if x.nome == nome), None)
    if not a:
        return
    def mv(ref):
        m = re.match(r'(\$?)([A-Z]+)(\$?\d+)$', ref)
        return m.group(1) + mapa.get(m.group(2), m.group(2)) + m.group(3) if m else ref
    def mv_f(f, aba_da_formula):
        def sub(m):
            sh = m.group(1)
            if (sh or aba_da_formula) != nome:
                return m.group(0)
            r1 = mv(m.group(2)); r2 = mv(m.group(3)) if m.group(3) else ''
            return ("'%s'!" % sh if sh else '') + r1 + (':' + r2 if r2 else '')
        return '=' + _REF.sub(sub, f[1:])
    for x in abas:
        for cc in x.cel.values():
            for kf in ('f', 'f0'):
                if isinstance(cc.get(kf), str) and cc[kf].startswith('='):
                    cc[kf] = mv_f(cc[kf], x.nome)
    a.cel = {mv(r): cc for r, cc in a.cel.items()}
    a.mesc = [':'.join(mv(p) for p in rng.split(':')) for rng in a.mesc]
    larg = {}
    for col, w in a.larg.items():
        larg[mapa.get(col, col)] = w
    larg.update(larg_extra or {})
    a.larg = larg


# detalhamento de cada célula (clicar → Detalhar)
DT = {'Pis.Cofins': {'F4': 'cred_prod', 'F5': 'tomados', 'F6': 'importacao', 'F9': 'saldo_pis', 'F10': 'saldo_cof', 'F14': 'serv', 'C14': 'serv', 'K3': 'serv', 'K9': 'serv', 'L9': 'serv',
                     'F15': 'loc', 'C15': 'loc', 'K2': 'loc', 'K10': 'loc', 'L10': 'loc', 'F16': 'prod_saida', 'F17': 'devol', 'K4': 'prod_fat', 'K11': 'prod_fat',
                     'F28': 'ret', 'F29': 'ret', 'C28': 'ret', 'C29': 'ret'},
      'Irpj.Csll': {'B11': 'ret', 'B14': 'ret', 'C11': 'ret', 'C14': 'ret', 'B5': 'scp_fat', 'G5': 'mat_fat'},
      'Economia': {}}


def detalhes(res, k):
    """Composição de cada campo da planilha: CFOP considerados (com as notas), notas de serviço, locações, retenções."""
    c = k['cfg']
    def cf_lin(x, inclui=True):
        return {'cfop': x['cfop'], 'desc': x.get('desc') or desc_cfop(x['cfop']), 'vc': _r(x['vc']), 'icms': _r(x['icms']), 'ipi': _r(x.get('ipi', 0)), 'st': _r(x.get('st', 0)),
                'base': _r(x['base']), 'notas': x.get('notas') or [], 'inclui': inclui, 'padrao': x.get('padrao', True)}
    deb, fora_deb = k['linhas_debito'], k.get('linhas_debito_fora') or []
    d = {}
    d['prod_saida'] = {'titulo': 'Saídas de produtos: CFOP considerados', 'tipo': 'cfop', 'toggle': 'debito_cfop',
                       'obs': 'Base = valor contábil − ICMS%s%s. Desmarque um CFOP para tirá-lo da base de débito, ou marque um que ficou de fora.' % (' − IPI' if c.get('excluir_ipi_debito') else '', ' − ICMS-ST' if c.get('excluir_st_debito') else ''),
                       'linhas': [cf_lin(x) for x in deb if x['sinal'] > 0] + [cf_lin(x, False) for x in fora_deb if x['sinal'] > 0]}
    d['devol'] = {'titulo': 'Devoluções de venda: CFOP considerados', 'tipo': 'cfop', 'toggle': 'debito_cfop',
                  'obs': 'Devolução de venda reduz a base de débito (Lei 10.833/03, art. 1º, §3º, V, b).',
                  'linhas': [cf_lin(x) for x in deb if x['sinal'] < 0] + [cf_lin(x, False) for x in fora_deb if x['sinal'] < 0]}
    d['prod_fat'] = {'titulo': 'Faturamento de produtos (valor contábil das vendas)', 'tipo': 'cfop', 'toggle': 'debito_cfop',
                     'obs': 'Soma do valor contábil dos CFOP de venda considerados.', 'linhas': d['prod_saida']['linhas']}
    d['cred_prod'] = {'titulo': 'Entradas de produtos: CFOP que geram crédito', 'tipo': 'cfop', 'toggle': 'creditos_cfop',
                      'obs': 'Base = valor contábil − ICMS%s. Uso e consumo, ativo e remessas ficam fora por padrão; marque para incluir.' % (' − IPI' if c.get('excluir_ipi_credito') else ''),
                      'linhas': [cf_lin(x) for x in k['linhas_credito']] + [cf_lin(x, False) for x in k.get('linhas_fora') or []]}
    imp = k.get('importacao') or {}
    d['importacao'] = {'titulo': 'Importação', 'tipo': 'texto', 'texto': ['Modo: %s' % {'pago': 'PIS/COFINS-Importação pago (destacado nas NF-e)', 'base': 'base sobre o valor contábil', 'manual': 'informado'}.get(imp.get('modo'), imp.get('modo') or ''),
                                                                     'PIS-Importação: %s · COFINS-Importação: %s' % (_b(imp.get('pis')), _b(imp.get('cofins'))), 'Valor contábil dos CFOP 3xxx: %s' % _b(imp.get('vc') or imp.get('base'))]}
    excl = set(c.get('tomados_excluir') or [])
    d['tomados'] = {'titulo': 'Serviços tomados (crédito de insumo)', 'tipo': 'lista', 'toggle': 'tomados_excluir',
                    'obs': 'Pessoa física não gera crédito. Desmarque um prestador para tirá-lo do crédito.',
                    'colunas': [('numero', 'Nota', 'txt'), ('data', 'Data', 'data'), ('participante', 'Prestador', 'txt'), ('cnpj', 'CNPJ/CPF', 'doc'), ('valor', 'Valor', 'brl'), ('fonte', 'Fonte', 'txt')],
                    'linhas': [dict({kk: t.get(kk) for kk in ('numero', 'data', 'participante', 'cnpj', 'valor', 'fonte')},
                                    inclui=not (t.get('cnpj') in excl or (c.get('excluir_pf') and (t.get('pf') or len(t.get('cnpj') or '') == 11))), chave=t.get('cnpj') or '') for t in k.get('tomados') or []]}
    d['serv'] = {'titulo': 'Prestação de serviços (NFS-e emitidas)', 'tipo': 'lista',
                 'colunas': [('numero', 'Nota', 'txt'), ('data', 'Data', 'data'), ('participante', 'Tomador', 'txt'), ('cnpj', 'CNPJ/CPF', 'doc'), ('valor', 'Valor', 'brl'), ('iss', 'ISS', 'brl'), ('fonte', 'Fonte', 'txt')],
                 'linhas': [{kk: t.get(kk) for kk in ('numero', 'data', 'participante', 'cnpj', 'valor', 'iss', 'fonte')} for t in k.get('prestados') or []],
                 'obs': 'Distribuição: Matriz %s · SCP %s (modo %s).' % (_b(k['resumo']['distribuicao']['serv_matriz']), _b(k['resumo']['distribuicao']['serv_scp']), k['resumo']['distribuicao']['modo'])}
    locs = [x for x in (res.get('locacoes') or {}).get('linhas', []) if x.get('tipo') == 'receita']
    d['loc'] = {'titulo': 'Locação (relatório do sistema)', 'tipo': 'lista',
                'colunas': [('numero', 'Nota', 'txt'), ('participante', 'Cliente', 'txt'), ('cnpj', 'CNPJ/CPF', 'doc'), ('valor', 'Valor', 'brl')],
                'linhas': [{kk: x.get(kk) for kk in ('numero', 'participante', 'cnpj', 'valor')} for x in locs],
                'obs': 'Distribuição: Matriz %s · SCP %s.' % (_b(k['resumo']['distribuicao']['loc_matriz']), _b(k['resumo']['distribuicao']['loc_scp']))}
    d['ret'] = {'titulo': 'Retenções nas notas de serviço prestado', 'tipo': 'lista',
                'colunas': [('numero', 'Nota', 'txt'), ('participante', 'Tomador', 'txt'), ('valor', 'Valor', 'brl'), ('pis', 'PIS', 'brl'), ('cofins', 'COFINS', 'brl'), ('irrf', 'IRRF', 'brl'), ('csll', 'CSLL', 'brl')],
                'linhas': [{kk: t.get(kk) for kk in ('numero', 'participante', 'valor', 'pis', 'cofins', 'irrf', 'csll')} for t in k.get('prestados') or [] if any((t.get(x) or 0) for x in ('pis', 'cofins', 'irrf', 'csll'))],
                'obs': 'Retenção vai para quem emitiu a nota: Matriz ou SCP conforme a distribuição.'}
    auto = k.get('auto') or {}
    for kk, nome, ck in (('saldo_pis', 'PIS', 'saldo_pis_ant'), ('saldo_cof', 'COFINS', 'saldo_cof_ant')):
        orig = 'informado pelo ADM' if c.get('saldo_manual') else ('fechamento de %s' % auto['saldo_origem'] if auto.get('saldo_origem') else 'sem fechamento anterior guardado (zero)')
        d[kk] = {'titulo': 'Saldo credor de %s de períodos anteriores' % nome, 'tipo': 'texto', 'texto': ['Valor: %s' % _b(c.get(ck)), 'Origem: %s' % orig]}
    d['saldo_icms'] = {'titulo': 'Saldo credor de ICMS do período anterior', 'tipo': 'texto', 'texto': ['Vem da aba Apuração ICMS/IPI (campo "Saldo credor do período anterior") ou do valor digitado aqui pelo ADM.']}
    for l in (res.get('apuracao') or {}).get('linhas', []):
        d['apu:' + l['cfop']] = {'titulo': 'CFOP %s %s: CST e notas' % (l['cfop'], l.get('desc') or desc_cfop(l['cfop'])), 'tipo': 'lista',
                                 'colunas': [('cst', 'CST', 'txt'), ('desc', 'Descrição', 'txt'), ('vc', 'Valor contábil', 'brl'), ('bc_icms', 'Base ICMS', 'brl'), ('icms', 'ICMS', 'brl'), ('ipi', 'IPI', 'brl'), ('notas', 'Notas', 'notas')],
                                 'linhas': [{kk: cs.get(kk) for kk in ('cst', 'desc', 'vc', 'bc_icms', 'icms', 'ipi', 'notas')} for cs in l.get('csts', [])]}
    return d


# ------------------------------------------------------------------ recálculo das fórmulas (valor digitado em qualquer célula)
_REF = re.compile(r"(?:'([^']+)'!)?(\$?[A-Z]{1,2}\$?\d+)(?::(\$?[A-Z]{1,2}\$?\d+))?")


def _col(l):
    n = 0
    for ch in l:
        n = n * 26 + ord(ch) - 64
    return n


def _letra(n):
    t = ''
    while n:
        m = (n - 1) % 26; t = chr(65 + m) + t; n = (n - 1) // 26
    return t


def _expr(f, aba):
    """Traduz a fórmula do Excel (SUM, MAX, IF, %, refs e intervalos, inclusive de outra aba) para Python."""
    out, i = [], 0
    corpo = f[1:]
    for m in _REF.finditer(corpo):
        out.append(corpo[i:m.start()])
        sh = m.group(1) or aba
        a1, a2 = m.group(2).replace('$', ''), (m.group(3) or '').replace('$', '')
        out.append('_rg(%r,%r,%r)' % (sh, a1, a2) if a2 else '_v(%r,%r)' % (sh, a1))
        i = m.end()
    out.append(corpo[i:])
    e = ''.join(out)
    e = re.sub(r'(\d+(?:\.\d+)?)%', r'(\1/100)', e)
    e = e.replace('<>', '!=')
    e = re.sub(r'(?<![<>!=])=(?!=)', '==', e)
    e = re.sub(r'\bSUM\(', '_sum(', e); e = re.sub(r'\bMAX\(', '_max(', e); e = re.sub(r'\bMIN\(', '_min(', e); e = re.sub(r'\bIF\(', '_if(', e)
    return e


def recalcular(abas, ov):
    """Aplica os valores digitados ({'Aba!REF': valor}) e recalcula as fórmulas da planilha, como o Excel faria."""
    por = {a.nome: a for a in abas}
    for chave, v in (ov or {}).items():
        nome, _, ref = chave.partition('!')
        a = por.get(nome)
        if not a or v is None:
            continue
        cel = a.cel.setdefault(ref, {'v': None, 'f': None, 'st': 'lin', 'fmt': 'num', 'b': False})
        cel['calc'] = cel.get('v'); cel['v'] = v; cel['ov'] = True; cel['f0'] = cel.get('f'); cel['f'] = None
    if not ov:
        return
    def _v(sh, ref):
        cc = por[sh].cel.get(ref) if sh in por else None
        v = cc.get('v') if cc else None
        return v if isinstance(v, (int, float)) and not isinstance(v, bool) else 0.0
    def _rg(sh, a1, a2):
        m1, m2 = re.match(r'([A-Z]+)(\d+)', a1), re.match(r'([A-Z]+)(\d+)', a2)
        c1, c2, r1, r2 = _col(m1.group(1)), _col(m2.group(1)), int(m1.group(2)), int(m2.group(2))
        return [_v(sh, '%s%d' % (_letra(cc), rr)) for rr in range(r1, r2 + 1) for cc in range(c1, c2 + 1)]
    def _flat(xs):
        for x in xs:
            if isinstance(x, list):
                yield from x
            else:
                yield x
    env = {'_v': _v, '_rg': _rg, '_sum': lambda *xs: sum(_flat(xs)), '_max': lambda *xs: max(_flat(xs)), '_min': lambda *xs: min(_flat(xs)),
           '_if': lambda c_, a_, b_: a_ if c_ else b_, '__builtins__': {}}
    for _ in range(12):
        mudou = False
        for a in abas:
            for ref, cc in a.cel.items():
                f = cc.get('f')
                if not f or not isinstance(f, str) or not f.startswith('='):
                    continue
                try:
                    v = round(float(eval(_expr(f, a.nome), env)), 6)
                except Exception:
                    continue
                if cc.get('fmt') != 'pct':
                    v = round(v, 2)
                if not isinstance(cc.get('v'), (int, float)) or abs((cc.get('v') or 0) - v) > 1e-6:
                    cc['v'] = v; mudou = True
        if not mudou:
            break


def economia_do_mes(planilha):
    """Do JSON da planilha de um fechamento guardado: carga comparativa, carga com estratégia, economia e honorário do mês."""
    ec = next((p for p in planilha or [] if p.get('nome') == 'Economia'), None)
    if not ec:
        return {}
    cel = ec.get('cel') or {}
    out = {}
    for ref, cc in cel.items():
        t = str(cc.get('v') or '').strip().lower()
        alvo = {'carga comparativa': 'carga_comp', 'carga com estratégia': 'carga_est', 'economia tributária': 'economia', 'honorário de êxito': 'honorario'}.get(t)
        if alvo:
            m = re.match(r'([A-Z]+)(\d+)$', ref)
            v = (cel.get('B' + m.group(2)) or {}).get('v') if m else None
            if isinstance(v, (int, float)):
                out[alvo] = round(v, 2)
    return out


def _num(v):
    if v in (None, ''):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    t = str(v).strip()
    try:
        return float(t.replace('.', '').replace(',', '.')) if ',' in t else float(t)
    except ValueError:
        return 0.0


def _b(v):
    return '{:,.2f}'.format(v or 0).replace(',', 'X').replace('.', ',').replace('X', '.')


def para_json(abas):
    return [{'nome': a.nome, 'larg': a.larg, 'cel': a.cel, 'mesc': a.mesc} for a in abas]


def xlsx(abas, titulo=''):
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    fina, media = Side(style='thin', color='7F7F7F'), Side(style='medium', color='0E2841')
    fmt = {'num': '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)', 'brl': '_-"R$"\\ * #,##0.00_-;\\-"R$"\\ * #,##0.00_-;_-"R$"\\ * "-"??_-;_-@_-', 'pct': '0.00%'}
    for a in abas:
        ws = wb.create_sheet(a.nome)
        for col, w in a.larg.items():
            ws.column_dimensions[col].width = w
        for ref, cc in a.cel.items():
            cel = ws[ref]
            cel.value = cc['f'] if cc.get('f') else cc['v']
            st = cc['st']
            if st in COR:
                bg, fg = COR[st]
                cel.fill = PatternFill('solid', fgColor=bg)
                cel.font = Font(bold=True, color=fg)
            elif st == 'cod':
                cel.font = Font(bold=True, color='FF0000', size=9)
            elif st == 'nota':
                cel.font = Font(italic=True, color='5B6478', size=9)
            else:
                cel.font = Font(bold=cc.get('b'))
            if st not in ('nota', 'cod') and cc['v'] is not None:
                cel.border = Border(left=fina, right=fina, top=media if st in ('tit', 'sub', 'totg') else fina, bottom=media if st in ('tit', 'sub', 'totg') else fina)
            if cc['fmt'] in fmt and cc['v'] is not None and not isinstance(cc['v'], str):
                cel.number_format = fmt[cc['fmt']]
            if st in ('tit', 'sub', 'cab'):
                cel.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        for rng in a.mesc:
            ws.merge_cells(rng)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _grade(a):
    """Converte a aba em linhas × colunas, com mesclagens (para HTML/PDF)."""
    from openpyxl.utils import column_index_from_string, get_column_letter, coordinate_to_tuple
    pos = {coordinate_to_tuple(r): v for r, v in a.cel.items()}
    if not pos:
        return [], []
    maxr = max(r for r, _ in pos); maxc = max(c for _, c in pos)
    usadas = sorted({c for _, c in pos})
    span, oculto = {}, set()
    for rng in a.mesc:
        i, f = rng.split(':')
        (r1, c1), (r2, c2) = coordinate_to_tuple(i), coordinate_to_tuple(f)
        span[(r1, c1)] = (r2 - r1 + 1, c2 - c1 + 1)
        for rr in range(r1, r2 + 1):
            for cc in range(c1, c2 + 1):
                if (rr, cc) != (r1, c1):
                    oculto.add((rr, cc))
    cols = list(range(min(usadas), maxc + 1))
    linhas = []
    for r in range(1, maxr + 1):
        linhas.append([(r, c, pos.get((r, c)), span.get((r, c)), (r, c) in oculto) for c in cols])
    larg = [a.larg.get(get_column_letter(c), 10) for c in cols]
    return linhas, larg


def _fmt(cc):
    v = cc['v']
    if v is None:
        return ''
    if isinstance(v, str) or cc['fmt'] == 'txt':
        return str(v)
    if cc['fmt'] == 'pct':
        return ('%.2f%%' % (v * 100)).replace('.', ',')
    if abs(v) < 0.005:
        return 'R$ -' if cc['fmt'] == 'brl' else '-'
    s = _b(abs(v))
    if cc['fmt'] == 'brl':
        return ('-R$ ' if v < 0 else 'R$ ') + s
    return '(%s)' % s if v < 0 else s


def html_aba(a, pdf=False):
    linhas, larg = _grade(a)
    tot = sum(larg) or 1
    out = ['<table style="border-collapse:collapse;width:100%%;font-size:%s">' % ('6.6pt' if pdf else '12px')]
    out.append('<tr>' + ''.join('<td style="width:%.1f%%;padding:0;border:0"></td>' % (100 * w / tot) for w in larg) + '</tr>')
    dono = {}
    if pdf:
        for row in linhas:
            for r, c, cc, sp, oc in row:
                if sp and cc:
                    for rr in range(r, r + sp[0]):
                        for c2 in range(c, c + sp[1]):
                            dono[(rr, c2)] = cc
    for row in linhas:
        if all(cc is None or oc for _, _, cc, _, oc in row) and not any((r, c) in dono for r, c, _, _, _ in row):
            out.append('<tr><td colspan="%d" style="height:%s;border:0"></td></tr>' % (len(row), '5px' if pdf else '8px'))
            continue
        cels = []
        for r, c, cc, sp, oc in row:
            if oc and pdf:
                st0 = dono.get((r, c), {}).get('st')
                bg = '#%s' % COR[st0][0] if st0 in COR else 'transparent'
                cels.append('<td style="background-color:%s;border-top:0.5px solid #7F7F7F;border-bottom:0.5px solid #7F7F7F"></td>' % bg)
                continue
            if oc:
                continue
            at = ''
            if sp and not pdf:
                at = ' rowspan="%d" colspan="%d"' % sp
            if cc is None:
                cels.append('<td%s style="border:0"></td>' % at)
                continue
            st = cc['st']
            css = 'padding:%s;' % ('2px 4px' if pdf else '4px 7px')
            if st in COR:
                bg, fg = COR[st]
                css += 'background-color:#%s;color:#%s;font-weight:bold;' % (bg, fg)
            elif st == 'cod':
                css += 'color:#E00000;font-weight:bold;font-size:%s;' % ('6pt' if pdf else '10.5px')
            elif st == 'nota':
                css += 'color:#5B6478;font-style:italic;font-size:%s;' % ('6pt' if pdf else '11px')
            elif cc.get('b'):
                css += 'font-weight:bold;'
            if st not in ('nota', 'cod'):
                css += 'border:%s solid #7F7F7F;' % ('0.5px' if pdf else '1px')
            num = not isinstance(cc['v'], str) and cc['v'] is not None and cc['fmt'] != 'txt'
            css += 'text-align:%s;' % ('center' if st in ('tit', 'sub', 'cab') else ('right' if num else 'left'))
            tip = '' if pdf or not cc.get('f') else ' title="%s"' % str(cc['f']).replace('"', '&quot;')
            cels.append('<td%s%s style="%s">%s</td>' % (at, tip, css, str(_fmt(cc)).replace('&', '&amp;').replace('<', '&lt;')))
        out.append('<tr>' + ''.join(cels) + '</tr>')
    out.append('</table>')
    return ''.join(out)


def pdf(abas, cab):
    import fitz, os
    buf = io.BytesIO()
    w = fitz.DocumentWriter(buf)
    mb = fitz.paper_rect('a4-l')
    area = mb + (24, 64, -24, -28)
    titulos = []
    for a in abas:
        story = fitz.Story('<body style="font-family:sans-serif"><p style="font-size:9pt;font-weight:bold;color:#0E2841;margin:0 0 6px 0">Aba %s</p>%s</body>' % (a.nome, html_aba(a, True)))
        mais = 1
        while mais:
            dev = w.begin_page(mb); mais, _ = story.place(area); story.draw(dev); w.end_page(); titulos.append(a.nome)
    w.close()
    doc = fitz.open('pdf', buf.getvalue())
    logo = os.path.join(os.path.dirname(__file__), '..', 'web', 'logo-colorida.png')
    n = len(doc)
    for i, pg in enumerate(doc, start=1):
        r = pg.rect
        if os.path.exists(logo):
            pg.insert_image(fitz.Rect(24, 16, 120, 44), filename=logo, keep_proportion=True)
        pg.insert_text((r.width - 360, 28), cab['titulo'], fontsize=10, color=(0.05, 0.16, 0.25), fontname='hebo')
        pg.insert_text((r.width - 360, 41), '%s · CNPJ %s · Competência %s' % (cab['empresa'][:46], cab['cnpj'], cab['competencia']), fontsize=7.5, color=(0.36, 0.39, 0.47))
        pg.draw_line((24, 54), (r.width - 24, 54), color=(0.84, 0.85, 0.89), width=0.6)
        pg.insert_text((24, r.height - 14), 'Marçal Contabilidade · Revisor do Fechamento', fontsize=7, color=(0.54, 0.58, 0.65))
        pg.insert_text((r.width - 90, r.height - 14), 'Página %d de %d' % (i, n), fontsize=7, color=(0.54, 0.58, 0.65))
    return doc.tobytes(deflate=True, garbage=3)
