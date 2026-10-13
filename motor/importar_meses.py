# -*- coding: utf-8 -*-
"""Lê planilhas de fechamento de meses anteriores (o Excel do escritório ou o exportado pela Estratégia SCP) e devolve,
por mês, o que o trimestre precisa: competência, faturamento da SCP e da Matriz e as retenções de IRRF/CSLL da SCP.

O layout varia de empresa para empresa, então a leitura é por rótulo, não por célula fixa:
- competência: título "PIS/COFINS - <Mês> <Ano>" (ou "FECHAMENTO <MÊS> <ANO>");
- faturamento: no quadro de resumo/economia, cabeçalho "LUCRO PRESUMIDO - SCP" (ou "LUCRO REAL ...") e a linha
  FATURAMENTO logo abaixo; cabeçalhos "APURAÇÃO ..." são das abas de PIS/COFINS e ficam de fora;
- retenções: tabela de notas com colunas IR/IRRF e CSLL (só as linhas da SCP, se houver a coluna Empresa) ou as
  linhas "(−) IRRF retido"/"(−) CSLL retida" da aba Irpj.Csll exportada."""
import io
import re
import unicodedata

MESES = {'janeiro': 1, 'fevereiro': 2, 'marco': 3, 'abril': 4, 'maio': 5, 'junho': 6, 'julho': 7, 'agosto': 8, 'setembro': 9,
         'outubro': 10, 'novembro': 11, 'dezembro': 12}
MAX_LIN, MAX_COL = 150, 40


def _n(t):
    t = unicodedata.normalize('NFKD', str(t or '')).encode('ascii', 'ignore').decode().lower()
    return re.sub(r'\s+', ' ', t).strip()


def _num(v):
    if isinstance(v, bool) or v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    t = str(v).strip().replace('R$', '').replace(' ', '')
    if not t or t in ('-',):
        return None
    try:
        return float(t.replace('.', '').replace(',', '.')) if ',' in t else float(t)
    except ValueError:
        return None


def _grade(ws):
    g = {}
    for row in ws.iter_rows(max_row=MAX_LIN, max_col=MAX_COL):
        for c in row:
            v = getattr(c, 'value', None)
            if v not in (None, ''):
                g[(c.row, c.column)] = v
    return g


def _competencia(texto):
    t = _n(texto)
    m = re.search(r'(?:pis/cofins|fechamento)\s*-?\s*(janeiro|fevereiro|marco|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro)\s*(?:de\s*)?(\d{4})', t)
    if m:
        return '%s-%02d' % (m.group(2), MESES[m.group(1)])
    return None


def _valor_abaixo(g, r, c, rotulo, ate=4):
    """Procura o rótulo nas linhas logo abaixo do cabeçalho (mesma coluna) e devolve o número à direita."""
    for rr in range(r + 1, r + 1 + ate):
        if rotulo in _n(g.get((rr, c), '')):
            for cc in range(c + 1, c + 4):
                v = _num(g.get((rr, cc)))
                if v is not None:
                    return v
            return 0.0
    return None


def ler(conteudo, nome=''):
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(conteudo), read_only=True, data_only=True)
    grades = [(ws.title, _grade(ws)) for ws in wb.worksheets]
    wb.close()
    out = {'arquivo': nome, 'competencia': None, 'receita': None, 'receita_matriz': None, 'irrf': 0.0, 'csll': 0.0, 'avisos': []}
    # competência: primeiro título encontrado, na ordem das abas
    for _, g in grades:
        for k in sorted(g):
            comp = _competencia(g[k]) if isinstance(g[k], str) else None
            if comp:
                out['competencia'] = comp
                break
        if out['competencia']:
            break
    # faturamento SCP × Matriz: primeiro quadro de resumo (cabeçalho sem "APURAÇÃO")
    for _, g in grades:
        scp = real = None
        for (r, c) in sorted(g):
            v = g[(r, c)]
            if not isinstance(v, str):
                continue
            t = _n(v)
            if t.startswith('apuracao') or t.startswith('regime anterior'):
                continue
            if scp is None and t.startswith('lucro presumido - scp'):
                scp = _valor_abaixo(g, r, c, 'faturamento')
            elif real is None and t.startswith('lucro real'):
                real = _valor_abaixo(g, r, c, 'faturamento')
        if scp is not None or real is not None:
            out['receita'] = scp if scp is not None else 0.0
            out['receita_matriz'] = real if real is not None else 0.0
            break
    # retenções da SCP
    achou_ret = False
    for nome_aba, g in grades:
        cab = None
        for (r, c) in sorted(g):
            t = _n(g[(r, c)]) if isinstance(g[(r, c)], str) else ''
            if t in ('ir', 'irrf'):
                linha = {cc: _n(g.get((r, cc), '')) for cc in range(1, MAX_COL + 1)}
                col_cs = next((cc for cc, x in linha.items() if x == 'csll'), None)
                if col_cs:
                    cab = (r, c, col_cs, next((cc for cc, x in linha.items() if x.startswith('empresa')), None))
                    break
        if cab:
            r0, c_ir, c_cs, c_emp = cab
            for rr in range(r0 + 1, MAX_LIN + 1):
                if any('total' in _n(g.get((rr, cc), '')) for cc in range(1, 6)):
                    break
                if c_emp and 'scp' not in _n(g.get((rr, c_emp), '')):
                    continue
                out['irrf'] += _num(g.get((rr, c_ir))) or 0
                out['csll'] += _num(g.get((rr, c_cs))) or 0
            achou_ret = True
            break
        if _n(nome_aba) == 'irpj.csll':
            for (r, c) in sorted(g):
                t = _n(g[(r, c)]) if isinstance(g[(r, c)], str) else ''
                if 'irrf retido' in t:
                    out['irrf'] = _num(g.get((r, c + 1))) or 0; achou_ret = True
                elif 'csll retida' in t:
                    out['csll'] = abs(_num(g.get((r, c + 1))) or 0); achou_ret = True
            if achou_ret:
                break
    out['irrf'], out['csll'] = round(out['irrf'], 2), round(out['csll'], 2)
    # quadro de fechamento do mês (economia): primeira ocorrência de cada rótulo
    rot = {'carga_comp': 'carga comparativa', 'carga_est': 'carga com estrategia', 'economia': 'economia tributaria', 'honorario': 'honorario'}
    for _, g in grades:
        for (r, c) in sorted(g):
            v = g[(r, c)]
            if not isinstance(v, str):
                continue
            t = _n(v)
            for k, alvo in rot.items():
                if k in out or not t.startswith(alvo):
                    continue
                for cc in range(c + 1, c + 4):
                    x = _num(g.get((r, cc)))
                    if x is not None:
                        out[k] = round(x, 2)
                        break
        if 'carga_est' in out:
            break
    if not out['competencia']:
        out['avisos'].append('Não achei o mês (título "PIS/COFINS - <Mês> <Ano>").')
    if out['receita'] is None:
        out['avisos'].append('Não achei o quadro com "LUCRO PRESUMIDO - SCP" e "FATURAMENTO".')
    if not achou_ret and out['receita']:
        out['avisos'].append('Sem tabela de retenções: IRRF e CSLL retidos da SCP ficaram zerados (confira em Parâmetros).')
    return out
