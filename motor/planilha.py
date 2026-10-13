# -*- coding: utf-8 -*-
"""Leitura da planilha do sistema do cliente (xlsx, xls, csv): acha a linha de cabeçalho, reconhece as
colunas por sinônimos e devolve os lançamentos. O layout varia por cliente; o que não for reconhecido
o usuário ajusta na tela e o mapeamento fica salvo em config/mapeamentos.json (chave = assinatura do cabeçalho)."""
import os, re, csv, json, io, hashlib
from .util import norm, num, digitos, data_iso

CAMPOS = [
    # campo, rótulo na tela, sinônimos (já normalizados)
    ('movimento', 'Entrada/Saída', ['tipo', 'movimento', 'e s', 'entrada saida', 'operacao', 'tipo movimento', 'es', 'tipo nota', 'tipo operacao']),
    ('chave', 'Chave de acesso', ['chave', 'chave acesso', 'chave de acesso', 'chave nfe', 'chave nf e', 'chave eletronica', 'chave do documento', 'chave cte']),
    ('numero', 'Número da nota', ['numero', 'n', 'no', 'num', 'nro', 'nota', 'nf', 'n nota', 'numero nota', 'nro nota', 'recibo', 'n recibo', 'numero recibo', 'contrato', 'n contrato', 'numero contrato', 'fatura', 'n fatura', 'num nota', 'numero nf', 'nro nf', 'nro documento', 'n documento', 'numero documento', 'documento', 'doc', 'n doc', 'nota fiscal', 'numero da nota', 'n da nota', 'nfse', 'n nf']),
    ('serie', 'Série', ['serie', 'ser', 'serie nf']),
    ('modelo', 'Modelo/Espécie', ['modelo', 'mod', 'especie', 'esp']),
    ('data_emissao', 'Data de emissão', ['emissao', 'data emissao', 'dt emissao', 'data de emissao', 'dt emis', 'data documento', 'data doc', 'dt doc', 'dt documento', 'data do documento']),
    ('data_lanc', 'Data de entrada/saída', ['data', 'entrada', 'saida', 'data entrada', 'data saida', 'dt entrada', 'dt saida', 'data lancamento', 'lancamento', 'data e s', 'data entrada saida', 'dt lancto', 'data movimento', 'dt movimento', 'dt entrada saida', 'data escrituracao', 'vencimento', 'data vencimento', 'data pagamento', 'data recebimento']),
    ('cnpj', 'CNPJ/CPF participante', ['cnpj', 'cpf', 'cnpj cpf', 'cpf cnpj', 'cnpj fornecedor', 'cnpj cliente', 'cnpj participante', 'cnpj emitente', 'cnpj destinatario', 'cnpj tomador', 'cnpj prestador', 'inscricao', 'cnpj cpf participante']),
    ('participante', 'Fornecedor/Cliente', ['fornecedor', 'cliente', 'razao social', 'nome', 'participante', 'emitente', 'destinatario', 'remetente', 'tomador', 'prestador', 'nome fornecedor', 'nome cliente', 'razao', 'transportadora', 'transportador', 'parceiro', 'descricao parceiro', 'nome parceiro', 'descricao parceiro empresa', 'razao social parceiro', 'locatario', 'locador', 'inquilino', 'nome locatario']),
    ('uf', 'UF', ['uf', 'estado', 'uf origem', 'uf destino']),
    ('cfop', 'CFOP', ['cfop', 'cfo', 'c f o p', 'cod fiscal', 'codigo fiscal']),
    ('cst_icms', 'CST ICMS', ['cst', 'cst icms', 'sit trib', 'situacao tributaria', 'cst csosn', 'csosn', 'cod sit trib', 'st icms', 'cst do icms', 'sit tributaria', 'tributacao icms', 'cod tributacao', 'codigo tributacao', 'tributacao', 'cst icms entrada']),
    ('cst_ipi', 'CST IPI', ['cst ipi', 'sit trib ipi']),
    ('cst_pis', 'CST PIS', ['cst pis', 'sit trib pis', 'cst pis cofins']),
    ('cst_cofins', 'CST COFINS', ['cst cofins', 'sit trib cofins']),
    ('valor_contabil', 'Valor contábil/total', ['valor contabil', 'vl contabil', 'valor total', 'total nota', 'valor nota', 'vl total', 'valor documento', 'total', 'valor', 'vl doc', 'valor nf', 'total da nota', 'valor da nota', 'vlr contabil', 'vlr total', 'valor do documento', 'vl documento', 'valor servico', 'valor servicos', 'vl servico', 'valor dos servicos', 'valor bruto', 'valor total servico', 'valor total dos servicos', 'valor aluguel', 'valor locacao', 'aluguel']),
    ('base_icms', 'Base ICMS', ['base icms', 'bc icms', 'base calculo', 'base de calculo', 'vl bc icms', 'base', 'base calculo icms', 'vlr base icms', 'base de calculo icms', 'bc']),
    ('aliq_icms', 'Alíquota ICMS', ['aliquota', 'aliq', 'aliq icms', 'aliquota icms', 'aliq aliq']),
    ('valor_icms', 'Valor ICMS', ['icms', 'valor icms', 'vl icms', 'icms creditado', 'imposto creditado', 'imposto', 'vlr icms', 'icms debitado', 'imposto debitado', 'credito icms']),
    ('isentas', 'Isentas/Não trib.', ['isentas', 'isentas nao tributadas', 'isenta', 'isentas ou nao tributadas', 'isentas icms', 'isentas nao tributadas icms']),
    ('outras', 'Outras', ['outras', 'outros', 'outras icms']),
    ('base_st', 'Base ICMS ST', ['base st', 'bc st', 'base icms st', 'base subst', 'base substituicao', 'bc icms st', 'base calculo st', 'base retencao', 'base icms retencao', 'base st retido']),
    ('valor_st', 'Valor ICMS ST', ['icms st', 'valor st', 'vl st', 'st', 'icms subst', 'valor icms st', 'vl icms st', 'icms substituicao', 'substituicao tributaria', 'icms retencao', 'icms retido', 'valor icms retido', 'vlr icms st']),
    ('base_ipi', 'Base IPI', ['base ipi', 'bc ipi', 'base calculo ipi', 'base de calculo ipi', 'vl bc ipi']),
    ('valor_ipi', 'Valor IPI', ['ipi', 'valor ipi', 'vl ipi', 'vlr ipi', 'ipi creditado', 'ipi debitado']),
    ('base_pis', 'Base PIS/COFINS', ['base pis', 'bc pis', 'base pis cofins', 'base cofins', 'bc pis cofins', 'base credito pis']),
    ('valor_pis', 'Valor PIS', ['pis', 'valor pis', 'vl pis', 'vlr pis', 'credito pis']),
    ('valor_cofins', 'Valor COFINS', ['cofins', 'valor cofins', 'vl cofins', 'vlr cofins', 'credito cofins']),
    ('valor_iss', 'Valor ISS', ['iss', 'valor iss', 'vl iss', 'issqn', 'valor issqn']),
    ('aliq_iss', 'Alíquota ISS', ['aliq iss', 'aliquota iss', 'aliq issqn']),
    ('iss_retido', 'ISS retido', ['iss retido', 'issqn retido', 'retido', 'iss ret', 'valor iss retido', 'vl iss retido']),
    ('irrf', 'IRRF retido', ['irrf', 'ir', 'ir retido', 'irrf retido', 'valor ir', 'vl ir', 'ir fonte']),
    ('csll_ret', 'CSLL retida', ['csll', 'csll retida', 'csll retido', 'valor csll']),
    ('inss_ret', 'INSS retido', ['inss', 'inss retido', 'valor inss', 'ret inss', 'inss retencao']),
    ('pis_ret', 'PIS retido', ['pis retido', 'pis ret']),
    ('cofins_ret', 'COFINS retida', ['cofins retido', 'cofins retida', 'cofins ret']),
    ('cod_servico', 'Código do serviço', ['codigo servico', 'cod servico', 'item lista', 'item lc', 'item lc 116', 'servico lc', 'codigo do servico', 'item servico', 'ctribnac', 'codigo tributacao nacional', 'subitem', 'item da lista']),
    ('base_iss', 'Base ISS', ['base iss', 'bc iss', 'base issqn', 'base calculo iss', 'base de calculo iss']),
    ('valor_liquido', 'Valor líquido', ['valor liquido', 'vl liquido', 'liquido', 'valor liquido nfse', 'vlr liquido']),
    ('descricao', 'Descrição do bem/serviço', ['descricao', 'bem', 'objeto', 'imovel', 'equipamento', 'descricao bem', 'bem locado', 'historico', 'descricao locacao']),
    ('acumulador', 'Acumulador/Natureza', ['acumulador', 'acum', 'natureza', 'natureza operacao', 'nat op']),
    ('data_cancelamento', 'Data de cancelamento', ['dt cancelamento', 'data cancelamento', 'cancelamento', 'dt cancel', 'data do cancelamento', 'cancelada em']),
    ('situacao', 'Situação (normal/cancelada)', ['situacao', 'sit', 'status', 'situacao documento', 'situacao nota', 'cod sit', 'sit doc', 'situacao do documento', 'cancelada']),
]
CAMPO_ROTULO = {c[0]: c[1] for c in CAMPOS}
NUMERICOS = {'base_ipi', 'valor_liquido', 'base_iss', 'valor_contabil', 'base_icms', 'aliq_icms', 'valor_icms', 'isentas', 'outras', 'base_st', 'valor_st', 'valor_ipi',
             'base_pis', 'valor_pis', 'valor_cofins', 'valor_iss', 'aliq_iss', 'irrf', 'csll_ret', 'inss_ret', 'pis_ret', 'cofins_ret', 'csrf_ret'}
# palavras que, presentes no cabeçalho e ausentes no sinônimo, desqualificam o casamento ("valor icms st" não é "valor icms")
DISCRIMINA = {'st', 'subst', 'substituicao', 'ipi', 'pis', 'cofins', 'iss', 'issqn', 'retido', 'retida', 'ret', 'irrf', 'ir', 'csll', 'inss',
              'base', 'bc', 'aliq', 'aliquota', 'cst', 'csosn', 'data', 'dt', 'cnpj', 'cpf', 'chave', 'serie', 'cfop', 'icms', 'emissao',
              'entrada', 'saida', 'uf', 'credito', 'nome', 'razao', 'isentas', 'outras', 'acumulador', 'modelo', 'especie', 'trib', 'tributaria', 'tributacao', 'rps', 'liquido', 'bruto', 'complemento', 'diferenca', 'retencao', 'fcp', 'difal', 'remet', 'fundo', 'antecipacao', 'cod',
              'codigo', 'tipo', 'contato', 'lancamento', 'unico', 'sequencia', 'conta', 'contabil', 'descricao', 'destino', 'origem', 'entrega', 'cidade', 'produto', 'fantasia'}

ARQ_MAPAS = os.path.join(os.environ.get('REVISOR_DADOS') or os.path.join(os.path.dirname(__file__), '..'), 'config', 'mapeamentos.json')


def _pdf_livro_servicos(caminho):
    """'Livro de Registro de Serviços Prestados/Adquiridos Mod. 1' (sistema do cliente): uma linha por nota com
    data (dd/mm), número, série, valor, material, subempreitada, saldo, alíquota do ISS e observação na linha de baixo.
    Devolve uma tabela com cabeçalho próprio; linhas de 'TOTAL DO DIA' e o resumo ficam fora."""
    import fitz
    doc = fitz.open(caminho)
    if 'LIVRO DE REGISTRO DE SERVI' not in doc[0].get_text().upper():
        return None
    ano = ''
    m = re.search(r'\d{2}/\d{2}/(\d{4})', doc[0].get_text())
    if m:
        ano = m.group(1)
    saida = [['Data', 'Número', 'Série', 'Valor Contábil', 'Valor Saldo', 'Alíq. ISS', 'Observações']]
    for pg in doc:
        linhas = {}
        for w in pg.get_text('words'):
            linhas.setdefault(round(w[1]), []).append((w[0], w[4]))
        ys = sorted(linhas)
        for i, y in enumerate(ys):
            ws = sorted(linhas[y])
            if not ws or ws[0][0] > 40 or not re.fullmatch(r'\d{2}/\d{2}', ws[0][1]):
                continue
            def col(x0, x1):
                return ' '.join(t for x, t in ws if x0 <= x < x1)
            obs = col(455, 9999)
            for y2 in ys[i + 1:i + 3]:
                if 0 < y2 - y <= 4:
                    obs = (obs + ' ' + ' '.join(t for x, t in sorted(linhas[y2]) if x >= 455)).strip()
            dd, mm = ws[0][1].split('/')
            saida.append(['%s-%s-%s' % (ano, mm, dd) if ano else ws[0][1], col(60, 140), col(140, 185), num(col(185, 260)),
                          num(col(370, 430)), num(col(430, 455)), obs])
    return saida if len(saida) > 1 else None


def _pdf_tabela(caminho):
    """Relatório em PDF (texto, não escaneado) -> linhas de tabela. Agrupa as palavras por linha (posição vertical),
    acha a linha de cabeçalho pelos mesmos sinônimos das planilhas e encaixa cada palavra na coluna pela posição horizontal."""
    import fitz
    doc = fitz.open(caminho)
    todas = []  # (pagina, y, [(x0, x1, texto)])
    for np_, pg in enumerate(doc):
        ws = sorted(pg.get_text('words'), key=lambda w: (round(w[3], 0), w[0]))
        linhas = []
        for w in ws:
            yc = (w[1] + w[3]) / 2
            if linhas and abs(linhas[-1][0] - yc) <= 2.5:
                linhas[-1][1].append((w[0], w[2], w[4]))
            else:
                linhas.append([yc, [(w[0], w[2], w[4])]])
        for yc, pal in linhas:
            todas.append((np_, yc, sorted(pal)))
    if not todas:
        return []
    def celulas(pal):
        out = []
        for x0, x1, t in pal:
            if out and x0 - out[-1][1] <= 4.5:
                out[-1] = (out[-1][0], x1, out[-1][2] + ' ' + t)
            else:
                out.append((x0, x1, t))
        return out
    melhor, idx_cab = 0, None
    for i, (_, _, pal) in enumerate(todas[:200]):
        cels = celulas(pal)
        if len(cels) < 3:
            continue
        pts = sum(1 for c in cels if _casar(norm(c[2])))
        if pts > melhor:
            melhor, idx_cab = pts, i
    if idx_cab is None or melhor < 2:
        return [[' '.join(t for _, _, t in pal)] for _, _, pal in todas]
    cab = celulas(todas[idx_cab][2])
    inicios = [c[0] for c in cab]
    fins = inicios[1:] + [1e9]
    txt_cab = [c[2] for c in cab]
    saida = [txt_cab]
    for _, _, pal in todas[idx_cab + 1:]:
        cels = [''] * len(cab)
        for x0, x1, t in pal:
            xc = (x0 + x1) / 2
            j = next((k for k in range(len(cab)) if inicios[k] - 6 <= xc < fins[k] - 6), None)
            if j is None:
                j = min(range(len(cab)), key=lambda k: abs(inicios[k] - x0))
            cels[j] = (cels[j] + ' ' + t).strip()
        if [c.strip() for c in cels] == [c.strip() for c in txt_cab]:
            continue  # cabeçalho repetido na página seguinte
        saida.append(cels)
    return saida


def ler_arquivo(caminho):
    """Devolve [{'aba': nome, 'linhas': [[valores]]}]."""
    ext = os.path.splitext(caminho)[1].lower()
    if ext == '.pdf':
        livro = _pdf_livro_servicos(caminho)
        if livro:
            return [{'aba': 'Livro de serviços', 'linhas': livro}]
        return [{'aba': 'PDF', 'linhas': _pdf_tabela(caminho)}]
    if ext in ('.xlsx', '.xlsm'):
        import openpyxl
        wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
        abas = []
        for ws in wb.worksheets:
            linhas = [list(r) for r in ws.iter_rows(values_only=True)]
            abas.append({'aba': ws.title, 'linhas': linhas})
        wb.close()
        return abas
    if ext == '.xls':
        import xlrd
        wb = xlrd.open_workbook(caminho)
        abas = []
        for sh in wb.sheets():
            linhas = []
            for r in range(sh.nrows):
                row = []
                for c in range(sh.ncols):
                    cel = sh.cell(r, c)
                    if cel.ctype == xlrd.XL_CELL_DATE:
                        try:
                            row.append(xlrd.xldate.xldate_as_datetime(cel.value, wb.datemode))
                        except Exception:
                            row.append(cel.value)
                    else:
                        row.append(cel.value)
                linhas.append(row)
            abas.append({'aba': sh.name, 'linhas': linhas})
        return abas
    # csv / txt
    raw = open(caminho, 'rb').read()
    for enc in ('utf-8-sig', 'cp1252', 'latin-1'):
        try:
            txt = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    try:
        dialect = csv.Sniffer().sniff(txt[:5000], delimiters=';,\t|')
        delim = dialect.delimiter
    except csv.Error:
        delim = ';'
    linhas = [r for r in csv.reader(io.StringIO(txt), delimiter=delim)]
    return [{'aba': os.path.basename(caminho), 'linhas': linhas}]


STOP = {'do', 'da', 'de', 'dos', 'das', 'e', 'a', 'o', 'em', 'p', 'para'}


def _casar(cab):
    """Para um cabeçalho normalizado, devolve [(pontuação, campo)]."""
    cab = ' '.join(t for t in cab.split() if t not in STOP)
    toks = set(cab.split())
    out = []
    for campo, _, sins in CAMPOS:
        best = 0
        for i, s in enumerate(sins):
            st = s.split()
            if cab == s:
                best = max(best, 100 + len(st) + (len(sins) - i) / 100.0)  # empate: vale o sinônimo listado antes
                continue
            if not set(st) <= toks:
                continue
            extra = toks - set(st)
            if extra & DISCRIMINA:
                continue
            best = max(best, 10 * len(st) - len(extra))
        if best > 0:
            out.append((best, campo))
    return out


def detectar(linhas):
    """Escolhe a linha de cabeçalho (entre as 40 primeiras) e casa colunas com campos."""
    melhor = (0, None, {})
    for i, row in enumerate(linhas[:40]):
        cabs = [norm(c) for c in row]
        if sum(1 for c in cabs if c) < 2:
            continue
        cand = []
        for j, c in enumerate(cabs):
            if not c or len(c) > 60:
                continue
            for pts, campo in _casar(c):
                cand.append((pts, j, campo))
        cand.sort(key=lambda c: (-c[0], c[1]))  # maior pontuação; empate: coluna mais à esquerda
        usados_col, mapa = set(), {}
        for pts, j, campo in cand:
            if j in usados_col or campo in mapa:
                continue
            mapa[campo] = j
            usados_col.add(j)
        score = len(mapa) + (3 if 'numero' in mapa or 'chave' in mapa else 0) + (1 if 'cfop' in mapa else 0)
        if score > melhor[0]:
            melhor = (score, i, mapa)
    _, linha_cab, mapa = melhor
    return linha_cab, mapa


def assinatura(linhas, linha_cab):
    if linha_cab is None:
        return ''
    cabs = '|'.join(norm(c) for c in linhas[linha_cab])
    return hashlib.md5(cabs.encode()).hexdigest()[:12]


def mapas_salvos():
    try:
        return json.load(open(ARQ_MAPAS))
    except Exception:
        return {}


def salvar_mapa(assin, linha_cab, mapa):
    m = mapas_salvos()
    m[assin] = {'linha_cab': linha_cab, 'mapa': mapa}
    os.makedirs(os.path.dirname(ARQ_MAPAS), exist_ok=True)
    json.dump(m, open(ARQ_MAPAS, 'w'), indent=1, ensure_ascii=False)


def _txt(v):
    if v is None:
        return ''
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def analisar(caminho, ajuste=None, servicos=False, locacoes=False):
    """Descreve cada aba: cabeçalho escolhido, mapa de colunas e amostra. `ajuste` = {aba: {linha_cab, mapa}}."""
    saida = []
    salvos = mapas_salvos()
    for aba in ler_arquivo(caminho):
        linhas = aba['linhas']
        linha_cab, mapa = detectar(linhas)
        origem = 'automatico'
        assin = assinatura(linhas, linha_cab)
        if assin in salvos:
            linha_cab, mapa = salvos[assin]['linha_cab'], salvos[assin]['mapa']
            origem = 'salvo'
        if ajuste and aba['aba'] in ajuste:
            linha_cab = ajuste[aba['aba']].get('linha_cab', linha_cab)
            mapa = {k: int(v) for k, v in ajuste[aba['aba']].get('mapa', {}).items() if v is not None and v != ''}
            origem = 'ajustado'
        if servicos and origem != 'ajustado' and linha_cab is not None:
            cab0 = [norm(_txt(c)) for c in linhas[linha_cab]]
            # "PIS - Débito", alíquota ou base não são retenção (relatório do Portal Nacional traz essas colunas)
            for campo in ('valor_pis', 'valor_cofins', 'pis_ret', 'cofins_ret', 'csll_ret'):
                j = mapa.get(campo)
                if j is not None and j < len(cab0) and any(w in cab0[j] for w in ('debito', 'aliquota', 'base', 'bc ', 'credito')):
                    mapa.pop(campo)
            # no relatório de serviços, "PIS", "COFINS", "CSLL" e "INSS" são os valores retidos
            for de, para in (('valor_pis', 'pis_ret'), ('valor_cofins', 'cofins_ret')):
                if de in mapa and para not in mapa:
                    mapa[para] = mapa.pop(de)
            # Portal Nacional: PIS + COFINS + CSLL retidos numa coluna só ("Contrib. Sociais Ret.")
            j = next((i for i, t in enumerate(cab0) if 'contrib' in t and 'socia' in t and 'ret' in t and not t.startswith('descr')), None)
            if j is not None and not any(k in mapa for k in ('pis_ret', 'cofins_ret', 'csll_ret')):
                mapa['csrf_ret'] = j
        filtro_loc = None
        if locacoes:
            # do relatório de locação só interessam o número da nota, o cliente e o valor da locação
            mapa = {k: v for k, v in mapa.items() if k in CAMPOS_LOCACAO}
            if linha_cab is not None:
                filtro_loc = _coluna_locacao(linhas[linha_cab + 1:])
        cab = [_txt(c) for c in linhas[linha_cab]] if linha_cab is not None and linha_cab < len(linhas) else []
        amostra = [[_txt(c) for c in r] for r in linhas[(linha_cab or 0) + 1:(linha_cab or 0) + 6]]
        n_dados = sum(1 for r in linhas[(linha_cab or 0) + 1:] if any(_txt(c) for c in r)) if linha_cab is not None else 0
        prova, alertas = _conferir(linhas, linha_cab, mapa, cab) if linha_cab is not None else ({}, [])
        if locacoes and linha_cab is not None:
            if filtro_loc:
                fora = filtro_loc['fora']
                alertas.append({'sev': 'info', 'campo': 'valor_contabil', 'txt': 'Só as linhas de locação entram (coluna "%s"): %d linha(s) de outro tipo ficaram fora (%s).'
                                % (cab[filtro_loc['col']] if filtro_loc['col'] < len(cab) else '', sum(fora.values()), ', '.join('%s: %d' % kv for kv in sorted(fora.items())))})
            if 'participante' not in mapa and 'cnpj' not in mapa:
                alertas.append({'sev': 'alerta', 'campo': 'participante', 'txt': 'O relatório não tem coluna de cliente (locatário/parceiro). Exporte do sistema com o nome ou CNPJ do cliente para ele aparecer.'})
        saida.append({'arquivo': os.path.basename(caminho), 'aba': aba['aba'], 'linha_cab': linha_cab, 'mapa': mapa, 'prova': prova, 'alertas': alertas,
                      'cabecalho': cab, 'amostra': amostra, 'linhas_dados': n_dados, 'origem': origem,
                      'assinatura': assin, 'filtro_loc': filtro_loc, 'utilizavel': linha_cab is not None and (('valor_contabil' in mapa) if locacoes else ('numero' in mapa or 'chave' in mapa))})
    return saida


CAMPOS_LOCACAO = ('numero', 'valor_contabil', 'participante', 'cnpj')


def _coluna_locacao(dados):
    """Relatório que mistura locação com outras vendas (ex.: TOP 'VENDA LOCAÇÃO' e 'VENDA NFS-E'): acha a coluna
    que diz o tipo da linha. Sem essa coluna, todas as linhas valem como locação."""
    dados = [r for r in dados if any(_txt(c) for c in r)]
    if not dados:
        return None
    ncol = max(len(r) for r in dados)
    melhor = None
    for j in range(ncol):
        vals = [norm(_txt(r[j])) if j < len(r) else '' for r in dados]
        com = sum(1 for v in vals if 'locac' in v or 'aluguel' in v)
        sem = [v for v in vals if v and not ('locac' in v or 'aluguel' in v)]
        if com and sem and len(set(vals)) <= 12 and (melhor is None or com > melhor[1]):
            melhor = (j, com, sem)
    if not melhor:
        return None
    j, _, sem = melhor
    fora = {}
    for r in dados:
        v = _txt(r[j]) if j < len(r) else ''
        if v and not ('locac' in norm(v) or 'aluguel' in norm(v)):
            fora[v] = fora.get(v, 0) + 1
    return {'col': j, 'fora': fora}


def eh_locacao(row, filtro):
    if not filtro:
        return True
    j = filtro['col']
    v = norm(_txt(row[j])) if j < len(row) else ''
    return 'locac' in v or 'aluguel' in v


def _conferir(linhas, linha_cab, mapa, cab):
    """Linha de prova (a primeira linha com valor) e alertas de coluna suspeita: valor zerado em todas as linhas,
    número da nota sempre vazio/zero, CFOP inválido, coluna de valor que não é número."""
    dados = [r for r in linhas[linha_cab + 1:] if any(_txt(c) for c in r)][:3000]
    def val(r, j):
        return r[j] if j is not None and j < len(r) else None
    prova, alertas = {}, []
    numericos = [j for c, j in mapa.items() if c in NUMERICOS]
    alvo = max(dados[:800], key=lambda r: sum(1 for j in numericos if num(val(r, j)) != 0) + (1 if _txt(val(r, mapa.get('chave'))) else 0)) if dados else []
    for campo, j in mapa.items():
        bruto = _txt(val(alvo, j))
        prova[campo] = {'coluna': cab[j] if j < len(cab) else '', 'valor': bruto[:40]}
    if not dados:
        return prova, alertas
    n = len(dados)
    for campo in ('valor_contabil', 'base_icms', 'valor_icms', 'valor_ipi', 'valor_iss'):
        j = mapa.get(campo)
        if j is None:
            continue
        nz = sum(1 for r in dados if num(val(r, j)) != 0)
        txt = sum(1 for r in dados if _txt(val(r, j)) and num(val(r, j)) == 0 and not re.match(r'^[\d.,\-\s]+$', _txt(val(r, j))))
        rot = CAMPO_ROTULO[campo]
        if nz == 0 and n >= 5:
            alertas.append({'sev': 'erro', 'campo': campo, 'txt': '%s: a coluna "%s" está zerada em todas as %d linhas. Confira se é a coluna certa.' % (rot, cab[j], n)})
        elif txt > n * 0.3:
            alertas.append({'sev': 'erro', 'campo': campo, 'txt': '%s: a coluna "%s" tem texto em vez de número.' % (rot, cab[j])})
    j = mapa.get('numero')
    if j is not None and 'chave' not in mapa:
        vazios = sum(1 for r in dados if not digitos(val(r, j)).strip('0'))
        if vazios > n * 0.5:
            alertas.append({'sev': 'erro', 'campo': 'numero', 'txt': 'Número da nota: a coluna "%s" está vazia ou zero em %d de %d linhas.' % (cab[j], vazios, n)})
    j = mapa.get('cfop')
    if j is not None:
        ruins = sum(1 for r in dados if not re.match(r'^[1-7]\d{3}$', digitos(val(r, j))[:4]))
        if ruins > n * 0.2:
            alertas.append({'sev': 'alerta', 'campo': 'cfop', 'txt': 'CFOP: %d de %d linhas sem um CFOP válido na coluna "%s".' % (ruins, n, cab[j])})
    for campo in ('valor_contabil',):
        if campo not in mapa:
            alertas.append({'sev': 'alerta', 'campo': campo, 'txt': 'Não achei a coluna de valor contábil/total: as conferências de valor não vão rodar.'})
    return prova, alertas


def registros(caminho, desc_abas):
    """Lê os lançamentos usando o mapa de cada aba. Linhas de continuação (só CFOP/valores, sem número)
    herdam a identificação da linha anterior, como nos livros fiscais exportados."""
    regs = []
    por_aba = {d['aba']: d for d in desc_abas}
    for aba in ler_arquivo(caminho):
        d = por_aba.get(aba['aba'])
        if not d or not d['utilizavel']:
            continue
        mapa, ini = d['mapa'], d['linha_cab'] + 1
        anterior = None
        for i, row in enumerate(aba['linhas'][ini:], start=ini + 1):
            def g(campo):
                j = mapa.get(campo)
                return row[j] if j is not None and j < len(row) else None
            if not any(_txt(c) for c in row):
                continue
            if d.get('filtro_loc') and not eh_locacao(row, d['filtro_loc']):
                continue
            primeiro = norm(next((c for c in row if _txt(c)), ''))
            if primeiro.startswith('total') or primeiro.startswith('subtotal'):
                continue
            r = {'arquivo': d['arquivo'], 'aba': aba['aba'], 'linha': i}
            for campo in mapa:
                v = g(campo)
                if campo in NUMERICOS:
                    r[campo] = num(v)
                elif campo in ('data_emissao', 'data_lanc', 'data_cancelamento'):
                    r[campo] = data_iso(v)
                elif campo == 'cnpj':
                    r[campo] = digitos(v)
                elif campo == 'chave':
                    r[campo] = digitos(v)
                elif campo in ('cfop',):
                    r[campo] = digitos(v)[:4]
                elif campo.startswith('cst'):
                    r[campo] = digitos(v)
                elif campo == 'numero':
                    r[campo] = digitos(v)
                else:
                    r[campo] = _txt(v)
            tem_id = bool(r.get('numero') or r.get('chave')) or (d.get('locacoes') and r.get('valor_contabil'))
            tem_valor = any(r.get(c) for c in NUMERICOS) or r.get('cfop')
            if not tem_id and tem_valor and anterior:
                for c in ('numero', 'chave', 'serie', 'modelo', 'cnpj', 'participante', 'data_emissao', 'data_lanc', 'uf', 'movimento'):
                    if not r.get(c) and anterior.get(c):
                        r[c] = anterior[c]
                r['continuacao'] = True
                tem_id = True
            if not tem_id:
                continue
            if r.get('numero') and not r.get('cfop') and not any(r.get(c) for c in NUMERICOS) and not r.get('chave'):
                anterior = r  # linha de cabeçalho da nota, valores vêm nas próximas
                continue
            regs.append(r)
            anterior = r
    return regs
