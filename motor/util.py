# -*- coding: utf-8 -*-
"""Funções utilitárias do Revisor do Fechamento."""
import re, unicodedata, datetime

TOL = 0.05  # tolerância de centavos para comparar valores


def num(v):
    """Converte '1.234,56', '1234.56', 1234.56 ou None em float."""
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace('R$', '').replace(' ', '')
    if not s or s in ('-', '--'):
        return 0.0
    neg = s.startswith('(') and s.endswith(')')
    s = s.strip('()')
    if ',' in s and '.' in s:
        if s.rfind(',') > s.rfind('.'):
            s = s.replace('.', '').replace(',', '.')
        else:
            s = s.replace(',', '')
    elif ',' in s:
        s = s.replace(',', '.')
    s = s.replace('%', '')
    try:
        f = float(s)
    except ValueError:
        return 0.0
    return -f if neg else f


def digitos(v):
    if v is None:
        return ''
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return re.sub(r'\D', '', str(v))


def numero_doc(v):
    """Número da nota sem zeros à esquerda (mantém '0' se for só zero)."""
    d = digitos(v)
    return d.lstrip('0') or ('0' if d else '')


def norm(s):
    """Texto normalizado para comparar cabeçalhos: minúsculo, sem acento, só letras/números."""
    s = '' if s is None else str(s)
    s = s.replace('º', ' n ').replace('°', ' n ').replace('%', ' aliq ')
    s = unicodedata.normalize('NFKD', s).encode('ascii', 'ignore').decode()
    s = re.sub(r'[^a-z0-9]+', ' ', s.lower())
    return s.strip()


def data_iso(v):
    """datetime, 'dd/mm/aaaa', 'aaaa-mm-dd...' ou número serial do Excel -> 'aaaa-mm-dd'."""
    if v is None or v == '':
        return ''
    if isinstance(v, datetime.datetime):
        return v.date().isoformat()
    if isinstance(v, datetime.date):
        return v.isoformat()
    if isinstance(v, (int, float)) and 20000 < float(v) < 80000:
        return (datetime.date(1899, 12, 30) + datetime.timedelta(days=int(v))).isoformat()
    s = str(v).strip()
    m = re.match(r'^(\d{4})-(\d{2})-(\d{2})', s)
    if m:
        return '%s-%s-%s' % m.groups()
    m = re.match(r'^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})', s)
    if m:
        d, mth, y = m.groups()
        if len(y) == 2:
            y = '20' + y
        return '%s-%02d-%02d' % (y, int(mth), int(d))
    return ''


def brl(v):
    s = '{:,.2f}'.format(abs(v or 0)).replace(',', 'X').replace('.', ',').replace('X', '.')
    return ('-' if (v or 0) < 0 else '') + 'R$ ' + s


def igual(a, b, tol=TOL):
    return abs((a or 0) - (b or 0)) <= tol


def raiz(cnpj):
    d = digitos(cnpj)
    return d[:8] if len(d) == 14 else d


def chave_valida(ch):
    """Dígito verificador (módulo 11) da chave de 44 posições de NF-e/CT-e/NFC-e/MDF-e."""
    if not re.fullmatch(r'\d{44}', ch or ''):
        return False
    pesos, soma, p = [], 0, 2
    for c in reversed(ch[:43]):
        soma += int(c) * p
        p = 2 if p == 9 else p + 1
    dv = 11 - soma % 11
    dv = 0 if dv >= 10 else dv
    return dv == int(ch[43])


def info_chave(ch):
    """Decompõe a chave de 44 posições."""
    if len(ch or '') != 44:
        return {}
    return {'uf_cod': ch[0:2], 'aamm': ch[2:6], 'cnpj': ch[6:20], 'modelo': ch[20:22],
            'serie': str(int(ch[22:25])), 'numero': str(int(ch[25:34]))}
