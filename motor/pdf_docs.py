# -*- coding: utf-8 -*-
"""Documentos Diversos (PDF geral de notas): lê o texto de cada página e identifica os documentos pela chave
de acesso (DANFE, DACTE) ou, nas NFS-e municipais, pelo número da nota e CNPJs da página."""
import os, re
from .util import chave_valida, info_chave, digitos, num

RE_CHAVE = re.compile(r'(?<!\d)((?:\d[\s.]?){43}\d)(?!\d)')
RE_CNPJ = re.compile(r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}')
RE_NFSE_NUM = re.compile(r'(?:N[uú]mero\s+da\s+NFS-?e|NFS-?e\s*N[ºo°.]*|N[uú]mero\s+da\s+Nota)\s*[:\-]?\s*(\d{1,15})', re.I)
RE_VALOR = re.compile(r'(?:Valor\s+(?:Total|L[ií]quido|do\s+Servi[cç]o|dos\s+Servi[cç]os|da\s+Nota))[^\d]{0,40}?([\d.]+,\d{2})', re.I)


def ler(caminho):
    nome = os.path.basename(caminho)
    try:
        import fitz
        doc = fitz.open(caminho)
    except Exception as e:
        return {'arquivo': nome, 'erro': str(e), 'paginas': 0, 'docs': [], 'sem_texto': []}
    docs, vistos, sem_texto = [], set(), []
    for i, pg in enumerate(doc, start=1):
        txt = pg.get_text() or ''
        if len(txt.strip()) < 20:
            sem_texto.append(i)
            continue
        achou = False
        for m in RE_CHAVE.finditer(txt):
            ch = digitos(m.group(1))
            if len(ch) == 44 and chave_valida(ch) and ch not in vistos:
                vistos.add(ch)
                inf = info_chave(ch)
                docs.append({'tipo': {'55': 'nfe', '65': 'nfce', '57': 'cte', '67': 'cte'}.get(inf['modelo'], 'outro'),
                             'chave': ch, 'numero': inf['numero'], 'serie': inf['serie'], 'cnpj_emit': inf['cnpj'],
                             'modelo': inf['modelo'], 'aamm': inf['aamm'], 'arquivo': nome, 'pagina': i})
                achou = True
        if not achou:
            mnum = RE_NFSE_NUM.search(txt)
            if mnum and re.search(r'NFS-?e|Nota\s+Fiscal\s+de\s+Servi', txt, re.I):
                n = mnum.group(1).lstrip('0') or '0'
                cnpjs = [digitos(c) for c in RE_CNPJ.findall(txt)]
                mv = RE_VALOR.search(txt)
                k = ('nfse', n, tuple(cnpjs[:2]))
                if k not in vistos:
                    vistos.add(k)
                    docs.append({'tipo': 'nfse', 'chave': '', 'numero': n, 'cnpjs': cnpjs, 'valor': num(mv.group(1)) if mv else 0.0,
                                 'arquivo': nome, 'pagina': i})
                achou = True
        if not achou:
            cnpjs = [digitos(c) for c in RE_CNPJ.findall(txt)][:3]
            if cnpjs:
                docs.append({'tipo': 'nao_identificado', 'arquivo': nome, 'pagina': i, 'cnpjs': cnpjs})
            else:
                sem_texto.append(i)
    return {'arquivo': nome, 'paginas': len(doc), 'docs': docs, 'sem_texto': sem_texto}
