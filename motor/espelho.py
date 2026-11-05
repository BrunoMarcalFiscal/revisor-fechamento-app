# -*- coding: utf-8 -*-
"""PDF de validação da nota, gerado do XML (NF-e, CT-e, NFS-e), com os apontamentos do Revisor. Serve para conferência: não é o DANFE."""
import html
import io
import fitz
from . import xml_docs
from .util import brl

CSS = 'body{font-family:sans-serif;font-size:8.5pt;color:#212E50} h1{font-size:13pt;margin:0} h2{font-size:9.5pt;margin:9px 0 3px;color:#212E50} p{margin:1px 0} ' \
      'table{border-collapse:collapse;width:100%} th{background-color:#212E50;color:#FFFFFF;font-size:7.5pt;padding:2px 3px;text-align:left} ' \
      'td{border:0.5px solid #C9D3E3;font-size:7.5pt;padding:2px 3px} .av{color:#8A5A00;font-size:8pt} .erro{color:#9B1C1C;font-weight:bold} .conf{color:#8A5A00;font-weight:bold}'
SIT = {'100': 'Autorizada', '101': 'Cancelada', '135': 'Cancelada', '151': 'Cancelada', '155': 'Cancelada', '110': 'Denegada', '301': 'Denegada', '302': 'Denegada', '303': 'Denegada'}


def _e(v):
    return html.escape(str(v if v is not None else ''))


def _d(v):
    v = str(v or '')
    return '%s/%s/%s' % (v[8:10], v[5:7], v[:4]) if len(v) >= 10 else v


def _doc(c):
    c = ''.join(x for x in str(c or '') if x.isdigit())
    if len(c) == 14:
        return '%s.%s.%s/%s-%s' % (c[:2], c[2:5], c[5:8], c[8:12], c[12:])
    return c


def _q(v):
    return ('%.4f' % (v or 0)).rstrip('0').rstrip('.').replace('.', ',')


def _tabela(cab, linhas):
    return '<table><tr>%s</tr>%s</table>' % (''.join('<th>%s</th>' % _e(c) for c in cab),
                                             ''.join('<tr>%s</tr>' % ''.join('<td>%s</td>' % _e(x) for x in l) for l in linhas))


def _corpo_nfe(d):
    e, de, t = d['emit'], d['dest'], d['totais']
    it = [[i.get('nItem'), i.get('cProd'), i.get('xProd'), i.get('NCM'), i.get('CFOP'), (i['icms'].get('orig', '') + i['icms'].get('CST', '')), i.get('uCom'), _q(i.get('qCom')), brl(i.get('vUnCom')), brl(i.get('vProd')),
           brl(i['icms'].get('vBC')), _q(i['icms'].get('pICMS')), brl(i['icms'].get('vICMS')), brl(i['icms'].get('vICMSST')), brl(i['ipi'].get('vIPI')), brl(i['pis'].get('vPIS')), brl(i['cofins'].get('vCOFINS'))] for i in d['itens']]
    h = '<h2>Emitente e destinatário</h2>' + _tabela(['', 'Nome', 'CNPJ/CPF', 'UF', 'IE'], [['Emitente', e.get('nome'), _doc(e.get('cnpj')), e.get('uf'), e.get('ie')], ['Destinatário', de.get('nome'), _doc(de.get('cnpj')), de.get('uf'), de.get('ie')]])
    h += '<h2>Itens</h2>' + _tabela(['#', 'Código', 'Descrição', 'NCM', 'CFOP', 'CST', 'Un', 'Qtd', 'V. unit.', 'V. prod.', 'BC ICMS', '%', 'ICMS', 'ST', 'IPI', 'PIS', 'COFINS'], it)
    h += '<h2>Totais</h2>' + _tabela(['BC ICMS', 'ICMS', 'BC ST', 'ST', 'Produtos', 'Frete', 'Desconto', 'IPI', 'PIS', 'COFINS', 'Total da nota'],
                                     [[brl(t.get(k)) for k in ('vBC', 'vICMS', 'vBCST', 'vST', 'vProd', 'vFrete', 'vDesc', 'vIPI', 'vPIS', 'vCOFINS', 'vNF')]])
    if d.get('refs'):
        h += '<h2>Notas referenciadas</h2><p>%s</p>' % _e(', '.join(x for x in d['refs'] if x))
    if d.get('infAdic'):
        h += '<h2>Informações complementares</h2><p>%s</p>' % _e(d['infAdic'])
    return h, '%s nº %s, série %s' % ('NF-e' if d.get('modelo') != '65' else 'NFC-e', d.get('numero'), d.get('serie')), t.get('vNF')


def _corpo_cte(d):
    ic = d.get('icms') or {}
    h = '<h2>Transporte</h2>' + _tabela(['Emitente', 'CNPJ', 'CFOP', 'Trajeto', 'Tomador', 'Remetente', 'Destinatário'],
                                        [[d['emit'].get('nome'), _doc(d['emit'].get('cnpj')), d.get('CFOP'), '%s → %s' % (d.get('UFIni'), d.get('UFFim')), _doc(d.get('tomador_cnpj')), _doc(d['rem'].get('cnpj')), _doc(d['dest'].get('cnpj'))]])
    h += '<h2>Valores</h2>' + _tabela(['Prestação', 'A receber', 'CST', 'BC ICMS', '%', 'ICMS', 'Carga'], [[brl(d.get('vTPrest')), brl(d.get('vRec')), ic.get('CST'), brl(ic.get('vBC')), _q(ic.get('pICMS')), brl(ic.get('vICMS')), brl(d.get('vCarga'))]])
    if d.get('nfes'):
        h += '<h2>NF-e transportadas</h2><p>%s</p>' % _e(', '.join(x for x in d['nfes'] if x))
    return h, 'CT-e nº %s, série %s' % (d.get('numero'), d.get('serie')), d.get('vTPrest')


def _corpo_nfse(d):
    pr, to, r = d.get('prestador') or {}, d.get('tomador') or {}, d.get('ret') or {}
    h = '<h2>Prestador e tomador</h2>' + _tabela(['', 'Nome', 'CNPJ/CPF'], [['Prestador', pr.get('nome'), _doc(pr.get('cnpj'))], ['Tomador', to.get('nome'), _doc(to.get('cnpj'))]])
    h += '<h2>Serviço</h2><p>%s %s</p>' % (_e(d.get('cTribNac') or ''), _e(d.get('xDescServ') or ''))
    h += '<h2>Valores</h2>' + _tabela(['Serviço', 'Líquido', 'ISS', '%', 'ISS retido', 'IRRF', 'CSLL', 'PIS', 'COFINS', 'INSS'],
                                      [[brl(d.get('vServ')), brl(d.get('vLiq')), brl(d.get('vISS')), _q(d.get('pAliq')), 'Sim' if d.get('issRetido') else 'Não'] + [brl(r.get(k)) for k in ('irrf', 'csll', 'pis', 'cofins', 'inss')]])
    return h, 'NFS-e nº %s' % d.get('numero'), d.get('vServ')


def gerar(caminho, nota):
    """Bytes do PDF da nota; None quando o arquivo não é um documento que dá para ler."""
    docs = xml_docs.ler(caminho)
    d = next((x for x in docs if x.get('tipo') in ('nfe', 'cte', 'nfse')), None)
    if not d:
        return None
    corpo, titulo, total = {'nfe': _corpo_nfe, 'cte': _corpo_cte, 'nfse': _corpo_nfse}[d['tipo']](d)
    cancelada = nota.get('cancelada_efetiva') or any(x.get('tipo') == 'evento' and x.get('cancelamento') for x in docs)
    sit = 'Cancelada' if cancelada else SIT.get(str(d.get('protocolo') or ''), '')
    cab = '<h1>%s</h1><p class="av">PDF de validação gerado a partir do XML. Não é o DANFE e não substitui o documento fiscal.</p>' % _e(titulo)
    cab += '<p>Emissão %s · Total %s%s%s</p>' % (_d(d.get('emissao')), brl(total), (' · Situação: %s' % sit) if sit else '', (' · Natureza: %s' % _e(d['natOp'])) if d.get('natOp') else '')
    if d.get('chave'):
        cab += '<p>Chave: %s</p>' % _e(d['chave'])
    ap = [a for a in nota.get('aponts', []) if a.get('sev') in ('erro', 'alerta')]
    if ap:
        lin = ''.join('<tr><td class="%s">%s</td><td>%s</td><td>%s</td><td>%s</td><td>%s</td></tr>' % ('erro' if a['sev'] == 'erro' else 'conf', 'Erro' if a['sev'] == 'erro' else 'Conferir',
                      _e(a.get('titulo')) + ': ' + _e(a.get('detalhe')), _e(a.get('xml')), _e(a.get('sis')), _e(a.get('correcao'))) for a in ap)
        corpo = '<h2>Apontamentos do Revisor</h2><table><tr><th>Situação</th><th>O que está errado</th><th>Na nota</th><th>No sistema</th><th>Como corrigir</th></tr>%s</table>' % lin + corpo
    story = fitz.Story('<body>%s%s</body>' % (cab, corpo), user_css=CSS)
    buf = io.BytesIO()
    w = fitz.DocumentWriter(buf)
    mb = fitz.paper_rect('a4-l')
    mais = 1
    while mais:
        dev = w.begin_page(mb)
        mais, _ = story.place(mb + (24, 24, -24, -24))
        story.draw(dev)
        w.end_page()
    w.close()
    return buf.getvalue()
