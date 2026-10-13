# -*- coding: utf-8 -*-
"""Leitor de XML de documentos fiscais: NF-e/NFC-e (procNFe), CT-e (procCTe), NFS-e (padrão nacional e ABRASF) e eventos.
Adaptado do Revisor Tributário. ler() devolve sempre uma LISTA (lote ABRASF traz várias NFS-e num arquivo)."""
import os, re
import xml.etree.ElementTree as ET
from .util import num

def _strip_ns(root):
    for el in root.iter():
        if isinstance(el.tag, str) and '}' in el.tag:
            el.tag = el.tag.split('}', 1)[1]
    return root

def _t(el, path, default=''):
    if el is None:
        return default
    x = el.find(path)
    return (x.text or '').strip() if x is not None and x.text is not None else default

def _n(el, path):
    return num(_t(el, path, '0'))

def _cnpj(el, tag='CNPJ'):
    v = _t(el, tag) or _t(el, 'CPF')
    return re.sub(r'\D', '', v)

def ler(caminho, conteudo=None):
    nome = os.path.basename(caminho)
    try:
        if conteudo is not None:
            root = _strip_ns(ET.fromstring(conteudo))
        else:
            root = _strip_ns(ET.parse(caminho).getroot())
    except ET.ParseError as e:
        return [{'tipo': 'erro', 'arquivo': nome, 'erro': str(e)}]
    tag = root.tag
    if root.find('.//infCte') is not None or tag == 'infCte':
        return [_cte(root, caminho)]  # antes da NF-e: o CT-e cita as NF-e em infDoc/infNFe
    if root.find('.//infNFe/ide') is not None or tag == 'infNFe':
        return [_nfe(root, caminho)]
    if root.find('.//infInut') is not None:
        inf = root.find('.//infInut')
        ret = root.find('.//retInutNFe/infInut')
        return [{'tipo': 'inut', 'arquivo': nome, 'cnpj': _t(inf, 'CNPJ'), 'modelo': _t(inf, 'mod'), 'serie': _t(inf, 'serie'),
                 'ini': int(_t(inf, 'nNFIni') or 0), 'fim': int(_t(inf, 'nNFFin') or 0), 'ano': _t(inf, 'ano'), 'just': _t(inf, 'xJust'),
                 'cStat': _t(ret, 'cStat') if ret is not None else '', 'dh': (_t(ret, 'dhRecbto') if ret is not None else '')[:10]}]
    if root.find('.//infEvento') is not None:
        return [_evento(root, caminho)]
    if root.find('.//infNFSe') is not None or root.find('.//infDPS') is not None:
        return [_nfse_nacional(root, caminho)]
    comps = root.findall('.//CompNfse')
    if comps:
        return [_nfse_abrasf(c, caminho) for c in comps]
    if root.find('.//InfNfse') is not None or 'Nfse' in tag:
        return [_nfse_abrasf(root, caminho)]
    return [{'tipo': 'desconhecido', 'arquivo': nome, 'raiz': tag}]

def _nfe(root, caminho):
    inf = root.find('.//infNFe')
    ide = inf.find('ide'); emit = inf.find('emit'); dest = inf.find('dest'); tot = inf.find('total/ICMSTot')
    chave = (inf.get('Id') or '').replace('NFe', '') or _t(root, './/protNFe/infProt/chNFe')
    d = {
        'tipo': 'nfe', 'arquivo': os.path.basename(caminho), 'chave': chave,
        'modelo': _t(ide, 'mod'), 'serie': _t(ide, 'serie'), 'numero': _t(ide, 'nNF'),
        'emissao': (_t(ide, 'dhEmi') or _t(ide, 'dEmi'))[:10], 'tpNF': _t(ide, 'tpNF'), 'natOp': _t(ide, 'natOp'),
        'idDest': _t(ide, 'idDest'), 'indFinal': _t(ide, 'indFinal'), 'finNFe': _t(ide, 'finNFe'),
        'emit': {'cnpj': _cnpj(emit), 'nome': _t(emit, 'xNome'), 'uf': _t(emit, 'enderEmit/UF'), 'ie': _t(emit, 'IE'), 'crt': _t(emit, 'CRT'), 'mun': _t(emit, 'enderEmit/cMun')},
        'dest': {'cnpj': _cnpj(dest) if dest is not None else '', 'nome': _t(dest, 'xNome'), 'uf': _t(dest, 'enderDest/UF'), 'ie': _t(dest, 'IE'), 'indIEDest': _t(dest, 'indIEDest'), 'mun': _t(dest, 'enderDest/cMun')},
        'totais': {k: _n(tot, k) for k in ['vBC','vICMS','vICMSDeson','vFCP','vBCST','vST','vFCPST','vProd','vFrete','vSeg','vDesc','vII','vIPI','vIPIDevol','vPIS','vCOFINS','vOutro','vNF','vFCPUFDest','vICMSUFDest','vICMSUFRemet']},
        'infAdic': _t(inf, 'infAdic/infCpl'),
        'itens': [],
        'ibscbs': inf.find('.//IBSCBS') is not None or inf.find('.//IBSCBSTot') is not None,
        'protocolo': _t(root, './/protNFe/infProt/cStat'),
        'refs': [(_t(r, 'refNFe') or '') for r in ide.findall('NFref')] if ide is not None else [],
    }
    for det in inf.findall('det'):
        p = det.find('prod'); imp = det.find('imposto')
        icms = None
        if imp is not None and imp.find('ICMS') is not None:
            icms = list(imp.find('ICMS'))[0] if len(list(imp.find('ICMS'))) else None
        ipi = imp.find('IPI') if imp is not None else None
        ipi_sub = None
        if ipi is not None:
            ipi_sub = ipi.find('IPITrib') if ipi.find('IPITrib') is not None else ipi.find('IPINT')
        pis = imp.find('PIS') if imp is not None else None
        pis_sub = list(pis)[0] if pis is not None and len(list(pis)) else None
        cof = imp.find('COFINS') if imp is not None else None
        cof_sub = list(cof)[0] if cof is not None and len(list(cof)) else None
        ufdest = imp.find('ICMSUFDest') if imp is not None else None
        ibs = det.find('.//IBSCBS')
        item = {
            'nItem': det.get('nItem'), 'cProd': _t(p, 'cProd'), 'xProd': _t(p, 'xProd'), 'NCM': _t(p, 'NCM'), 'CEST': _t(p, 'CEST'),
            'CFOP': _t(p, 'CFOP'), 'uCom': _t(p, 'uCom'), 'qCom': _n(p, 'qCom'), 'vUnCom': _n(p, 'vUnCom'), 'vProd': _n(p, 'vProd'),
            'vFrete': _n(p, 'vFrete'), 'vSeg': _n(p, 'vSeg'), 'vDesc': _n(p, 'vDesc'), 'vOutro': _n(p, 'vOutro'), 'indTot': _t(p, 'indTot'),
            'infAdProd': _t(det, 'infAdProd'), 'vIPIDevol': _n(det, 'impostoDevol/IPI/vIPIDevol'), 'pDevol': _n(det, 'impostoDevol/pDevol'),
            'icms': {}, 'ipi': {}, 'pis': {}, 'cofins': {}, 'ufdest': {}, 'ibscbs': ibs is not None,
        }
        if icms is not None:
            item['icms'] = {'grupo': icms.tag, 'orig': _t(icms, 'orig'), 'CST': _t(icms, 'CST') or _t(icms, 'CSOSN'), 'modBC': _t(icms, 'modBC'),
                            'pRedBC': _n(icms, 'pRedBC'), 'vBC': _n(icms, 'vBC'), 'pICMS': _n(icms, 'pICMS'), 'vICMS': _n(icms, 'vICMS'),
                            'vBCST': _n(icms, 'vBCST'), 'pMVAST': _n(icms, 'pMVAST'), 'pICMSST': _n(icms, 'pICMSST'), 'vICMSST': _n(icms, 'vICMSST'),
                            'vFCP': _n(icms, 'vFCP'), 'vFCPST': _n(icms, 'vFCPST'), 'vICMSDeson': _n(icms, 'vICMSDeson'), 'motDesICMS': _t(icms, 'motDesICMS'),
                            'vBCSTRet': _n(icms, 'vBCSTRet'), 'vICMSSTRet': _n(icms, 'vICMSSTRet'), 'pCredSN': _n(icms, 'pCredSN'), 'vCredICMSSN': _n(icms, 'vCredICMSSN')}
        if ipi_sub is not None:
            item['ipi'] = {'CST': _t(ipi_sub, 'CST'), 'vBC': _n(ipi_sub, 'vBC'), 'pIPI': _n(ipi_sub, 'pIPI'), 'vIPI': _n(ipi_sub, 'vIPI'), 'cEnq': _t(ipi, 'cEnq')}
        if pis_sub is not None:
            item['pis'] = {'CST': _t(pis_sub, 'CST'), 'vBC': _n(pis_sub, 'vBC'), 'pPIS': _n(pis_sub, 'pPIS'), 'vPIS': _n(pis_sub, 'vPIS')}
        if cof_sub is not None:
            item['cofins'] = {'CST': _t(cof_sub, 'CST'), 'vBC': _n(cof_sub, 'vBC'), 'pCOFINS': _n(cof_sub, 'pCOFINS'), 'vCOFINS': _n(cof_sub, 'vCOFINS')}
        if ufdest is not None:
            item['ufdest'] = {'vBCUFDest': _n(ufdest, 'vBCUFDest'), 'pICMSUFDest': _n(ufdest, 'pICMSUFDest'), 'pICMSInter': _n(ufdest, 'pICMSInter'),
                              'vICMSUFDest': _n(ufdest, 'vICMSUFDest'), 'vICMSUFRemet': _n(ufdest, 'vICMSUFRemet'), 'vFCPUFDest': _n(ufdest, 'vFCPUFDest')}
        d['itens'].append(item)
    return d

def _cte(root, caminho):
    inf = root.find('.//infCte')
    ide = inf.find('ide'); emit = inf.find('emit'); rem = inf.find('rem'); dest = inf.find('dest'); exped = inf.find('exped'); receb = inf.find('receb')
    chave = (inf.get('Id') or '').replace('CTe', '')
    toma = ''
    t3 = ide.find('toma3') if ide is not None else None
    t4 = ide.find('toma4') if ide is not None else None
    tomador_cnpj = ''
    if t3 is not None:
        toma = _t(t3, 'toma')
    elif t4 is not None:
        toma = _t(t4, 'toma'); tomador_cnpj = _cnpj(t4)
    mapa = {'0': rem, '1': exped, '2': receb, '3': dest}
    if not tomador_cnpj and toma in mapa and mapa[toma] is not None:
        tomador_cnpj = _cnpj(mapa[toma])
    icms_el = inf.find('imp/ICMS')
    icms = list(icms_el)[0] if icms_el is not None and len(list(icms_el)) else None
    d = {
        'tipo': 'cte', 'arquivo': os.path.basename(caminho), 'chave': chave, 'modelo': _t(ide, 'mod'), 'serie': _t(ide, 'serie'), 'numero': _t(ide, 'nCT'),
        'emissao': _t(ide, 'dhEmi')[:10], 'CFOP': _t(ide, 'CFOP'), 'natOp': _t(ide, 'natOp'), 'tpCTe': _t(ide, 'tpCTe'), 'tpServ': _t(ide, 'tpServ'),
        'UFIni': _t(ide, 'UFIni'), 'UFFim': _t(ide, 'UFFim'), 'cMunIni': _t(ide, 'cMunIni'), 'cMunFim': _t(ide, 'cMunFim'),
        'emit': {'cnpj': _cnpj(emit), 'nome': _t(emit, 'xNome'), 'uf': _t(emit, 'enderEmit/UF'), 'crt': _t(emit, 'CRT')},
        'rem': {'cnpj': _cnpj(rem) if rem is not None else '', 'uf': _t(rem, 'enderReme/UF')},
        'dest': {'cnpj': _cnpj(dest) if dest is not None else '', 'uf': _t(dest, 'enderDest/UF')},
        'tomador': toma, 'tomador_cnpj': tomador_cnpj,
        'vTPrest': _n(inf, 'vPrest/vTPrest'), 'vRec': _n(inf, 'vPrest/vRec'),
        'icms': {'grupo': icms.tag if icms is not None else '', 'CST': _t(icms, 'CST'), 'vBC': _n(icms, 'vBC'), 'pICMS': _n(icms, 'pICMS'), 'vICMS': _n(icms, 'vICMS'), 'pRedBC': _n(icms, 'pRedBC'),
                 'vBCOutraUF': _n(icms, 'vBCOutraUF'), 'pICMSOutraUF': _n(icms, 'pICMSOutraUF'), 'vICMSOutraUF': _n(icms, 'vICMSOutraUF')} if icms is not None else {},
        'nfes': [(_t(x, 'chave') or '') for x in inf.findall('.//infNFe')],
        'vCarga': _n(inf, './/infCarga/vCarga'),
        'protocolo': _t(root, './/protCTe/infProt/cStat'),
    }
    return d

def _evento(root, caminho):
    inf = root.find('.//infEvento')
    tp = _t(inf, 'tpEvento') or _t(inf, 'detEvento/tpEvento')
    if not tp and (root.find('.//e101101') is not None or root.find('.//e105102') is not None):
        tp = '110111'  # cancelamento de NFS-e nacional
    return {'tipo': 'evento', 'arquivo': os.path.basename(caminho),
            'chave': _t(inf, 'chNFe') or _t(inf, 'chCTe') or _t(inf, 'chNFSe'),
            'tpEvento': tp, 'cancelamento': tp in ('110111', '110112'),
            'descEvento': _t(inf, 'detEvento/descEvento'), 'xCorrecao': _t(inf, 'detEvento/xCorrecao'),
            'dh': _t(inf, 'dhEvento')[:10], 'cStat': _t(root, './/retEvento/infEvento/cStat')}

def _nfse_nacional(root, caminho):
    inf = root.find('.//infNFSe'); dps = root.find('.//infDPS')
    emit = inf.find('emit') if inf is not None else None
    prest = dps.find('prest') if dps is not None else None
    toma = dps.find('toma') if dps is not None else None
    serv = dps.find('serv') if dps is not None else None
    val = dps.find('valores') if dps is not None else None
    vs = inf.find('valores') if inf is not None else None
    fed = val.find('trib/tribFed') if val is not None else None
    pc = fed.find('piscofins') if fed is not None else None
    tp_ret = _t(val, 'trib/tribMun/tpRetISSQN')
    vpis = _n(pc, 'vPis'); vcof = _n(pc, 'vCofins')
    ret_pc = _t(pc, 'tpRetPisCofins') in ('1', '3')  # 1 retido, 2 não retido (leiaute 1.0)
    return {
        'tipo': 'nfse', 'leiaute': 'nacional', 'arquivo': os.path.basename(caminho),
        'chave': (inf.get('Id') or '').replace('NFS', '') if inf is not None else '',
        'numero': _t(inf, 'nNFSe'), 'emissao': (_t(dps, 'dhEmi') or _t(inf, 'dhProc'))[:10], 'competencia': _t(dps, 'dCompet'),
        'prestador': {'cnpj': _cnpj(emit) or _cnpj(prest), 'nome': _t(emit, 'xNome') or _t(prest, 'xNome'), 'uf': _t(emit, 'enderNac/UF')},
        'tomador': {'cnpj': _cnpj(toma) if toma is not None else '', 'nome': _t(toma, 'xNome')},
        'cTribNac': _t(serv, 'cServ/cTribNac'), 'xDescServ': _t(serv, 'cServ/xDescServ'),
        'vServ': _n(val, 'vServPrest/vServ') or _n(vs, 'vCalcDR'), 'vISS': _n(vs, 'vISSQN') or _n(val, 'trib/tribMun/vISSQN'),
        'pAliq': _n(vs, 'pAliqAplic') or _n(val, 'trib/tribMun/pAliq'), 'vLiq': _n(vs, 'vLiq'), 'vBC': _n(vs, 'vBC'),
        'issRetido': tp_ret in ('2', '3'),
        'ret': {'irrf': _n(fed, 'vRetIRRF'), 'csll': _n(fed, 'vRetCSLL'), 'inss': _n(fed, 'vRetCP'),
                'pis': vpis if ret_pc else 0.0, 'cofins': vcof if ret_pc else 0.0},
        'pisCofins': {'CST': _t(pc, 'CST'), 'vPis': vpis, 'vCofins': vcof},
        'cancelada': False,
    }

def _nfse_abrasf(root, caminho):
    inf = root.find('.//InfNfse')
    if inf is None:
        return {'tipo': 'desconhecido', 'arquivo': os.path.basename(caminho), 'raiz': root.tag}
    serv = inf.find('.//Servico')
    if serv is None:
        serv = root.find('.//Servico')
    val = serv.find('Valores') if serv is not None else None
    def cnpj_de(*paths):
        for p in paths:
            v = re.sub(r'\D', '', _t(inf, p))
            if v:
                return v
        return ''
    iss_ret = (_t(val, 'IssRetido') or _t(serv, 'IssRetido')) == '1'
    return {
        'tipo': 'nfse', 'leiaute': 'abrasf', 'arquivo': os.path.basename(caminho), 'chave': _t(inf, 'CodigoVerificacao'),
        'numero': _t(inf, 'Numero'), 'emissao': _t(inf, 'DataEmissao')[:10], 'competencia': _t(inf, 'Competencia')[:10] or _t(inf, './/Competencia')[:10],
        'prestador': {'cnpj': cnpj_de('.//PrestadorServico/IdentificacaoPrestador/CpfCnpj/Cnpj', './/PrestadorServico/IdentificacaoPrestador/Cnpj',
                                      './/IdentificacaoPrestador/CpfCnpj/Cnpj', './/Prestador/CpfCnpj/Cnpj', './/IdentificacaoPrestador/Cnpj'),
                      'nome': _t(inf, './/PrestadorServico/RazaoSocial'), 'uf': _t(inf, './/PrestadorServico/Endereco/Uf')},
        'tomador': {'cnpj': cnpj_de('.//TomadorServico/IdentificacaoTomador/CpfCnpj/Cnpj', './/Tomador/IdentificacaoTomador/CpfCnpj/Cnpj',
                                    './/TomadorServico/IdentificacaoTomador/CpfCnpj/Cpf', './/Tomador/IdentificacaoTomador/CpfCnpj/Cpf'),
                    'nome': _t(inf, './/TomadorServico/RazaoSocial') or _t(inf, './/Tomador/RazaoSocial')},
        'cTribNac': _t(serv, 'ItemListaServico'), 'xDescServ': _t(serv, 'Discriminacao'),
        'vServ': _n(val, 'ValorServicos'), 'vISS': _n(val, 'ValorIss') or _n(val, 'ValorIssRetido'), 'pAliq': (lambda x: x * 100 if 0 < x < 1 else x)(_n(val, 'Aliquota')), 'vBC': _n(val, 'BaseCalculo'),
        'vLiq': _n(val, 'ValorLiquidoNfse') or _n(inf, './/ValorLiquidoNfse'),
        'issRetido': iss_ret,
        'ret': {'irrf': _n(val, 'ValorIr'), 'csll': _n(val, 'ValorCsll'), 'inss': _n(val, 'ValorInss'),
                'pis': _n(val, 'ValorPis'), 'cofins': _n(val, 'ValorCofins')},
        'pisCofins': {'CST': '', 'vPis': _n(val, 'ValorPis'), 'vCofins': _n(val, 'ValorCofins')},
        'cancelada': root.find('.//NfseCancelamento') is not None,
    }
