# -*- coding: utf-8 -*-
"""Destino provável do item comprado, pela descrição, NCM e valor unitário: combustível, uso e consumo
(alimentação, escritório, limpeza, EPI, ferramenta) ou ativo imobilizado. Se a empresa VENDE aquela NCM
(aparece nas saídas), o item é tratado como mercadoria/insumo dela e não recebe classificação."""
import re
from .util import norm

LIMITE_ATIVO = 1200.0  # valor unitário a partir do qual o bem não pode ir direto para despesa (RIR/2018, art. 313)

# (categoria, rótulo, palavras-chave, prefixos de NCM)
REGRAS = [
    ('combustivel', 'combustível', ['gasolina', 'etanol', 'diesel', 'alcool hidratado', 'alcool comum', 'gnv', 'biodiesel'],
     ['271012', '27101921', '27101922', '22071', '22072', '271121', '38260']),
    ('consumo', 'alimentação/refeição', ['refeicao', 'refeicoes', 'almoco', 'jantar', 'lanche', 'marmita', 'marmitex', 'salgado', 'sanduiche',
                                        'agua mineral', 'agua sem gas', 'agua com gas', 'refrigerante', 'suco', 'cafe', 'acucar', 'biscoito',
                                        'chocolate', 'energetico', 'red bull', 'cerveja', 'bombom', 'bala', 'chiclete', 'salgadinho', 'prestigio'],
     ['0201', '0202', '0207', '0401', '0402', '0406', '0901', '1701', '1704', '1806', '1905', '2009', '2101', '2106', '2201', '2202', '2203']),
    ('consumo', 'material de escritório', ['papel a4', 'papel sulfite', 'sulfite', 'caneta', 'lapis', 'grampeador', 'grampo', 'clips', 'envelope',
                                          'toner', 'cartucho', 'post it', 'caderno', 'bloco de notas', 'pasta suspensa', 'etiqueta adesiva'],
     ['4802', '4817', '4820', '9608', '9609', '9610', '3215']),
    ('consumo', 'limpeza/higiene', ['detergente', 'desinfetante', 'sabao', 'sabonete', 'papel higienico', 'papel toalha', 'alcool gel', 'alcool 70',
                                   'vassoura', 'rodo', 'saco de lixo', 'saco lixo', 'agua sanitaria', 'limpador', 'multiuso', 'esponja', 'pano de chao'],
     ['3401', '3402', '4818', '9603', '380894']),
    ('consumo', 'EPI/uniforme', ['luva', 'bota', 'botina', 'capacete', 'oculos de protecao', 'oculos seguranca', 'protetor auricular', 'mascara',
                                'uniforme', 'jaleco', 'avental', 'camiseta', 'calca brim'],
     ['6116', '6402', '6403', '6506', '900490']),
    ('consumo', 'ferramenta', ['broca', 'serra copo', 'serra', 'disco de corte', 'disco corte', 'lixa', 'alicate', 'chave de fenda', 'chave phillips',
                              'martelo', 'trena', 'estilete', 'pincel', 'jogo de chaves', 'soquete'],
     ['8202', '8203', '8204', '8205', '8207', '6804', '6805']),
]
# ativo: NCM típica de bem durável + valor unitário alto, ou palavra-chave forte + valor unitário alto
ATIVO_NCM = ['8427', '8471', '8415', '8418', '8450', '8528', '8517', '8443', '8467', '8456', '8457', '8458', '8459', '8460', '8461',
             '8462', '8463', '8464', '8465', '8414', '8502', '8701', '8703', '8704', '8711', '9401', '9403', '8525', '8521']
ATIVO_PALAVRAS = ['empilhadeira', 'maquina', 'equipamento', 'compressor', 'notebook', 'computador', 'impressora', 'monitor', 'ar condicionado',
                  'gerador', 'torno', 'fresadora', 'servidor', 'veiculo', 'automovel', 'caminhao', 'mesa', 'cadeira', 'armario', 'estante',
                  'geladeira', 'refrigerador', 'televisor', 'smart tv', 'celular', 'smartphone', 'furadeira', 'parafusadeira', 'esmerilhadeira',
                  'prensa', 'balanca', 'transpaleteira', 'paleteira']

_RE = {}


def _tem(texto, palavra):
    r = _RE.get(palavra)
    if r is None:
        r = _RE[palavra] = re.compile(r'(?<![a-z0-9])' + re.escape(palavra) + r'(?![a-z0-9])')
    return bool(r.search(texto))


def classificar(item, ncms_vendidos=frozenset()):
    """-> (categoria, rótulo, motivo) ou None. categoria: combustivel | consumo | ativo."""
    desc = norm(item.get('xProd', ''))
    ncm = (item.get('NCM') or '').strip()
    if ncm[:4] in ncms_vendidos:
        return None
    vun = item.get('vUnCom') or 0
    if not vun and item.get('qCom'):
        vun = (item.get('vProd') or 0) / item['qCom']
    for cat, rot, pals, ncms in REGRAS:
        if cat == 'combustivel' and ('lubrificante' in desc or 'oleo motor' in desc or 'aditivo' in desc and 'gasolina' not in desc):
            continue
        p = next((x for x in pals if _tem(desc, x)), None)
        n = next((x for x in ncms if ncm.startswith(x)), None)
        if p or n:
            if cat == 'consumo' and rot == 'ferramenta' and vun >= LIMITE_ATIVO:
                return ('ativo', 'ferramenta/equipamento durável', 'valor unitário %s' % _brl(vun))
            return (cat, rot, ('descrição "%s"' % p) if p else ('NCM %s' % ncm))
    if vun >= LIMITE_ATIVO:
        n = next((x for x in ATIVO_NCM if ncm.startswith(x)), None)
        p = next((x for x in ATIVO_PALAVRAS if _tem(desc, x)), None)
        if n or p:
            return ('ativo', 'ativo imobilizado', '%s, valor unitário %s' % (('descrição "%s"' % p) if p else ('NCM %s' % ncm), _brl(vun)))
    return None


def _brl(v):
    return 'R$ ' + '{:,.2f}'.format(v).replace(',', 'X').replace('.', ',').replace('X', '.')
