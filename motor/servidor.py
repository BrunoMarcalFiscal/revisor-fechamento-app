# -*- coding: utf-8 -*-
"""Servidor local do Revisor do Fechamento (só biblioteca padrão + openpyxl/PyMuPDF já instalados)."""
import os, json, zipfile, io, shutil, re, traceback, urllib.parse
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from . import comparador, planilha, exportar, relatorios, estrategico, scp, scp_planilha

RAIZ = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
# dados (sessão, resultados, histórico, configurações) podem ficar fora do código: REVISOR_DADOS (disco do servidor)
DADOS = os.path.abspath(os.environ.get('REVISOR_DADOS') or RAIZ)
SESSAO = os.path.join(DADOS, 'sessao')
SAIDA = os.path.join(DADOS, 'saida')
WEB = os.path.join(RAIZ, 'web')
CFG = os.path.join(SESSAO, 'config.json')
AJUSTES = os.path.join(SESSAO, 'ajustes_planilha.json')
CORRIGIDOS = os.path.join(SESSAO, 'corrigidos.json')
APURACAO = os.path.join(SESSAO, 'apuracao.json')
SCP_CFG = os.path.join(SESSAO, 'scp.json')
MANUAIS = os.path.join(SESSAO, 'manuais.json')
CANCELADAS = os.path.join(SESSAO, 'canceladas.json')
HIST = os.path.join(DADOS, 'fechamentos')
ADM_ARQ = os.path.join(DADOS, 'config', 'admin.json')
ADM_TOKENS = set()

# ------------------------------------------------------------------ acesso restrito (colaborador)
# REVISOR_PERFIL=colaborador: exige login (REVISOR_USUARIO / REVISOR_SENHA), só a empresa de REVISOR_EMPRESA (CNPJ)
# e sem Estratégia SCP nem Visão estratégica (bloqueado no servidor, não só escondido na tela).
PERFIL = (os.environ.get('REVISOR_PERFIL') or 'completo').strip().lower()
EMPRESA_OK = re.sub(r'\D', '', os.environ.get('REVISOR_EMPRESA') or '')
USUARIO = os.environ.get('REVISOR_USUARIO') or ''
SENHA = os.environ.get('REVISOR_SENHA') or ''
LOGIN_ON = bool(USUARIO and SENHA)
SESSOES = {}
BLOQ_COLAB = ('/api/scp', '/api/estrategico', '/api/adm')


def colaborador():
    return PERFIL == 'colaborador'


def _cookie_tok(handler):
    for parte in (handler.headers.get('Cookie') or '').split(';'):
        k, _, v = parte.strip().partition('=')
        if k == 'rf_sessao':
            return v
    return ''


def logado(handler):
    return not LOGIN_ON or _cookie_tok(handler) in SESSOES


def empresa_permitida(cnpj):
    return not EMPRESA_OK or re.sub(r'\D', '', cnpj or '')[:8] == EMPRESA_OK[:8]


LOGIN_HTML = '''<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Revisor do Fechamento</title>
<style>:root{--navy:#0E2841;--borda:#E1E6EE}body{margin:0;min-height:100vh;display:flex;align-items:center;justify-content:center;background:#F4F6FA;font:14px/1.5 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;color:#16202E}
form{background:#fff;border:1px solid var(--borda);border-radius:10px;padding:28px 26px;width:min(360px,calc(100vw - 32px));box-sizing:border-box}
.top{background:var(--navy);margin:-28px -26px 22px;padding:16px 26px;border-radius:10px 10px 0 0}.top img{height:26px}
h1{font-size:17px;margin:0 0 4px;color:var(--navy)}p{margin:0 0 18px;color:#5A6476;font-size:12.5px}
label{display:block;font-size:12px;font-weight:600;color:#5A6476;margin:0 0 12px}input{display:block;width:100%;box-sizing:border-box;margin-top:4px;border:1px solid var(--borda);border-radius:6px;padding:9px 10px;font:inherit}
input:focus{outline:2px solid #BFD6F2;border-color:#2563A8}button{width:100%;background:var(--navy);color:#fff;border:0;border-radius:6px;padding:10px;font:inherit;font-weight:600;cursor:pointer;margin-top:6px}
.erro{color:#B42318;font-size:12.5px;margin:0 0 10px}</style></head><body>
<form method="post" action="/login"><div class="top"><img src="/web/logo-branca.png" alt="Marçal" onerror="this.style.display='none'"></div>
<h1>Revisor do Fechamento</h1><p>Entre com o usuário e a senha que o escritório passou.</p>__ERRO__
<label>Usuário<input name="usuario" autocomplete="username" required autofocus></label><label>Senha<input name="senha" type="password" autocomplete="current-password" required></label>
<button type="submit">Entrar</button></form></body></html>'''


def _hash_pin(pin, sal):
    import hashlib
    return hashlib.pbkdf2_hmac('sha256', pin.encode(), bytes.fromhex(sal), 200000).hex()


def adm_ok(handler):
    return handler.headers.get('X-ADM', '') in ADM_TOKENS


def _id_fech(res):
    return '%s_%s' % ((res.get('empresa') or {}).get('cnpj') or 'semcnpj', res.get('competencia') or 'semmes')


def arquivar(res=None):
    """Guarda a sessão atual (arquivos, ajustes, correções e resultado) em fechamentos/<cnpj>_<aaaa-mm>/."""
    res = res or _json(os.path.join(SAIDA, 'resultado.json'), None)
    if not res or not res.get('notas'):
        return None
    fid = _id_fech(res)
    dest = os.path.join(HIST, fid)
    meta_ant = _json(os.path.join(dest, 'meta.json'), {})
    if os.path.isdir(dest):
        shutil.rmtree(dest)
    shutil.copytree(SESSAO, os.path.join(dest, 'sessao'))
    os.makedirs(dest, exist_ok=True)
    shutil.copy(os.path.join(SAIDA, 'resultado.json'), os.path.join(dest, 'resultado.json'))
    cor = _json(CORRIGIDOS, {})
    pend = sum(1 for n in res['notas'] for a in n['aponts'] if a['sev'] in ('erro', 'alerta') and a.get('k') not in cor)
    totap = sum(1 for n in res['notas'] for a in n['aponts'] if a['sev'] in ('erro', 'alerta'))
    import datetime
    meta = {'id': fid, 'cnpj': res['empresa'].get('cnpj'), 'empresa': res['empresa'].get('nome'), 'uf': res['empresa'].get('uf'), 'competencia': res.get('competencia'),
            'salvo_em': datetime.datetime.now().strftime('%d/%m/%Y %H:%M'), 'status': meta_ant.get('status', 'em revisão'),
            'apontamentos': totap, 'pendentes': pend, 'notas': res['resumo']['notas'],
            'indicadores': estrategico.indicadores(res, (_json(APURACAO, {}) or {}).get('calc')), 'recebido_em': recebido_em(),
            'fechado_em': meta_ant.get('fechado_em', '')}
    json.dump(meta, open(os.path.join(dest, 'meta.json'), 'w', encoding='utf-8'), ensure_ascii=False)
    return meta


def recebido_em():
    """Data do primeiro arquivo recebido na sessão (para medir o tempo do fechamento)."""
    import datetime
    ts = []
    for c in comparador.CARDS:
        pasta = os.path.join(SESSAO, c)
        if os.path.isdir(pasta):
            ts += [os.path.getmtime(os.path.join(pasta, f)) for f in os.listdir(pasta) if not f.startswith('.')]
    return datetime.datetime.fromtimestamp(min(ts)).strftime('%d/%m/%Y %H:%M') if ts else ''


def _comp_ant(comp, k=1):
    a, m = int(comp[:4]), int(comp[5:7])
    for _ in range(k):
        a, m = (a - 1, 12) if m == 1 else (a, m - 1)
    return '%04d-%02d' % (a, m)


def calcular_scp(res, cfg_novo=None):
    """Estratégia SCP: junta o que foi salvo, o que veio do histórico (saldo credor do mês anterior e meses do trimestre) e calcula."""
    salvo = _json(SCP_CFG, {})
    comp0, cnpj0 = res.get('competencia') or '', (res.get('empresa') or {}).get('cnpj')
    if salvo and (salvo.get('cnpj'), salvo.get('competencia')) != (cnpj0, comp0):
        salvo = {}  # configuração de outra empresa/competência (ou sem marca): não vale para este fechamento
    cfg = dict(salvo.get('cfg') or {})
    if not salvo and len(comp0) >= 7:
        # mês novo: herda do mês anterior as escolhas que não mudam (SCP, CFOP de crédito, critérios, regime anterior, compensações)
        ant = _json(os.path.join(HIST, '%s_%s' % (cnpj0, _comp_ant(comp0)), 'sessao', 'scp.json'), {})
        for kk, v in (ant.get('cfg') or {}).items():
            if kk in ('scp_cnpj', 'scp_nome', 'matriz_nome', 'creditos_cfop', 'excluir_ipi_credito', 'excluir_ipi_debito', 'excluir_st_debito', 'deduzir_devolucoes', 'importacao',
                      'excluir_pf', 'prioridade', 'modo', 'compensa', 'ant_carga_pct', 'retidas_scp', 'exito_pct') or kk.startswith('ant_'):
                cfg[kk] = v
    if cfg_novo is not None:
        cfg.update(cfg_novo)
    dados = scp.extrair(res)
    comp, cnpj = res.get('competencia') or '', (res.get('empresa') or {}).get('cnpj')
    auto = {}
    if len(comp) >= 7:
        ant = _json(os.path.join(HIST, '%s_%s' % (cnpj, _comp_ant(comp)), 'sessao', 'scp.json'), {})
        mt = (ant.get('calc') or {}).get('matriz') or {}
        if mt:
            auto['saldo_pis_ant'] = mt['pis']['saldo_credor_final']; auto['saldo_cof_ant'] = mt['cofins']['saldo_credor_final']
            auto['saldo_origem'] = _comp_ant(comp)
        mes = int(comp[5:7]); ini = mes - (mes - 1) % 3
        outros = []
        for k in range(1, mes - ini + 1):
            ca = _comp_ant(comp, k)
            h = _json(os.path.join(HIST, '%s_%s' % (cnpj, ca), 'sessao', 'scp.json'), {})
            rs_ = (h.get('calc') or {}).get('resumo') or {}
            if rs_:
                outros.append({'competencia': ca, 'receita': rs_['scp']['receita'], 'irrf': (h['calc'].get('retencoes') or {}).get('scp', {}).get('irrf', 0),
                               'csll': (h['calc'].get('retencoes') or {}).get('scp', {}).get('csll', 0), 'auto': True,
                               'receita_matriz': (rs_.get('matriz') or {}).get('receita', 0), 'lucro': (h.get('cfg') or {}).get('lucro_matriz') or 0,
                               'irrf_m': (h['calc'].get('retencoes') or {}).get('matriz', {}).get('irrf', 0), 'csll_m': (h['calc'].get('retencoes') or {}).get('matriz', {}).get('csll', 0),
                               **scp_planilha.economia_do_mes(h['calc'].get('planilha') or [])})
        if outros:
            auto['tri_outros_meses'] = outros
    efetivo = dict(cfg)
    for k, v in auto.items():
        if k == 'saldo_origem':
            continue
        if k in ('saldo_pis_ant', 'saldo_cof_ant') and cfg.get('saldo_manual'):
            continue
        if k == 'tri_outros_meses' and cfg.get('tri_manual'):
            continue
        efetivo[k] = v
    calc = scp.calcular(dados, efetivo)
    calc['auto'] = auto
    calc['ret_a_recolher'] = ((res.get('painel') or {}).get('impostos') or {}).get('ret_a_recolher', 0)
    calc['planilha'] = scp_planilha.para_json(scp_planilha.montar(res, calc, _json(APURACAO, {})))
    calc['detalhes'] = scp_planilha.detalhes(res, calc)
    calc['fontes'] = {'prestados': len(dados['prestados']), 'tomados': len(dados['tomados']), 'locacoes': len(dados['locacoes']),
                      'fonte_prestados': dados['prestados'][0]['fonte'] if dados['prestados'] else '', 'fonte_tomados': dados['tomados'][0]['fonte'] if dados['tomados'] else '',
                      'icms_pis_saidas': dados.get('icms_pis_saidas')}
    os.makedirs(SESSAO, exist_ok=True)
    json.dump({'cnpj': cnpj0, 'competencia': comp0, 'cfg': cfg, 'calc': calc}, open(SCP_CFG, 'w', encoding='utf-8'), ensure_ascii=False, default=list)
    return calc


def listar_fechamentos():
    out = []
    if os.path.isdir(HIST):
        for d in sorted(os.listdir(HIST)):
            m = _json(os.path.join(HIST, d, 'meta.json'), None)
            if m:
                out.append(m)
    return sorted(out, key=lambda m: (m.get('empresa') or '', m.get('competencia') or ''), reverse=False)


def saldo_anterior(res):
    """Saldo credor de ICMS/IPI do fechamento do mês anterior da mesma empresa (se houver)."""
    comp = res.get('competencia') or ''
    if len(comp) < 7:
        return None
    a, m = int(comp[:4]), int(comp[5:7])
    a, m = (a - 1, 12) if m == 1 else (a, m - 1)
    fid = '%s_%04d-%02d' % (res['empresa'].get('cnpj'), a, m)
    ap = _json(os.path.join(HIST, fid, 'sessao', 'apuracao.json'), {})
    calc = ap.get('calc') or {}
    if not calc:
        return None
    return {'icms': round(-calc['saldo'], 2) if calc.get('saldo', 0) < 0 else 0, 'ipi': round(-calc['ipiSaldo'], 2) if calc.get('ipiSaldo', 0) < 0 else 0,
            'origem': '%02d/%04d' % (m, a)}
EXTS = {'entrada': {'.xml'}, 'saida': {'.xml'}, 'frete': {'.xml'}, 'prestado': {'.xml'}, 'tomado': {'.xml'},
        'diversos': {'.pdf', '.xml'}, 'planilha': {'.xlsx', '.xlsm', '.xls', '.csv', '.txt'},
        'planilha_servicos': {'.xlsx', '.xlsm', '.xls', '.csv', '.txt', '.pdf'}, 'serv_entrada': {'.xlsx', '.xlsm', '.xls', '.csv', '.txt', '.pdf'},
        'serv_saida': {'.xlsx', '.xlsm', '.xls', '.csv', '.txt', '.pdf'}, 'planilha_locacoes': {'.xlsx', '.xlsm', '.xls', '.csv', '.txt'}}


def _json(path, padrao):
    try:
        return json.load(open(path, encoding='utf-8'))
    except Exception:
        return padrao


def _nome_seguro(nome):
    nome = os.path.basename(nome.replace('\\', '/'))
    return re.sub(r'[^\w.\- ()]', '_', nome) or 'arquivo'


def _salvar(card, nome, dados):
    pasta = os.path.join(SESSAO, card)
    os.makedirs(pasta, exist_ok=True)
    salvos = 0
    if nome.lower().endswith('.zip'):
        with zipfile.ZipFile(io.BytesIO(dados)) as z:
            for info in z.infolist():
                if info.is_dir() or '__MACOSX' in info.filename:
                    continue
                n = info.filename
                try:
                    n = n.encode('cp437').decode('utf-8')
                except Exception:
                    pass
                conteudo = z.read(info)
                if n.lower().endswith('.zip'):
                    salvos += _salvar(card, n, conteudo)
                    continue
                if os.path.splitext(n)[1].lower() in EXTS[card]:
                    salvos += _gravar(pasta, _nome_seguro(n), conteudo)
        return salvos
    if os.path.splitext(nome)[1].lower() in EXTS[card]:
        return _gravar(pasta, _nome_seguro(nome), dados)
    return 0


def _gravar(pasta, nome, dados):
    destino = os.path.join(pasta, nome)
    if os.path.exists(destino) and open(destino, 'rb').read() == dados:
        return 1
    base, ext = os.path.splitext(nome)
    k = 1
    while os.path.exists(destino):
        k += 1
        destino = os.path.join(pasta, '%s (%d)%s' % (base, k, ext))
    open(destino, 'wb').write(dados)
    return 1


def estado():
    cards = {}
    for card in comparador.CARDS:
        pasta = os.path.join(SESSAO, card)
        arqs = []
        if os.path.isdir(pasta):
            for f in sorted(os.listdir(pasta)):
                if not f.startswith('.'):
                    arqs.append({'nome': f, 'tamanho': os.path.getsize(os.path.join(pasta, f))})
        cards[card] = {'arquivos': arqs, 'qtd': len(arqs), 'bytes': sum(a['tamanho'] for a in arqs)}
    return {'cards': cards, 'manuais': len(_json(MANUAIS, []) or []), 'config': _json(CFG, {}), 'tem_resultado': os.path.exists(os.path.join(SAIDA, 'resultado.json'))}


def analisar_planilhas():
    ajustes = _json(AJUSTES, {})
    out = []
    for card, pref in comparador.PLANILHAS:
        pasta = os.path.join(SESSAO, card)
        if not os.path.isdir(pasta):
            continue
        for f in sorted(os.listdir(pasta)):
            if os.path.splitext(f)[1].lower() in EXTS[card]:
                try:
                    abas = planilha.analisar(os.path.join(pasta, f), ajustes.get(pref + f), servicos=(card in comparador.SERV_CARDS), locacoes=(card == 'planilha_locacoes'))
                except Exception as e:
                    abas = [{'arquivo': f, 'aba': '', 'erro': str(e), 'utilizavel': False}]
                for a in abas:
                    a['card'] = card
                out.extend(abas)
    return {'abas': out, 'campos': [{'campo': c, 'rotulo': r} for c, r, _ in planilha.CAMPOS]}


class H(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _env(self, codigo, corpo, tipo='application/json; charset=utf-8', extra=None):
        if isinstance(corpo, (dict, list)):
            corpo = json.dumps(corpo, ensure_ascii=False).encode('utf-8')
        elif isinstance(corpo, str):
            corpo = corpo.encode('utf-8')
        self.send_response(codigo)
        self.send_header('Content-Type', tipo)
        self.send_header('Content-Length', str(len(corpo)))
        self.send_header('Cache-Control', 'no-store')
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(corpo)

    def _corpo(self):
        n = int(self.headers.get('Content-Length') or 0)
        return self.rfile.read(n) if n else b''

    def _portao(self, metodo):
        """Login e bloqueios do acesso de colaborador, antes de qualquer rota."""
        u = urllib.parse.urlparse(self.path)
        q = dict(urllib.parse.parse_qsl(u.query))
        if LOGIN_ON and u.path == '/login':
            if metodo == 'GET':
                return self._env(200, LOGIN_HTML.replace('__ERRO__', ''), 'text/html; charset=utf-8') or True
            f = dict(urllib.parse.parse_qsl(self._corpo().decode('utf-8', 'ignore')))
            import secrets, hmac
            if hmac.compare_digest(f.get('usuario', ''), USUARIO) and hmac.compare_digest(f.get('senha', ''), SENHA):
                tok = secrets.token_urlsafe(32); SESSOES[tok] = f.get('usuario')
                self.send_response(303); self.send_header('Location', '/')
                self.send_header('Set-Cookie', 'rf_sessao=%s; Path=/; HttpOnly; SameSite=Lax%s; Max-Age=43200' % (tok, '; Secure' if os.environ.get('RENDER') else ''))
                self.send_header('Content-Length', '0'); self.end_headers()
                return True
            return self._env(401, LOGIN_HTML.replace('__ERRO__', '<div class="erro">Usuário ou senha incorretos.</div>'), 'text/html; charset=utf-8') or True
        if u.path == '/api/importar-fechamento' and metodo == 'POST':
            # carga inicial do servidor: zip de um fechamento (pasta fechamentos/<id>), protegida por token próprio
            import hmac
            tok = os.environ.get('REVISOR_IMPORT_TOKEN') or ''
            if not tok or not hmac.compare_digest(self.headers.get('X-Import-Token', ''), tok):
                return self._env(403, {'erro': 'token inválido'}) or True
            fid = _nome_seguro(q.get('id', ''))
            if not fid or not empresa_permitida(fid.split('_')[0]):
                return self._env(400, {'erro': 'fechamento de outra empresa'}) or True
            dest = os.path.join(HIST, fid)
            shutil.rmtree(dest, ignore_errors=True); os.makedirs(dest, exist_ok=True)
            with zipfile.ZipFile(io.BytesIO(self._corpo())) as z:
                for m in z.infolist():
                    alvo = os.path.abspath(os.path.join(dest, m.filename))
                    if not alvo.startswith(os.path.abspath(dest) + os.sep) or os.path.basename(alvo) == 'scp.json':
                        continue
                    if m.is_dir():
                        os.makedirs(alvo, exist_ok=True); continue
                    os.makedirs(os.path.dirname(alvo), exist_ok=True)
                    open(alvo, 'wb').write(z.read(m))
            return self._env(200, {'ok': True, 'id': fid, 'lista': [m['id'] for m in listar_fechamentos()]}) or True
        if LOGIN_ON and u.path == '/sair':
            SESSOES.pop(_cookie_tok(self), None)
            self.send_response(303); self.send_header('Location', '/login'); self.send_header('Set-Cookie', 'rf_sessao=; Path=/; Max-Age=0')
            self.send_header('Content-Length', '0'); self.end_headers()
            return True
        if not logado(self) and not u.path.startswith('/web/logo'):
            if u.path.startswith('/api/'):
                return self._env(401, {'erro': 'Sessão expirada. Entre de novo.'}) or True
            self.send_response(303); self.send_header('Location', '/login'); self.send_header('Content-Length', '0'); self.end_headers()
            return True
        if colaborador():
            if u.path.startswith(BLOQ_COLAB) or (u.path == '/api/relatorio' and q.get('tipo') in ('scp', 'estrategico')):
                return self._env(403, {'erro': 'Este acesso não inclui a Estratégia.'}) or True
        return False

    def do_GET(self):
        if self._portao('GET'):
            return
        u = urllib.parse.urlparse(self.path)
        q = dict(urllib.parse.parse_qsl(u.query))
        try:
            if u.path == '/api/perfil':
                return self._env(200, {'perfil': PERFIL, 'usuario': SESSOES.get(_cookie_tok(self), ''), 'login': LOGIN_ON, 'empresa': EMPRESA_OK})
            if u.path in ('/', '/index.html'):
                return self._env(200, open(os.path.join(WEB, 'index.html'), 'rb').read(), 'text/html; charset=utf-8')
            if u.path.startswith('/web/'):
                f = os.path.join(WEB, _nome_seguro(u.path[5:]))
                tipo = {'.png': 'image/png', '.js': 'text/javascript', '.css': 'text/css'}.get(os.path.splitext(f)[1], 'application/octet-stream')
                return self._env(200, open(f, 'rb').read(), tipo)
            if u.path == '/api/estado':
                return self._env(200, estado())
            if u.path == '/api/planilha':
                return self._env(200, analisar_planilhas())
            if u.path == '/api/resultado':
                r_ = _json(os.path.join(SAIDA, 'resultado.json'), {})
                if r_ and not empresa_permitida((r_.get('empresa') or {}).get('cnpj')):
                    return self._env(200, {})
                return self._env(200, r_)
            if u.path == '/api/apuracao':
                return self._env(200, _json(APURACAO, {}))
            if u.path == '/api/adm':
                return self._env(200, {'configurado': os.path.exists(ADM_ARQ), 'ativo': adm_ok(self)})
            if u.path == '/api/fechamentos':
                lst = [m for m in listar_fechamentos() if empresa_permitida(m.get('cnpj'))]
                if colaborador():
                    lst = [{k: v for k, v in m.items() if k != 'indicadores'} for m in lst]
                return self._env(200, {'lista': lst, 'atual': _json(os.path.join(SESSAO, 'meta.json'), {})})
            if u.path == '/api/scp':
                res = _json(os.path.join(SAIDA, 'resultado.json'), None)
                if not res:
                    return self._env(404, {'erro': 'Calcule antes.'})
                return self._env(200, calcular_scp(res))
            if u.path == '/api/manuais':
                return self._env(200, {'notas': _json(MANUAIS, [])})
            if u.path == '/api/estrategico':
                res = _json(os.path.join(SAIDA, 'resultado.json'), None)
                if not res:
                    return self._env(404, {'erro': 'Calcule antes.'})
                return self._env(200, estrategico.montar(res, _json(APURACAO, {}), listar_fechamentos(), recebido_em()))
            if u.path == '/api/relatorio':
                res = _json(os.path.join(SAIDA, 'resultado.json'), None)
                if not res:
                    return self._env(404, {'erro': 'Calcule antes de gerar o relatório.'})
                fmt = q.get('formato', 'json')
                if q.get('tipo') == 'scp':
                    res = dict(res, scp_calc=calcular_scp(res))
                if q.get('tipo') == 'estrategico':
                    res = dict(res, estrategico=estrategico.montar(res, _json(APURACAO, {}), listar_fechamentos(), recebido_em()))
                    if q.get('cbs'):
                        res['estrategico']['reforma']['ref_cbs'] = float(q['cbs'])
                    if q.get('ibs'):
                        res['estrategico']['reforma']['ref_ibs'] = float(q['ibs'])
                dados = relatorios.gerar(res, _json(APURACAO, {}), q.get('tipo', 'impostos'), fmt, receita=q.get('receita', 'consolidado'),
                                         direcao=q.get('direcao', 'ambos'), quem=q.get('quem', 'ambos'))
                if fmt == 'json':
                    return self._env(200, dados)
                nome = 'Relatorio_%s_%s_%s.%s' % (q.get('tipo', ''), res['empresa'].get('cnpj', ''), (res.get('competencia') or '').replace('-', ''), fmt)
                tipo_mime = 'application/pdf' if fmt == 'pdf' else 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                return self._env(200, dados, tipo_mime, {'Content-Disposition': '%s; filename="%s"' % ('inline' if fmt == 'pdf' else 'attachment', nome)})
            if u.path == '/api/corrigidos':
                return self._env(200, _json(CORRIGIDOS, {}))
            if u.path == '/api/exportar':
                res = _json(os.path.join(SAIDA, 'resultado.json'), None)
                if not res:
                    return self._env(404, {'erro': 'Calcule antes de exportar.'})
                dados = exportar.xlsx(res, _json(CORRIGIDOS, {}), _json(APURACAO, {}))
                nome = 'Revisao_fechamento_%s_%s.xlsx' % (res['empresa'].get('cnpj', ''), (res.get('competencia') or '').replace('-', ''))
                return self._env(200, dados, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                                 {'Content-Disposition': 'attachment; filename="%s"' % nome})
            return self._env(404, {'erro': 'não encontrado'})
        except Exception as e:
            traceback.print_exc()
            return self._env(500, {'erro': str(e)})

    def do_POST(self):
        if self._portao('POST'):
            return
        u = urllib.parse.urlparse(self.path)
        q = dict(urllib.parse.parse_qsl(u.query))
        try:
            if u.path == '/api/upload':
                card = q.get('card')
                if card not in EXTS:
                    return self._env(400, {'erro': 'card inválido'})
                n = _salvar(card, q.get('nome', 'arquivo'), self._corpo())
                return self._env(200, {'salvos': n, 'estado': estado()})
            if u.path == '/api/remover':
                card, nome = q.get('card'), _nome_seguro(q.get('nome', ''))
                f = os.path.join(SESSAO, card or '', nome)
                if card in EXTS and os.path.isfile(f):
                    os.remove(f)
                return self._env(200, estado())
            if u.path == '/api/limpar':
                card = q.get('card')
                if card == 'tudo':
                    arquivar()  # "Novo fechamento" guarda o atual no histórico antes de limpar
                    for f in (os.path.join(SESSAO, 'meta.json'),):
                        if os.path.exists(f):
                            os.remove(f)
                alvos = comparador.CARDS if card == 'tudo' else [card]
                for c in alvos:
                    if c in EXTS:
                        shutil.rmtree(os.path.join(SESSAO, c), ignore_errors=True)
                if (card == 'tudo' or card in dict(comparador.PLANILHAS)) and os.path.exists(AJUSTES):
                    os.remove(AJUSTES)
                if card == 'tudo':
                    for f in (CORRIGIDOS, APURACAO, SCP_CFG, MANUAIS, CANCELADAS):
                        if os.path.exists(f):
                            os.remove(f)
                if card == 'tudo':
                    for f in (os.path.join(SAIDA, 'resultado.json'), CFG):
                        if os.path.exists(f):
                            os.remove(f)
                return self._env(200, estado())
            if u.path in ('/api/adm/criar', '/api/adm/entrar', '/api/adm/sair'):
                import secrets
                b = json.loads(self._corpo() or b'{}')
                pin = str(b.get('pin') or '')
                if u.path.endswith('sair'):
                    ADM_TOKENS.discard(self.headers.get('X-ADM', ''))
                    return self._env(200, {'ok': True})
                if u.path.endswith('criar'):
                    if os.path.exists(ADM_ARQ):
                        return self._env(403, {'erro': 'O PIN do ADM já foi criado.'})
                    if len(pin) < 4:
                        return self._env(400, {'erro': 'O PIN precisa ter pelo menos 4 dígitos.'})
                    sal = secrets.token_hex(16)
                    os.makedirs(os.path.dirname(ADM_ARQ), exist_ok=True)
                    json.dump({'sal': sal, 'hash': _hash_pin(pin, sal)}, open(ADM_ARQ, 'w'))
                else:
                    cfg_adm = _json(ADM_ARQ, None)
                    if not cfg_adm or _hash_pin(pin, cfg_adm['sal']) != cfg_adm['hash']:
                        return self._env(403, {'erro': 'PIN incorreto.'})
                tk = secrets.token_hex(16)
                ADM_TOKENS.add(tk)
                return self._env(200, {'token': tk})
            if u.path == '/api/scp':
                res = _json(os.path.join(SAIDA, 'resultado.json'), None)
                if not res:
                    return self._env(404, {'erro': 'Calcule antes.'})
                if not adm_ok(self):
                    return self._env(403, {'erro': 'Só o ADM altera a Estratégia SCP. Entre no modo ADM.'})
                b = json.loads(self._corpo() or b'{}')
                if b.get('reset'):
                    if os.path.exists(SCP_CFG):
                        os.remove(SCP_CFG)
                    return self._env(200, calcular_scp(res, {}))
                return self._env(200, calcular_scp(res, b.get('cfg') or {}))
            if u.path == '/api/fechamento/status':
                b = json.loads(self._corpo() or b'{}')
                meta = arquivar() or {}
                if meta:
                    meta['status'] = b.get('status', 'em revisão')
                    if meta['status'] == 'fechado' and not meta.get('fechado_em'):
                        import datetime
                        meta['fechado_em'] = datetime.datetime.now().strftime('%d/%m/%Y %H:%M')
                    json.dump(meta, open(os.path.join(HIST, meta['id'], 'meta.json'), 'w', encoding='utf-8'), ensure_ascii=False)
                    json.dump(meta, open(os.path.join(SESSAO, 'meta.json'), 'w', encoding='utf-8'), ensure_ascii=False)
                return self._env(200, meta)
            if u.path == '/api/fechamento/abrir':
                fid = _nome_seguro(q.get('id', ''))
                orig = os.path.join(HIST, fid)
                if not os.path.isdir(orig) or not empresa_permitida(fid.split('_')[0]):
                    return self._env(404, {'erro': 'fechamento não encontrado'})
                arquivar()  # guarda o que está aberto antes de trocar
                shutil.rmtree(SESSAO, ignore_errors=True)
                shutil.copytree(os.path.join(orig, 'sessao'), SESSAO)
                os.makedirs(SAIDA, exist_ok=True)
                shutil.copy(os.path.join(orig, 'resultado.json'), os.path.join(SAIDA, 'resultado.json'))
                json.dump(_json(os.path.join(orig, 'meta.json'), {}), open(os.path.join(SESSAO, 'meta.json'), 'w', encoding='utf-8'), ensure_ascii=False)
                return self._env(200, estado())
            if u.path == '/api/apuracao':
                b = json.loads(self._corpo() or b'{}')
                os.makedirs(SESSAO, exist_ok=True)
                json.dump(b, open(APURACAO, 'w', encoding='utf-8'), ensure_ascii=False)
                return self._env(200, {'ok': True})
            if u.path == '/api/corrigido':
                # resolução de apontamento: {chaves, status: 'corrigido'|'justificado'|null, obs}; observação da nota: chave 'obs:<nota>'
                b = json.loads(self._corpo() or b'{}')
                c = _json(CORRIGIDOS, {})
                import datetime
                agora = datetime.datetime.now().strftime('%d/%m/%Y %H:%M')
                status = b.get('status', 'corrigido' if b.get('corrigido', True) else None)
                for k in b.get('chaves', []):
                    if status:
                        c[k] = {'status': status, 'obs': (b.get('obs') or '').strip(), 'em': agora}
                    else:
                        c.pop(k, None)
                os.makedirs(SESSAO, exist_ok=True)
                json.dump(c, open(CORRIGIDOS, 'w', encoding='utf-8'), ensure_ascii=False)
                return self._env(200, c)
            if u.path == '/api/planilha/ajuste':
                b = json.loads(self._corpo() or b'{}')
                aj = _json(AJUSTES, {})
                card = b.get('card') if b.get('card') in dict(comparador.PLANILHAS) else 'planilha'
                chave_aj = dict(comparador.PLANILHAS)[card] + b['arquivo']
                aj.setdefault(chave_aj, {})[b['aba']] = {'linha_cab': int(b['linha_cab']), 'mapa': {k: int(v) for k, v in b['mapa'].items() if v not in (None, '')}}
                os.makedirs(SESSAO, exist_ok=True)
                json.dump(aj, open(AJUSTES, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
                # memoriza o layout (pela assinatura do cabeçalho detectado) para as próximas planilhas iguais
                for d in planilha.analisar(os.path.join(SESSAO, card, _nome_seguro(b['arquivo'])), aj[chave_aj], servicos=(card in comparador.SERV_CARDS), locacoes=(card == 'planilha_locacoes')):
                    if d['aba'] == b['aba'] and d.get('assinatura'):
                        planilha.salvar_mapa(d['assinatura'], d['linha_cab'], d['mapa'])
                return self._env(200, analisar_planilhas())
            if u.path == '/api/scp/importar-meses':
                res = _json(os.path.join(SAIDA, 'resultado.json'), None)
                if not res:
                    return self._env(404, {'erro': 'Calcule antes.'})
                if not adm_ok(self):
                    return self._env(403, {'erro': 'Só o ADM altera a Estratégia SCP. Entre no modo ADM.'})
                import base64
                from . import importar_meses
                b = json.loads(self._corpo() or b'{}')
                comp = res.get('competencia') or ''
                mes = int(comp[5:7]) if len(comp) >= 7 else 0
                ini = mes - (mes - 1) % 3
                validos = {'%s-%02d' % (comp[:4], m) for m in range(ini, mes)} if mes else set()
                salvo = _json(SCP_CFG, {})
                cfg_atual = (salvo.get('cfg') or {}) if (salvo.get('cnpj'), salvo.get('competencia')) == ((res.get('empresa') or {}).get('cnpj'), comp) else {}
                lista = {x.get('competencia'): dict(x) for x in (cfg_atual.get('tri_outros_meses') or []) if x.get('competencia')}
                importados, ignorados = [], []
                for a in b.get('arquivos') or []:
                    try:
                        r = importar_meses.ler(base64.b64decode(a.get('b64') or ''), a.get('nome') or '')
                    except Exception as e:
                        ignorados.append({'arquivo': a.get('nome'), 'motivo': 'não consegui ler a planilha (%s)' % e}); continue
                    if not r['competencia']:
                        ignorados.append({'arquivo': r['arquivo'], 'motivo': '; '.join(r['avisos']) or 'mês não encontrado'}); continue
                    if r['competencia'] not in validos:
                        ignorados.append({'arquivo': r['arquivo'], 'motivo': '%s não é mês anterior do trimestre de %s' % (r['competencia'], comp)}); continue
                    if r['receita'] is None:
                        ignorados.append({'arquivo': r['arquivo'], 'motivo': '; '.join(r['avisos'])}); continue
                    lista[r['competencia']] = dict(lista.get(r['competencia']) or {}, competencia=r['competencia'], receita=round(r['receita'], 2),
                                                   receita_matriz=round(r['receita_matriz'] or 0, 2), irrf=r['irrf'], csll=r['csll'], auto=False, fonte=r['arquivo'])
                    lista[r['competencia']].setdefault('lucro', 0.0); lista[r['competencia']].setdefault('irrf_m', 0.0); lista[r['competencia']].setdefault('csll_m', 0.0)
                    for kk in ('carga_est', 'carga_comp', 'economia', 'honorario'):
                        if r.get(kk) is not None:
                            lista[r['competencia']][kk] = r[kk]
                    importados.append(dict(r))
                if not importados:
                    return self._env(200, {'importados': [], 'ignorados': ignorados, 'calc': calcular_scp(res)})
                nova = [lista[k] for k in sorted(lista)]
                return self._env(200, {'importados': importados, 'ignorados': ignorados, 'calc': calcular_scp(res, {'tri_outros_meses': nova, 'tri_manual': True})})
            if u.path == '/api/canceladas':
                b = json.loads(self._corpo() or b'{}')
                marc = _json(CANCELADAS, {})
                if b.get('cancelada') is None:
                    marc.pop(b.get('chave', ''), None)
                else:
                    marc[b.get('chave', '')] = bool(b['cancelada'])
                os.makedirs(SESSAO, exist_ok=True)
                json.dump(marc, open(CANCELADAS, 'w', encoding='utf-8'), ensure_ascii=False)
                return self._env(200, {'canceladas': marc})
            if u.path == '/api/manuais':
                b = json.loads(self._corpo() or b'{}')
                notas = [x for x in (b.get('notas') or []) if isinstance(x, dict)]
                os.makedirs(SESSAO, exist_ok=True)
                json.dump(notas, open(MANUAIS, 'w', encoding='utf-8'), ensure_ascii=False)
                return self._env(200, {'notas': notas})
            if u.path == '/api/calcular':
                b = json.loads(self._corpo() or b'{}')
                cfg = {'cnpj': re.sub(r'\D', '', b.get('cnpj', '')), 'competencia': b.get('competencia', '')}
                os.makedirs(SESSAO, exist_ok=True)
                json.dump(cfg, open(CFG, 'w'), ensure_ascii=False)
                res = comparador.executar(SESSAO, cfg['cnpj'], cfg['competencia'], _json(AJUSTES, {}))
                com_doc = [n for n in res.get('notas', []) if n.get('arquivo')]
                de_fora = [n for n in com_doc if any(a.get('cod') == 'OUTRA_EMPRESA' for a in n.get('aponts', []))]
                if EMPRESA_OK and (not empresa_permitida((res.get('empresa') or {}).get('cnpj')) or (com_doc and len(de_fora) * 2 > len(com_doc))):
                    nome_ = (res.get('empresa') or {}).get('nome') if not empresa_permitida((res.get('empresa') or {}).get('cnpj')) else ''
                    return self._env(403, {'erro': 'Este acesso é só para a empresa de CNPJ %s. A maior parte dos arquivos enviados é de outra empresa%s. Tire esses arquivos e calcule de novo.' % (EMPRESA_OK, (' (%s)' % nome_) if nome_ else '')})
                os.makedirs(SAIDA, exist_ok=True)
                sa = saldo_anterior(res)
                if sa and res.get('apuracao'):
                    res['apuracao']['saldo_anterior'] = sa
                json.dump(res, open(os.path.join(SAIDA, 'resultado.json'), 'w', encoding='utf-8'), ensure_ascii=False, default=list)
                meta = arquivar(json.loads(json.dumps(res, default=list)))
                if meta:
                    json.dump(meta, open(os.path.join(SESSAO, 'meta.json'), 'w', encoding='utf-8'), ensure_ascii=False)
                return self._env(200, json.loads(json.dumps(res, ensure_ascii=False, default=list)))
            return self._env(404, {'erro': 'não encontrado'})
        except Exception as e:
            traceback.print_exc()
            return self._env(500, {'erro': '%s: %s' % (type(e).__name__, e)})


def rodar(porta=8766, abrir=True):
    os.makedirs(SESSAO, exist_ok=True)
    porta = int(os.environ.get('PORT') or porta)
    host = os.environ.get('REVISOR_HOST') or ('0.0.0.0' if os.environ.get('RENDER') else '127.0.0.1')
    if host != '127.0.0.1' and not LOGIN_ON:
        raise SystemExit('Servidor aberto na rede exige REVISOR_USUARIO e REVISOR_SENHA.')
    srv = ThreadingHTTPServer((host, porta), H)
    url = 'http://localhost:%d/' % porta
    print('Revisor do Fechamento em', url, '(Ctrl+C para encerrar)')
    if abrir:
        import webbrowser
        webbrowser.open(url)
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass
