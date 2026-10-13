# -*- coding: utf-8 -*-
"""Revisor do Fechamento: confere os XML (entradas, saídas, frete, serviços) e o PDF de notas contra a planilha do sistema do cliente.
Uso: python3 revisor_fechamento.py            abre a tela em http://localhost:8766/
     python3 revisor_fechamento.py --lote DIR  roda sem tela sobre DIR/<card>/ e grava saida/"""
import sys, os, json
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from motor import servidor, comparador, exportar

if __name__ == '__main__':
    if '--lote' in sys.argv:
        pasta = sys.argv[sys.argv.index('--lote') + 1]
        res = comparador.executar(pasta)
        os.makedirs('saida', exist_ok=True)
        json.dump(res, open('saida/resultado.json', 'w', encoding='utf-8'), ensure_ascii=False, default=list)
        open('saida/revisao.xlsx', 'wb').write(exportar.xlsx(json.loads(json.dumps(res, default=list))))
        print(res['resumo'])
    else:
        servidor.rodar(abrir='--sem-navegador' not in sys.argv)
