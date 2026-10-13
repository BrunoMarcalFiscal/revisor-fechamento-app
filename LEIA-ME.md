# Revisor do Fechamento

Ferramenta local (não ligada ao Cockpit) que confere o lançamento do sistema do cliente contra os documentos fiscais do mês. Pensada para empresas do Lucro Real, com muita entrada e saída.

## Como usar

1. Duplo clique em **Revisor do Fechamento.command** (ou `python3 revisor_fechamento.py`). Abre `http://localhost:8766/`.
2. Na aba **Documentos**, arraste em cada quadro (aceita arquivos soltos, pastas inteiras ou .zip):
   - XML de Entrada, XML de Saída, XML de Frete de Entrada (CT-e), XML de Serviços Prestados, XML de Serviços Tomados;
   - Documentos Diversos: o PDF geral de notas (DANFE, DACTE, NFS-e);
   - **Planilha do sistema** (xlsx, xls ou csv): a cópia da tela do sistema, como a nota foi lançada. É a referência das NF-e e CT-e.
   - **Relatório de serviços** (xlsx, xls ou csv): o relatório do sistema com os serviços prestados e tomados. Confere as NFS-e (valor, base e ISS, alíquota, ISS retido, IRRF, CSLL, PIS, COFINS, INSS, item da LC 116, líquido). Prestado/tomado sai da coluna Tipo (Prestado/Tomado ou S/E) ou do nome da aba.
3. Confira as colunas reconhecidas da planilha. Se faltar alguma, clique em **ajustar colunas**: o layout fica lembrado para a próxima planilha igual.
4. CNPJ e competência são opcionais (identifico pelos XML e pela planilha).
5. **Calcular**. A aba **Resultado** mostra: totais, quadro por movimento (XML × sistema), ranking "onde estão os erros" e a lista de notas. Clique numa nota para ver o que está errado, o valor na nota, o valor no sistema, como corrigir e a comparação linha a linha (CFOP/CST) com a célula errada marcada.
6. **Exportar Excel**: abas Apontamentos (uma linha por correção, com a linha da planilha onde está), Notas e Resumo.

Modo sem tela: `python3 revisor_fechamento.py --lote PASTA` (PASTA com subpastas entrada, saida, frete, prestado, tomado, diversos, planilha). Grava `saida/resultado.json` e `saida/revisao.xlsx`.

Massa de teste: `python3 exemplo/gerar_exemplo.py` cria `exemplo/lote/` (empresa fictícia com um erro plantado por regra).

## O que é conferido

**Presença**: XML sem lançamento; lançamento sem XML (e se consta no PDF); documento do PDF sem lançamento; nota lançada duas vezes; nota cancelada lançada com valor; XML de outra empresa ou de outra competência.

**Cabeçalho**: valor total, CNPJ do participante (e filial), chave, série, data de entrada antes da emissão, data de emissão.

**Entradas (NF-e)**, linha a linha por CFOP/CST:
- CFOP convertido do fornecedor (5→1, 6→2, 7→3; 5405→1403; ST→x403) e aceitos por família (compra para revenda, industrialização, uso e consumo, ativo, transferência, devolução, remessas);
- CFOP estadual × interestadual trocado; compra com ST lançada em CFOP sem ST e o contrário;
- CST de ICMS aceito conforme o CST/CSOSN do fornecedor; origem da mercadoria;
- crédito de ICMS maior que o destacado (ou que o permitido do Simples), crédito em uso e consumo, crédito integral de ativo (CIAP), ICMS destacado sem crédito, base e alíquota;
- ICMS-ST e IPI;
- PIS/COFINS: crédito em monofásico/alíquota zero na revenda, crédito em uso e consumo, compra tributada sem crédito, base do crédito sem excluir o ICMS (Lei 14.592/2023), alíquotas do não cumulativo.

**Saídas (NF-e)**: CFOP, CST, base, alíquota, ICMS, ST, IPI, PIS/COFINS e CST de PIS/COFINS iguais ao XML.

**Frete (CT-e)**: empresa é a tomadora? CFOP 1.35x/2.35x pela UF de início, crédito de ICMS × destacado, frete sem crédito de PIS/COFINS.

**Serviços (NFS-e nacional e ABRASF)**: valor, ISS, alíquota, ISS retido (marcado ou não), IRRF, CSLL, INSS, PIS e COFINS retidos.

Cada conferência só roda se a planilha tiver a coluna correspondente; o painel avisa quais faltaram.

## Aba 3: Apuração ICMS/IPI

Fechamento a partir da planilha do sistema (o livro como foi lançado):
- **Resumo por CFOP** com Valor contábil, Base ICMS, Valor ICMS, Base IPI, Valor IPI, Isentas e Outras; clique no CFOP para abrir os CST. Quando a planilha não tem CST, Base IPI, Isentas ou Outras, o valor vem do XML casado ou é calculado, e a tela avisa.
- **Ajustes** (cada um com valor, descrição e código SPED E111): crédito do Simples Nacional (pCredSN do XML ou texto do rodapé; fora do ajuste quando já creditado na nota ou lançado em uso e consumo/ativo/retorno), DIFAL de uso e consumo (2556, por dentro ou por fora, alíquota interna editável), estorno de débito e estorno de crédito (com sugestões tiradas dos apontamentos).
- **Código E111:** já vem o prefixo da Tabela 5.1.1 (UF + 0 + tipo: 0 outros débitos, 1 estorno de créditos, 2 outros créditos, 3 estorno de débitos); os 4 dígitos finais são da tabela da UF.
- **Resumo:** débitos × créditos de ICMS e de IPI, com saldo credor anterior, e o resultado: imposto a recolher ou saldo credor.
- Os ajustes ficam salvos em `sessao/apuracao.json` e vão para a aba "Apuração ICMS-IPI" do Excel.

## Aba 4: Painel da empresa

Visão gerencial do mês a partir dos XML (documentos reais), com o sistema ao lado onde faz sentido: faturamento bruto (inclui remessas de entrega futura 5116/5117; o 5922 fica fora para não contar duas vezes) − devoluções = líquido, serviços prestados, impostos do período (ICMS/IPI da apuração, destacados nas vendas, ISS, retenções sofridas e a recolher), carga tributária, clientes que mais compram (matriz e filiais juntas), fornecedores de matéria-prima (pelo CFOP lançado), compras por finalidade, fornecedores do Simples, produtos mais vendidos, matérias-primas mais compradas e vendas por destino. Os arquivos não trazem guias pagas: impostos aqui são apurados/destacados. Vai para a aba "Painel da empresa" do Excel.

## Aba 5: Visão estratégica (PDF e Excel)

Evolução mês a mês (dos fechamentos guardados), margem bruta estimada, saldo credor e o que fazer com ele, créditos não aproveitados com valor, projeção do Lucro Real (PIS/COFINS não cumulativo e estimativa mensal de IRPJ/CSLL), simulação da reforma em 2027 (CBS/IBS com alíquota de referência editável), concentração e risco, e indicadores do fechamento. Cada bloco traz as premissas.

## Relatório de locações

Quadro próprio na aba Documentos para o relatório de locações tirado do sistema do cliente (Excel/CSV). Colunas reconhecidas sozinhas (locatário, CNPJ/CPF, valor, data/vencimento, contrato/recibo, bem, tipo receita/despesa) e ajustáveis; o layout fica lembrado. Receita de locação entra no faturamento, nos relatórios (receita Locações) e no ranking de clientes; despesa de locação de PJ entra como crédito na projeção de PIS/COFINS. Avisa valor zerado, linha sem locatário e lançamento repetido.

## Aba 6: Estratégia SCP (Matriz Lucro Real × SCP Lucro Presumido)

Motor em `motor/scp.py` (extrair do fechamento + calcular com memória), parâmetros em `sessao/scp.json` (não alteram os dados importados).
- Produtos ficam na Matriz. Débito = valor contábil − ICMS (Tema 69) − IPI e ICMS-ST (DL 1.598/77, art. 12, §4º) − devoluções de venda.
- Créditos da Matriz por CFOP marcável (padrão: insumo, revenda, energia, frete), base = valor contábil − ICMS (Lei 14.592/2023) − IPI recuperável; devolução de compra reduz; importação pelo PIS/COFINS-Importação pago; serviços tomados de PJ; aluguéis PJ; outros créditos informados; saldo credor anterior (automático do mês anterior ou digitado).
- Modo automático: a Matriz fica com serviço/locação só até os créditos zerarem o PIS/COFINS; o resto vai para a SCP (cumulativo 0,65%/3%; IRPJ/CSLL 32%, trimestral com adicional acima de R$ 60 mil). Modo manual: proporção escolhida.
- PIS e COFINS apurados separados: créditos do período, saldo anterior, créditos usados, retenção (só deduz o devido; o que sobra é restituição, não crédito) e saldo credor a transportar.
- Comparativo do mês: estratégia × tudo no Lucro Real × tudo no Presumido; alertas de risco (documentação da SCP, crédito de insumo de receita da SCP, bens da locação, ISS em locação, limite do Presumido).
- Conferência com a planilha do escritório: `python3 exemplo/validar_planilha_junho.py "<planilha>"` reproduz a planilha de fechamento de um mês centavo a centavo (o script fica só no Mac, fora do repositório).

## Relatórios de serviços (PDF ou Excel)

Dois quadros: Saída de serviços (prestados) e Entrada de serviços (tomados). O PDF é lido pela posição do texto (precisa ter texto, não escaneado). Campos padronizados: número, data, participante, CNPJ/CPF, código do serviço, valor, base e valor do ISS, ISS retido, IRRF, CSLL, PIS, COFINS, INSS e líquido.

## Aba 7: Relatórios (PDF e Excel)

- **Demonstrativo por imposto:** resumo (ICMS próprio, IPI, ICMS-ST, DIFAL de venda a não contribuinte, ISS prestado e retido, IRRF, CSRF, INSS) e o detalhe de cada um.
- **Entradas e saídas:** consolidado ou por receita (Produtos, Serviços, Locações), entradas, saídas ou ambos; sem canceladas; mostra se a nota está no sistema.
- **Clientes e fornecedores:** ranking por valor, produtos × serviços, regime, impostos e retenções.
Locação é identificada pela descrição/item da NFS-e (locação, aluguel, cessão de uso, item 03) ou pelo CFOP x908/x909 da NF-e.

## Rotina

- **Fechamentos:** cada Calcular guarda o fechamento em `fechamentos/<cnpj>_<aaaa-mm>/` (arquivos, ajustes, correções e resultado). "Novo fechamento" guarda antes de limpar; o botão Fechamentos lista e reabre. A situação (em revisão, corrigido, fechado) e o % resolvido ficam no topo.
- **Saldo credor anterior** de ICMS e IPI vem do fechamento do mês anterior da mesma empresa.
- **Conferência das colunas:** a planilha mostra a linha de prova (como uma linha real foi lida) e alerta coluna zerada, com texto ou número vazio; com alerta grave, o Calcular pede confirmação.
- **Correção em lote:** filtrando um tipo de apontamento, dá para corrigir ou justificar todos de uma vez.
- **Pendências do cliente:** a aba Resultado monta a lista do que pedir (XML de CT-e, de meses anteriores, eventos de cancelamento, números pulados, relatório de serviços) e copia a mensagem pronta.

## Como os apontamentos são organizados

- **Um apontamento por causa.** A mesma regra em várias linhas da nota vira um só (valores somados); o que é consequência de outro erro (total diferente porque faltou uma linha, base diferente porque o ICMS está errado) aparece como "Também:" dentro do erro de origem.
- **Categoria** em cada apontamento (Lançamento, Identificação, Valores, CFOP, CST, ICMS e crédito, ST, IPI, PIS/COFINS, Frete, ISS, Retenções, Devolução, Cancelamento e numeração); o ranking e o filtro seguem as categorias.
- **Por que é erro:** fundamento curto de cada regra.
- **Pontos conferidos:** cada nota mostra o que foi conferido (✓), o que tem erro (✗) ou é para conferir (!), e o que não pôde ser conferido porque a planilha não tem a coluna (–).
- **Corrigido / Justificar:** cada apontamento pode ser marcado como corrigido ou justificado (com texto obrigatório). A nota também aceita uma observação livre. Tudo fica salvo em `sessao/corrigidos.json` e vai para o Excel (Resolução, Justificativa, Observação da nota, Pontos não conferidos).

## Estrutura

```
revisor_fechamento.py        inicia a tela / modo --lote
motor/xml_docs.py            NF-e, CT-e, NFS-e (nacional e ABRASF, lote), eventos de cancelamento
motor/pdf_docs.py            PDF: chave de acesso (com dígito verificador) e NFS-e por número/CNPJ
motor/planilha.py            leitura xlsx/xls/csv, cabeçalho, sinônimos de colunas, linhas de continuação
motor/regras_base.py         CFOP de entrada esperado, CST aceitos, grupos de CFOP, PIS/COFINS
motor/comparador.py          casamento nota × lançamento e todas as regras (catálogo CATALOGO)
motor/consolidar.py          um apontamento por causa, categorias, fundamento, pontos conferidos
motor/devolucoes.py          devolução × nota de origem, canceladas, inutilizadas, numeração
motor/exportar.py            Excel
motor/servidor.py            servidor local (porta 8766)
web/index.html               tela
sessao/                      arquivos do fechamento em andamento     saida/  resultado
config/mapeamentos.json      layouts de planilha lembrados
```

## Limites conhecidos

- O destino da mercadoria (revenda, consumo, ativo) não está no XML: o revisor aceita os CFOP da família e aponta só o que é incompatível; crédito não tomado sai como "Conferir".
- PDF escaneado (sem texto) não é lido; o painel lista as páginas.
- Validado com massa sintética. Na primeira planilha real de cada cliente, conferir as colunas reconhecidas antes de calcular.
