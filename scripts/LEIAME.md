# Formulário revisado e gráficos — RSL Diversidade em recomendação de notícias

## Arquivos
- `Qualidade_e_extracao_de_dados_revisado.xlsx`
  - **Extração de dados**: formulário com a taxonomia revisada (cabeçalhos em amarelo = campos recodificados ou novos; última coluna = notas da revisão por estudo).
  - **Livro de códigos**: cada categoria com definição e critério de uso.
  - **Mapeamento antigo-novo**: o que aconteceu com cada categoria antiga e por quê.
  - **Log de alterações**: valor original × revisado, por estudo e campo.
  - **Extração (original)**: cópia intacta do formulário anterior.
  - Abas Qualidade e Taxa de citações: inalteradas.
- `gerar_graficos.py`: lê o formulário revisado (sem normalização de texto: as categorias já são controladas).
- `graficos/`: figuras em PNG (300 dpi) e PDF, e `tabelas/` com as contagens em CSV.

## Como regenerar
Coloque o .xlsx revisado e `artigos_aceitos_etapa_final.xls` na pasta do script e rode `python gerar_graficos.py`.

## Observações
- O estudo "The research and implementation of a large-scale real-time news recommendation algorithm" está com a linha vazia no formulário e sem PDF na pasta: entra só no gráfico de publicações por ano (N = 70 nas demais análises).
- "Relevancy and Diversity in News Recommendations" também não tem PDF na pasta; foi recodificado a partir da extração existente.
- As categorias de limitações foram atribuídas por palavras-chave sobre o texto livre e devem ser revisadas manualmente.

## Atualização (11/11/2026)

- Gráficos regenerados com os arquivos corrigidos: 70 estudos incluídos (o estudo de Zhao et al.
  passou a "Rejected" no Parsifal e foi retirado da planilha de extração).
- `02_funil_selecao`: agora segue as etapas do PRISMA 2020 (533 → 206 duplicatas → 327 triados →
  244 excluídos → 83 buscados → 1 não recuperado → 82 avaliados → 12 excluídos → 70 incluídos).
  As quantidades 83 e 82 foram informadas pela autora (constantes no início do script).
- `03_motivos_exclusao`: critérios separados pela data da decisão (até 28/09/2026 = triagem;
  a partir de 01/10/2026 = texto completo). A exclusão sem critério registrado corresponde ao
  texto completo não recuperado (CE12). O 12.º estudo excluído no texto completo não é
  identificável na exportação e permanece no primeiro período (245, e não 244).
- `04_bases_de_dados`: a correspondência entre bases passou a usar título e DOI.
- `graficos_sem_titulo/`: mesmas figuras sem título interno, para uso com \caption no LaTeX
  (gerar com a variável de ambiente SEM_TITULOS=1).

## Atualização (12/11/2026)

- Entrada: nova exportação do Parsifal (CE9 para "A Multiple Perceptual...", CE12 para o estudo
  de Zhao et al.).
- `REDUNDANTES_TEXTO_COMPLETO`: "Incorporating Accuracy and Diversity in a News Recommender
  System" (registro da Scopus) é tratado como exclusão do texto completo por CE11 (publicação
  redundante de "Accuracy meets Diversity..."), e não como duplicata removida antes da triagem.
  Com isso: 205 duplicatas, 328 triados, 245 excluídos na triagem, 12 excluídos no texto completo.
  Se o registro for atualizado no Parsifal para "Rejected"/CE11, a lista pode ser esvaziada.
- Gráficos alterados: `02_funil_selecao` e `03_motivos_exclusao`. Os demais não foram afetados
  e foram mantidos como na versão anterior.

## Atualização (12/11/2026, noite)

- Entrada: exportação do Parsifal com "Incorporating Accuracy..." (Scopus) e "Improving selection
  diversity..." (Scopus) reclassificados como excluídos por CE11; os demais registros dessas
  publicações permanecem como duplicatas (204 duplicatas no total).
- A regra REDUNDANTES_TEXTO_COMPLETO foi removida (o Parsifal já registra a exclusão).
- Nova regra ETAPA_DA_DECISAO: "Improving selection diversity..." é contabilizado na triagem,
  etapa em que a sobreposição foi identificada (a data no Parsifal é a da reclassificação).
- Números: 533 → 204 duplicatas → 329 triados → 246 excluídos na triagem → 83 → 1 não
  recuperado → 82 → 12 excluídos → 70.
- Gráficos alterados: 02_funil_selecao e 03_motivos_exclusao.

## Atualização (13/11/2026)

- Entradas: exportação do Parsifal com a atualização das buscas de 13/11/2026 (545 registros) e
  planilha de qualidade e extração com 71 estudos (inclusão de "A Study of Normative Diversity
  Metrics in News Recommendations" e QA11 recalculada com citações coletadas em 13/11/2026).
- N_BUSCADOS_TEXTO_COMPLETO = 87 e N_AVALIADOS_TEXTO_COMPLETO = 86 (busca original + atualização).
- ETAPA_DA_DECISAO inclui as cinco exclusões da atualização feitas na triagem (triagem e texto
  completo ocorreram no mesmo dia, 13/11/2026, e não podem ser separados pela data).
- Fluxo das bases: 545 → 207 duplicatas → 338 → 251 excluídos na triagem → 87 → 1 não
  recuperado → 86 → 15 excluídos → 71 incluídos. Os três estudos do backward snowballing não
  estão no Parsifal e não aparecem nos gráficos 02 e 03 (todos excluídos no texto completo).
- Todos os gráficos foram regenerados, pois o conjunto de estudos incluídos passou de 70 para 71.
- Nova pasta Sensibilidade/: análise de sensibilidade por qualidade (analise_sensibilidade.py e
  tabelas). Uso: python analise_sensibilidade.py Qualidade_e_extracao_de_dados.xlsx pasta_saida
