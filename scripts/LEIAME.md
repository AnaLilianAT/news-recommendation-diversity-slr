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
