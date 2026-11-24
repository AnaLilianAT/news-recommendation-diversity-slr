# Scripts de análise

## Arquivos
- `gerar_graficos.py`: lê o formulário de extração e a exportação do Parsifal e gera as figuras (PNG em 300 dpi e PDF), organizadas por questão de pesquisa, e as tabelas de contagem em CSV (subpasta `tabelas/`), incluindo os motivos de exclusão por etapa.
- `analise_sensibilidade.py`: compara a distribuição das categorias dos principais campos do formulário no conjunto completo de estudos incluídos e no subconjunto de qualidade moderada ou alta (Q ≥ 0,50), conforme previsto no protocolo.
- `requirements.txt`: dependências.

## Entradas
- `quality_and_data_extraction/Qualidade_e_extracao_de_dados.xlsx`
  - **Qualidade**: respostas às questões de avaliação da qualidade e escore de cada estudo.
  - **Extração de dados**: formulário com as categorias controladas.
  - **Livro de códigos**: definição e critério de uso de cada categoria.
  - **Taxa de citações por ano**: taxa média de citações por ano de cada estudo e seu percentil.
- `selection/artigos aceitos etapa final.xls`: exportação do Parsifal.

## Como executar
Ver o `README.md` na raiz do repositório. Os dois scripts recebem os caminhos dos arquivos de entrada e da pasta de saída como argumentos:

```bash
python scripts/gerar_graficos.py <formulario.xlsx> <parsifal.xls> <pasta_saida>
python scripts/analise_sensibilidade.py <formulario.xlsx> <pasta_saida>
```

Defina `SEM_TITULOS=1` para gerar as figuras sem título.
