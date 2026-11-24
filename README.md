# news-recommendation-diversity-slr

Supplementary materials and reproducibility resources for a systematic literature review on diversification in news recommender systems, including the review protocol, search results, selection decisions, quality assessment, data extraction and analysis scripts.

Materiais suplementares da Revisão Sistemática da Literatura (RSL) **"Diversificação em Sistemas de Recomendação de Notícias"**, desenvolvida como Trabalho de Conclusão de Curso em Ciência da Computação na Universidade Federal de Santa Maria (UFSM).

- **Autora:** Ana Lilian Alfonso Toledo
- **Orientador:** Gabriel Machado Lunardi

## Estrutura do repositório

| Pasta / arquivo | Conteúdo |
|---|---|
| `protocol/protocol.pdf` | Protocolo da RSL, versão 2.0 (15/10/2026), com histórico de emendas e registro de desvios na execução. |
| `search/acm.bib`, `ieee.bib`, `scopus.bib`, `wos.bib` | Exportações bibliográficas das quatro bases (buscas de 22/09/2026 e atualização de 13/11/2026). As *strings* executadas em cada base estão no protocolo. |
| `selection/artigos aceitos etapa final.xls` | Exportação do Parsifal com os 545 registros, a situação de cada um (incluído, excluído ou duplicado) e o critério de exclusão aplicado. |
| `selection/snowballing.csv` | Estudos identificados por *backward snowballing* e avaliados por texto completo, com o critério de exclusão e a justificativa. |
| `quality_and_data_extraction/Qualidade_e_extracao_de_dados.xlsx` | Avaliação da qualidade, taxa de citações por ano, formulário de extração dos 71 estudos incluídos e livro de códigos. |
| `scripts/` | *Scripts* em Python que geram as tabelas, os gráficos e a análise de sensibilidade (ver `scripts/LEIAME.md`). |

## Como reproduzir as análises

Com Python 3.10 ou superior, na raiz do repositório:

```bash
pip install -r scripts/requirements.txt

# Gráficos e tabelas de contagem
python scripts/gerar_graficos.py \
  quality_and_data_extraction/Qualidade_e_extracao_de_dados.xlsx \
  "selection/artigos aceitos etapa final.xls" \
  saida/graficos

# Análise de sensibilidade (exclusão dos estudos de qualidade baixa)
python scripts/analise_sensibilidade.py \
  quality_and_data_extraction/Qualidade_e_extracao_de_dados.xlsx \
  saida/sensibilidade
```

Para gerar as figuras sem título (como no texto do TCC), defina `SEM_TITULOS=1` antes do primeiro comando.

## Observações

- Os PDFs dos estudos não são redistribuídos, por questões de direitos autorais.
- As exportações das bases são mantidas como foram obtidas; a remoção de duplicatas e as decisões de seleção estão registradas na exportação do Parsifal.

## Como citar

TOLEDO, A. L. A.; LUNARDI, G. M. *Protocolo de Revisão Sistemática da Literatura sobre Diversificação em Sistemas de Recomendação de Notícias*. Versão 2.0, 2026. Disponível em: <https://github.com/AnaLilianAT/news-recommendation-diversity-slr>.
