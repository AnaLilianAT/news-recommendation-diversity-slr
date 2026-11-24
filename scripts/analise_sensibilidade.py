"""Análise de sensibilidade da síntese em relação à qualidade metodológica.

Compara a distribuição das categorias dos principais campos do formulário de extração no
conjunto completo de estudos incluídos com a distribuição no subconjunto de estudos de
qualidade moderada ou alta (Q >= 0,50), conforme previsto no protocolo.

Uso: python analise_sensibilidade.py Qualidade_e_extracao_de_dados.xlsx pasta_saida
"""
import sys, os
import pandas as pd

ARQ, OUT = sys.argv[1], sys.argv[2]
os.makedirs(OUT, exist_ok=True)
LIMIAR_DIFERENCA = 10.0  # pontos percentuais

CAMPOS = [
    ("RQ1", "Paradigma do recomendador base"),
    ("RQ1", "Mecanismo de diversificação (principal)"),
    ("RQ1", "Origem do método de diversificação"),
    ("RQ2", "Etapa da diversificação"),
    ("RQ3", "Aspecto diversificado"),
    ("RQ3", "Representação da notícia"),
    ("RQ3", "Medida de similaridade/distância"),
    ("RQ4", "Origem dos dados"),
    ("RQ4", "Dataset público?"),
    ("RQ5", "Métricas de diversidade"),
    ("RQ6", "Cenário de avaliação"),
    ("RQ6", "Análises realizadas"),
    ("RQ6", "Significância estatística"),
    ("RQ6", "Código disponível?"),
    ("RQ7", "Relação com a diversidade"),
    ("RQ7", "Controle da relação/trade-off"),
    ("RQ8", "Limitações declaradas (categorias)"),
    ("RQ8", "Limitações identificadas na RSL (categorias)"),
]

q = pd.read_excel(ARQ, "Qualidade")
e = pd.read_excel(ARQ, "Extração de dados")
qa = [f"QA{i}" for i in range(1, 11)]

def escore(row):
    vals = []
    for c in qa:
        v = row[c]
        if pd.isna(v) or str(v).strip().upper() in ("N/A", "NA"):
            continue
        vals.append(float(str(v).replace(",", ".")))
    return sum(vals) / len(vals) if vals else float("nan")

q["Q"] = q.apply(escore, axis=1)
q["Faixa"] = pd.cut(q["Q"], [-0.01, 0.4999999, 0.7499999, 1.0], labels=["Baixa", "Moderada", "Alta"])
df = e.merge(q[["Título", "Q", "Faixa"]], on="Título", how="left")
assert df["Q"].notna().all(), "estudo sem escore de qualidade"
sub = df[df["Q"] >= 0.5]
n_all, n_sub = len(df), len(sub)

faixas = q["Faixa"].value_counts().reindex(["Alta", "Moderada", "Baixa"]).fillna(0).astype(int)
faixas.to_frame("estudos").to_csv(os.path.join(OUT, "00_faixas_de_qualidade.csv"), sep=";", encoding="utf-8-sig")
df.loc[df["Q"] < 0.5, ["Título", "Q"]].sort_values("Q").to_csv(
    os.path.join(OUT, "00_estudos_qualidade_baixa.csv"), sep=";", index=False, encoding="utf-8-sig")

def contagem(d, campo):
    s = d[campo].dropna().astype(str).str.split(";").explode().str.strip()
    s = s[s != ""]
    # cada estudo conta uma vez por categoria
    return s.groupby(level=0).apply(lambda x: pd.Series(x.unique())).value_counts() if False else \
        s.reset_index().drop_duplicates().set_index("index")[campo].value_counts()

linhas, resumo = [], []
for rq, campo in CAMPOS:
    c_all, c_sub = contagem(df, campo), contagem(sub, campo)
    cats = sorted(set(c_all.index) | set(c_sub.index), key=lambda c: -c_all.get(c, 0))
    for c in cats:
        pa = 100 * c_all.get(c, 0) / n_all
        ps = 100 * c_sub.get(c, 0) / n_sub
        linhas.append(dict(RQ=rq, Campo=campo, Categoria=c, n_todos=int(c_all.get(c, 0)),
                           pct_todos=round(pa, 1), n_moderada_alta=int(c_sub.get(c, 0)),
                           pct_moderada_alta=round(ps, 1), diferenca_pp=round(ps - pa, 1)))
    top_all = list(c_all.index[:3]); top_sub = list(c_sub.index[:3])
    maior = max((abs(l["diferenca_pp"]) for l in linhas if l["Campo"] == campo), default=0)
    resumo.append(dict(RQ=rq, Campo=campo,
                       categoria_mais_frequente_todos=top_all[0] if top_all else "",
                       categoria_mais_frequente_moderada_alta=top_sub[0] if top_sub else "",
                       mesma_categoria_mais_frequente=(top_all[:1] == top_sub[:1]),
                       mesmas_3_principais=(set(top_all) == set(top_sub)),
                       maior_diferenca_pp=maior,
                       diferenca_acima_limiar=maior >= LIMIAR_DIFERENCA))

pd.DataFrame(linhas).to_csv(os.path.join(OUT, "01_distribuicoes_comparadas.csv"), sep=";", index=False, encoding="utf-8-sig")
r = pd.DataFrame(resumo)
r.to_csv(os.path.join(OUT, "02_resumo_por_campo.csv"), sep=";", index=False, encoding="utf-8-sig")
print(f"Estudos: {n_all}; qualidade moderada ou alta: {n_sub}")
print(faixas.to_dict())
print(r.to_string())
