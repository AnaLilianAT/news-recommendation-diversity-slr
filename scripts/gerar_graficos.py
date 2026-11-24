# -*- coding: utf-8 -*-
"""
Geração dos gráficos da RSL "Diversidade em sistemas de recomendação de notícias"
— versão para o formulário com a TAXONOMIA REVISADA.

Entradas (mesma pasta do script, ou passe os caminhos como argumentos):
  1. Qualidade_e_extracao_de_dados_revisado.xlsx -> abas "Qualidade", "Taxa de citações por ano",
     "Extração de dados" (categorias conforme a aba "Livro de códigos")
  2. artigos_aceitos_etapa_final.xls -> exportação da ferramenta de RSL (ano, base, status, critério)
     Estudos incluídos = registros com status "Accepted".

Uso:
  python gerar_graficos.py [planilha_extracao.xlsx] [artigos.xls] [pasta_saida]

Saída: uma subpasta por questão de pesquisa com cada figura em PNG (300 dpi) e PDF
(vetorial, para LaTeX), e uma pasta "tabelas" com as contagens de cada figura em CSV.

Observações metodológicas
- Como as categorias já estão padronizadas no formulário, o script apenas separa os
  valores por ";" — não há mais normalização de texto aqui.
- Estudos sem extração (linha vazia no formulário) ficam fora das análises de conteúdo,
  mas entram no gráfico de publicações por ano.
- Campos multivalorados: contagens de ESTUDOS por categoria; percentuais sobre N
  (as somas podem passar de 100 %). O mecanismo principal é univalorado (soma 100 %).
- Categorias com menos de MIN_FREQ estudos são agrupadas em "Outros".
"""
import os
import re
import sys
import unicodedata
from collections import Counter, OrderedDict

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
import pandas as pd
from matplotlib.colors import LinearSegmentedColormap

# --------------------------------------------------------------------------------------
# Configuração
# --------------------------------------------------------------------------------------
BASE = os.path.dirname(os.path.abspath(__file__))
ARQ_EXTRACAO = sys.argv[1] if len(sys.argv) > 1 else os.path.join(BASE, "Qualidade_e_extracao_de_dados_revisado.xlsx")
ARQ_ARTIGOS = sys.argv[2] if len(sys.argv) > 2 else os.path.join(BASE, "artigos_aceitos_etapa_final.xls")
SAIDA = sys.argv[3] if len(sys.argv) > 3 else os.path.join(BASE, "graficos")

MIN_FREQ = 2
# Defina a variável de ambiente SEM_TITULOS=1 para gerar figuras sem título (para uso com \caption no LaTeX)
MOSTRAR_TITULOS = not os.environ.get("SEM_TITULOS")
FORMATOS = ("png", "pdf")
DPI = 300
ANO_PARCIAL = 2026
# Etapas da seleção que não podem ser derivadas apenas da exportação do Parsifal
# (informadas pela autora): relatos buscados para o texto completo e relatos avaliados.
N_BUSCADOS_TEXTO_COMPLETO = 87  # 83 da busca original + 4 da atualização de 13/11/2026
N_AVALIADOS_TEXTO_COMPLETO = 86  # 82 da busca original + 4 da atualização
# Decisões de exclusão registradas a partir desta data correspondem à avaliação do texto
# completo (a triagem por título e resumo foi concluída em 28/09/2026).
INICIO_TEXTO_COMPLETO = "2026-10-01"
# Exclusões cuja etapa não corresponde à data da última atualização no Parsifal (o status foi
# reclassificado depois da etapa em que a decisão foi tomada). Formato: (título, base, etapa).
ETAPA_DA_DECISAO = [
    # Sobreposição substancial com "Improving consumption diversity via graph-based topic nudging"
    # (CE11), identificada durante a triagem por título e resumo; reclassificada em 12/11/2026.
    ("Improving selection diversity using hybrid graph-based news recommenders", "Scopus", "Triagem"),
    # Atualização da busca (13/11/2026): exclusões realizadas na triagem por título e resumo
    # (triagem e avaliação do texto completo ocorreram no mesmo dia).
    ("A Knowledge Graph-Driven Framework For Precise Cross-Media NewsRecommendation And Coordinated Public Opinion Guidance", "Web of Science", "Triagem"),
    ("SentiGraphRec: sentiment perception entity graph for personalized news recommendation", "Web of Science", "Triagem"),
    ("Towards an Adaptive Turn in Journalism: Conceptualizing the Algorithmic Co-Constitution of the News Media-Audience Relationship", "Web of Science", "Triagem"),
    ("The algorithmic perception paradox: personalization, news diversity, and trust among young UAE news users", "Web of Science", "Triagem"),
    ("Do We Care About Personalization and Explainability? An Interview Study with News Recommendation Engineers", "ACM Digital Library", "Triagem"),
]
PERIODOS = [(2009, 2017, "2009–2017"), (2018, 2021, "2018–2021"),
            (2022, 2024, "2022–2024"), (2025, 2026, "2025–2026")]

AZUL = "#2a78d6"
CATEG = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#6250d6", "#e34948"]
CINZA = "#a8a7a2"
TXT1, TXT2, GRADE = "#0b0b0b", "#52514e", "#e6e5e0"
SEQ = LinearSegmentedColormap.from_list(
    "azul", ["#f4f8fd", "#cde2fb", "#9ec5f4", "#6da7ec", "#3987e5", "#256abf", "#184f95", "#0d366b"])

plt.rcParams.update({
    "font.family": "DejaVu Sans", "font.size": 10, "axes.titlesize": 12,
    "axes.titleweight": "bold", "axes.labelsize": 10, "axes.labelcolor": TXT2,
    "xtick.color": TXT2, "ytick.color": TXT1, "axes.edgecolor": "#c3c2b7",
    "axes.spines.top": False, "axes.spines.right": False, "figure.dpi": 100,
    "savefig.bbox": "tight", "savefig.pad_inches": 0.15, "pdf.fonttype": 42,
})

# Rótulos curtos para gráficos (os rótulos completos estão no livro de códigos)
CURTO = {
    "Seleção de subconjunto por dissimilaridade (MMR, dispersão, DPP)": "Dissimilaridade (MMR, dispersão, DPP)",
    "Maximização de cobertura (funções submodulares)": "Cobertura submodular",
    "Calibração, cotas e restrições de distribuição": "Calibração/cotas/restrições",
    "Objetivo de diversidade na função de perda": "Diversidade na função de perda",
    "Regularização de representações": "Regularização de representações",
    "Recompensa de diversidade (aprendizado por reforço/bandits)": "Recompensa de diversidade (RL/bandits)",
    "Ponderação de escores com componente de diversidade": "Ponderação de escores",
    "Hibridização de recomendadores (conteúdo + colaborativo)": "Hibridização CB + CF",
    "Expansão ou reorientação do perfil do usuário": "Expansão/reorientação do perfil",
    "Agrupamento ou passeios aleatórios em grafo": "Agrupamento/passeios em grafo",
    "Intervenção comportamental (nudging)": "Nudging",
    "Exploração/aleatorização": "Exploração/aleatorização",
    "Dissimilaridade intralista (ILD/ILS/ILAD)": "Dissimilaridade intralista (ILD)",
    "Uniformidade da distribuição (entropia, Gini)": "Uniformidade (entropia, Gini)",
    "Divergência em relação a distribuição de referência": "Divergência (KL, Wasserstein)",
    "Diversidade de consumo/exposição observada": "Consumo/exposição observada",
    "Diversidade percebida (questionário)": "Diversidade percebida",
    "Métricas normativas (RADio e afins)": "Normativas (RADio)",
    "Cobertura/contagem de categorias": "Cobertura de categorias",
    "Distância ao histórico do usuário": "Distância ao histórico",
    "Diversidade entre listas/usuários": "Entre listas/usuários",
    "Ganho de diversidade sem perda significativa": "Sem perda significativa",
    "Parâmetro explícito com análise de sensibilidade": "Parâmetro com análise de sensibilidade",
    "Parâmetro explícito sem análise sistemática": "Parâmetro sem análise sistemática",
    "Correspondência por similaridade de conteúdo": "Similaridade de conteúdo",
    "Fatoração de matrizes/fatores latentes": "Fatoração de matrizes",
    "Entidades/conceitos (grafo, ontologia, taxonomia)": "Entidades/conceitos",
    "Conteúdo textual (forma não especificada)": "Texto (forma não especificada)",
    "Dados de experimento próprio com usuários": "Experimento próprio com usuários",
    "Classificação/conjunto (AUC, precisão, revocação, F1)": "Classificação (AUC, P, R, F1)",
    "Comportamento observado (CTR, cliques, tempo de leitura)": "Comportamento (CTR, cliques, dwell)",
    "Beyond-accuracy (novidade, serendipidade, cobertura)": "Beyond-accuracy",
    "Experimento online/de campo (A/B)": "Experimento online/campo (A/B)",
    "Estudo com usuários (controlado)": "Estudo com usuários",
    "Offline (dados históricos)": "Offline",
}
ORDEM_ETAPAS = ["Modelagem do usuário", "Geração/recuperação de candidatos", "Treinamento do modelo",
                "Pontuação/ranqueamento integrado", "Pós-processamento (re-ranking)",
                "Interface/comportamento do usuário"]


# --------------------------------------------------------------------------------------
# Leitura e utilidades
# --------------------------------------------------------------------------------------
def chave(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]", "", s.lower())


def dividir(valor):
    if pd.isna(valor):
        return []
    return [p.strip() for p in str(valor).split(";") if p.strip() and p.strip() != "—"]


def rotulos(valor, curto=True):
    out = []
    for t in dividir(valor):
        t = CURTO.get(t, t) if curto else t
        if t not in out:
            out.append(t)
    return out


def carregar():
    ext = pd.read_excel(ARQ_EXTRACAO, sheet_name="Extração de dados")
    ext.columns = [c.strip() for c in ext.columns]
    qual = pd.read_excel(ARQ_EXTRACAO, sheet_name="Qualidade")
    cit = pd.read_excel(ARQ_EXTRACAO, sheet_name="Taxa de citações por ano")
    art = pd.read_excel(ARQ_ARTIGOS)
    ext["k"] = ext["Título"].map(chave)
    qual["k"] = qual["Título"].map(chave)
    cit["k"] = cit["Título"].map(chave)
    art["k"] = art["title"].map(chave)
    art["doi_n"] = art["doi"].astype(str).str.lower().str.strip().where(art["doi"].notna())
    art["atualizado"] = pd.to_datetime(art["updated_at"], unit="D", origin="1899-12-30")
    aceitos = art[art["status"] == "Accepted"].drop_duplicates("k")
    df = ext.merge(aceitos[["k", "year", "source"]], on="k", how="left")
    df = df.merge(qual.drop(columns="Título"), on="k", how="left")
    df = df.merge(cit.drop(columns="Título"), on="k", how="left")
    df = df.rename(columns={"year": "Ano", "source": "Base de origem"})
    for c in df.columns:
        if str(c).startswith("Taxa média") or str(c).startswith("Percentil"):
            df[c] = pd.to_numeric(df[c].astype(str).str.extract(r"^\s*(-?\d+(?:[.,]\d+)?)")[0]
                                  .str.replace(",", "."), errors="coerce")
    df["Ano"] = df["Ano"].astype("Int64")
    df["codificado"] = df["Mecanismo de diversificação (principal)"].notna()
    sem = df.loc[~df["codificado"], "Título"].tolist()
    if sem:
        print("ATENÇÃO: estudos sem extração (fora das análises de conteúdo):", sem)
    return df, art


def lista(df, coluna):
    return df[coluna].map(rotulos)


def contar_listas(listas, min_freq=MIN_FREQ, outros=True):
    c = Counter(r for lst in listas for r in set(lst))
    if outros and min_freq > 1:
        raros = {k for k, v in c.items() if v < min_freq}
        if len(raros) > 1:
            n_outros = sum(1 for lst in listas if set(lst) & raros)
            c = Counter({k: v for k, v in c.items() if k not in raros})
            c["Outros"] = c.get("Outros", 0) + n_outros
    s = pd.Series(c, dtype=int).sort_values(ascending=False)
    if "Outros" in s.index:
        s = pd.concat([s.drop("Outros"), s[["Outros"]]])
    return s


# --------------------------------------------------------------------------------------
# Contagens
# --------------------------------------------------------------------------------------
def agrupar_raros(listas, min_freq=MIN_FREQ):
    c = Counter(r for lst in listas for r in set(lst))
    raros = {k for k, v in c.items() if v < min_freq}
    if len(raros) <= 1:
        return listas
    return listas.map(lambda lst: list(OrderedDict.fromkeys("Outros" if r in raros else r for r in lst)))


# --------------------------------------------------------------------------------------
# Saída
# --------------------------------------------------------------------------------------
def salvar(fig, pasta, nome, tabela=None):
    d = os.path.join(SAIDA, pasta)
    os.makedirs(d, exist_ok=True)
    for ext in FORMATOS:
        fig.savefig(os.path.join(d, f"{nome}.{ext}"), dpi=DPI)
    plt.close(fig)
    if tabela is not None:
        dt = os.path.join(SAIDA, "tabelas")
        os.makedirs(dt, exist_ok=True)
        tabela.to_csv(os.path.join(dt, f"{nome}.csv"), encoding="utf-8-sig", sep=";")
    print("ok:", pasta, nome)


def titulo(ax, texto, sub=None):
    if MOSTRAR_TITULOS:
        ax.set_title(texto, loc="left", color=TXT1, pad=22 if sub else 10)
        if sub:
            ax.text(0, 1.015, sub, transform=ax.transAxes, fontsize=9, color=TXT2, va="bottom")


def barras_h(contagem, n_total, titulo_txt, pasta, nome, sub=None, xlabel="Número de estudos",
             cor=AZUL, cores=None, destaque=None, largura=8.0):
    if "Outros" in contagem.index:
        contagem = pd.concat([contagem.drop("Outros"), contagem[["Outros"]]])
    s = contagem.iloc[::-1]
    altura = max(2.2, 0.36 * len(s) + 1.3)
    fig, ax = plt.subplots(figsize=(largura, altura))
    if cores is None:
        cores = [CINZA if (k == "Outros" or (destaque and k in destaque)) else cor for k in s.index]
    ax.barh(s.index, s.values, color=cores, height=0.68, edgecolor="white", linewidth=1)
    xmax = s.max()
    for i, v in enumerate(s.values):
        ax.text(v + xmax * 0.012, i, f"{v} ({100 * v / n_total:.0f}%)", va="center",
                fontsize=9, color=TXT2)
    ax.set_xlim(0, xmax * 1.2)
    ax.set_xlabel(xlabel)
    ax.xaxis.set_major_locator(mticker.MaxNLocator(integer=True))
    ax.grid(axis="x", color=GRADE, linewidth=0.8)
    ax.set_axisbelow(True)
    ax.tick_params(axis="y", length=0)
    ax.spines["left"].set_visible(False)
    titulo(ax, titulo_txt, sub)
    tab = pd.DataFrame({"estudos": contagem.values,
                        "percentual": (100 * contagem.values / n_total).round(1)}, index=contagem.index)
    salvar(fig, pasta, nome, tab)


def heatmap(df_lista_a, df_lista_b, titulo_txt, pasta, nome, rotulo_a, rotulo_b, sub=None,
            ordem_a=None, ordem_b=None, normalizar_linha=False, largura=None, quebra_b=18, rot=45):
    """Matriz de coocorrência (nº de estudos) entre dois campos multivalorados."""
    pares = Counter()
    for la, lb in zip(df_lista_a, df_lista_b):
        for a in set(la):
            for b in set(lb):
                pares[(a, b)] += 1
    ca = Counter(a for la in df_lista_a for a in set(la))
    cb = Counter(b for lb in df_lista_b for b in set(lb))
    ordem_a = ordem_a or [k for k, _ in sorted(ca.items(), key=lambda x: (x[0] == "Outros", -x[1]))]
    ordem_b = ordem_b or [k for k, _ in sorted(cb.items(), key=lambda x: (x[0] == "Outros", -x[1]))]
    m = pd.DataFrame([[pares.get((a, b), 0) for b in ordem_b] for a in ordem_a],
                     index=ordem_a, columns=ordem_b)
    valores = m.div(pd.Series(ca)[ordem_a], axis=0) * 100 if normalizar_linha else m
    w = largura or max(6.5, 0.72 * len(ordem_b) + 3.8)
    h = max(3.0, 0.42 * len(ordem_a) + 2.2)
    fig, ax = plt.subplots(figsize=(w, h))
    im = ax.imshow(valores.values, cmap=SEQ, aspect="auto", vmin=0)
    vmax = valores.values.max()
    for i in range(len(ordem_a)):
        for j in range(len(ordem_b)):
            n = m.values[i, j]
            if n == 0:
                continue
            txt = f"{valores.values[i, j]:.0f}%" if normalizar_linha else str(n)
            ax.text(j, i, txt, ha="center", va="center", fontsize=8.5,
                    color="white" if valores.values[i, j] > vmax * 0.55 else TXT1)
    ax.set_xticks(range(len(ordem_b)))
    ax.set_xticklabels([quebrar(b, quebra_b) for b in ordem_b], rotation=rot,
                       ha="right" if rot else "center", rotation_mode="anchor", fontsize=9)
    ax.set_yticks(range(len(ordem_a)))
    ax.set_yticklabels([f"{a} (n={ca[a]})" for a in ordem_a], fontsize=9)
    ax.set_xlabel(rotulo_b)
    ax.set_xticks(np.arange(-.5, len(ordem_b), 1), minor=True)
    ax.set_yticks(np.arange(-.5, len(ordem_a), 1), minor=True)
    ax.grid(which="minor", color="white", linewidth=1.5)
    ax.tick_params(which="both", length=0)
    for sp in ax.spines.values():
        sp.set_visible(False)
    cb_ = fig.colorbar(im, ax=ax, fraction=0.03, pad=0.02)
    cb_.set_label("% dos estudos da linha" if normalizar_linha else "Número de estudos", color=TXT2)
    cb_.outline.set_visible(False)
    cb_.locator = mticker.MaxNLocator(integer=True)
    cb_.update_ticks()
    titulo(ax, titulo_txt, sub)
    salvar(fig, pasta, nome, m)


def barras_empilhadas_100(linhas, categorias, cores, titulo_txt, pasta, nome, sub=None,
                          rotulo_y="", legenda_cols=3, ordem_linhas=None):
    """linhas: dict nome_linha -> Counter(categoria -> n)."""
    ordem_linhas = ordem_linhas or list(linhas.keys())
    tot = {r: sum(linhas[r].values()) for r in ordem_linhas}
    fig, ax = plt.subplots(figsize=(9, max(2.6, 0.45 * len(ordem_linhas) + 1.9)))
    y = np.arange(len(ordem_linhas))[::-1]
    esquerda = np.zeros(len(ordem_linhas))
    dados = {}
    for cat, cor in zip(categorias, cores):
        v = np.array([100 * linhas[r].get(cat, 0) / tot[r] if tot[r] else 0 for r in ordem_linhas])
        dados[cat] = v
        ax.barh(y, v, left=esquerda, color=cor, height=0.66, edgecolor="white", linewidth=1.5, label=cat)
        for yi, l, vi in zip(y, esquerda, v):
            if vi >= 7:
                ax.text(l + vi / 2, yi, f"{vi:.0f}%", ha="center", va="center", fontsize=8,
                        color="white" if cor not in ("#eda100", "#e87ba4", CINZA, "#d7d6d0") else TXT1)
        esquerda += v
    ax.set_yticks(y)
    ax.set_yticklabels([f"{r} (n={tot[r]})" for r in ordem_linhas])
    ax.set_xlim(0, 100)
    ax.xaxis.set_major_formatter(mticker.PercentFormatter())
    ax.set_ylabel(rotulo_y)
    ax.tick_params(axis="y", length=0)
    ax.spines["left"].set_visible(False)
    alt_ax = fig.get_size_inches()[1] * (fig.subplotpars.top - fig.subplotpars.bottom)
    ax.legend(loc="upper center", bbox_to_anchor=(0.4, -0.38 / alt_ax),
              ncol=legenda_cols, frameon=False, fontsize=8.5)
    titulo(ax, titulo_txt, sub)
    salvar(fig, pasta, nome, pd.DataFrame(dados, index=ordem_linhas).round(1))


def quebrar(texto, n=22):
    palavras, linhas, atual = str(texto).split(), [], ""
    for p in palavras:
        if len(atual) + len(p) + 1 > n and atual:
            linhas.append(atual)
            atual = p
        else:
            atual = f"{atual} {p}".strip()
    linhas.append(atual)
    return "\n".join(linhas)




# --------------------------------------------------------------------------------------
# Gráficos
# --------------------------------------------------------------------------------------
def principal_complementar(df, N, pasta):
    """Mecanismo principal (1 por estudo) + menções como complementar, em barras empilhadas."""
    prin = Counter(r for l in lista(df, "Mecanismo de diversificação (principal)") for r in l)
    comp = Counter(r for l in lista(df, "Mecanismos complementares") for r in l)
    cats = [k for k, _ in sorted(prin.items(), key=lambda x: (x[0] == "Outros", -x[1] - comp.get(x[0], 0)))]
    cats += [k for k in comp if k not in cats]
    fig, ax = plt.subplots(figsize=(9, 0.4 * len(cats) + 1.6))
    y = np.arange(len(cats))[::-1]
    v1 = np.array([prin.get(c, 0) for c in cats]); v2 = np.array([comp.get(c, 0) for c in cats])
    ax.barh(y, v1, color=AZUL, height=0.66, edgecolor="white", label="Mecanismo principal")
    ax.barh(y, v2, left=v1, color="#86b6ef", height=0.66, edgecolor="white", label="Mecanismo complementar")
    for yi, a, b in zip(y, v1, v2):
        ax.text(a + b + 0.3, yi, f"{a}" + (f" + {b}" if b else ""), va="center", fontsize=9, color=TXT2)
    ax.set_yticks(y); ax.set_yticklabels(cats)
    ax.set_xlabel("Número de estudos"); ax.set_xlim(0, (v1 + v2).max() * 1.18)
    ax.grid(axis="x", color=GRADE); ax.set_axisbelow(True)
    ax.tick_params(axis="y", length=0); ax.spines["left"].set_visible(False)
    ax.legend(frameon=False, fontsize=8.5, loc="lower right")
    titulo(ax, "Mecanismos de diversificação: principal e complementares",
           f"N = {N}; cada estudo tem um mecanismo principal e pode combinar outros")
    salvar(fig, pasta, "02_mecanismos_principal_complementar",
           pd.DataFrame({"principal": v1, "complementar": v2}, index=cats))


def metodo_literatura(valor):
    s = str(valor).lower(); achados = []
    for nome, p in OrderedDict([
            ("MMR", r"\bmmr\b"), ("Topic Diversification", r"topic diversification"), ("DPP", r"\bdpp\b"),
            ("Max-Min / Max-Sum", r"max-min|max-sum"), ("K-means", r"k-means"), ("Interleaving", r"interleaving"),
            ("ILD-TFIDF / ILD-BERT", r"ild-"), ("SentiRec", r"sentirec"), ("SKNN", r"sknn"),
            ("Passeio aleatório / RWR", r"rwr|rp³β|passeio aleatório"), ("Bandits (LinUCB)", r"linucb"),
            ("Deep Q-Learning / SlateQ", r"dqn|slateq"), ("Otimização submodular", r"sieve|matroide|scp|semantic volume"),
            ("RADio", r"radio"), ("Modelos neurais (MANNeR, XNRS, FreeLB)", r"manner|xnrs|freelb")]).items():
        if re.search(p, s):
            achados.append(nome)
    return achados


METODOS_NOMEADOS = OrderedDict([
    ("MMR", r"\bmmr\b"), ("Topic Diversification", r"topic diversification"), ("K-means", r"k-means"),
    ("DPP", r"\bdpp\b"), ("RP³β", r"rp³β"), ("RADio", r"radio"), ("Max-Min dispersion", r"max-min"),
    ("Max-Sum dispersion", r"max-sum"), ("SKNN", r"sknn"), ("SentiRec", r"sentirec"), ("LinUCB", r"linucb"),
    ("Interleaving (round-robin)", r"interleaving"), ("ILD-TFIDF", r"ild-tfidf"), ("ILD-BERT", r"ild-bert"),
    ("FreeLB", r"freelb"), ("SCP (Contextual Submodular Prediction)", r"\bscp\b"), ("MANNeR", r"manner"),
    ("XNRS", r"xnrs"), ("Trocas locais com restrições de matroide", r"matroide"),
    ("DQN com dominância de Pareto", r"dqn"), ("Passeio aleatório com seleção gulosa", r"passeio aleatório"),
    ("RWR (Random Walk with Restart)", r"\brwr\b"), ("SlateQ", r"slateq"), ("Wolpertinger", r"wolpertinger"),
    ("Kacimi e Gamper (2011; 2012)", r"kacimi"), ("Sieve-Streaming", r"sieve"),
    ("Maximizing Semantic Volume", r"semantic volume"),
])


def metodos_nomeados(valor):
    s = str(valor).lower()
    return [n for n, p in METODOS_NOMEADOS.items() if re.search(p, s)]


def main():
    dfall, art = carregar()
    df = dfall[dfall["codificado"]].reset_index(drop=True)
    N = len(df)
    sub_multi = f"N = {N} estudos; um estudo pode estar em mais de uma categoria"
    sub_uni = f"N = {N} estudos; uma categoria por estudo"

    periodo = df["Ano"].map(lambda a: next((p[2] for p in PERIODOS if p[0] <= a <= p[1]), None)
                            if pd.notna(a) else None)
    ordem_p = [p[2] for p in PERIODOS]
    L = {"periodo": periodo.map(lambda p: [p] if p else [])}
    for k, c in [("mec", "Mecanismo de diversificação (principal)"), ("etapa", "Etapa da diversificação"),
                 ("aspecto", "Aspecto diversificado"), ("ref", "Referência da diversidade"),
                 ("metrica", "Métricas de diversidade"), ("tec", "Técnica do recomendador base"),
                 ("par", "Paradigma do recomendador base"), ("cen", "Cenário de avaliação"),
                 ("outras", "Outras dimensões avaliadas"), ("sim", "Medida de similaridade/distância"),
                 ("rn", "Representação da notícia"), ("rel", "Relação com a diversidade"),
                 ("ctrl", "Controle da relação/trade-off"), ("orig", "Origem dos dados")]:
        L[k] = agrupar_raros(lista(df, c))
    L["familia"] = L["mec"]  # usado pelas funções de evolução e qualidade

    # ============================== Visão geral =====================================
    P0 = "00_visao_geral"
    anos = dfall["Ano"].dropna().astype(int)
    serie = anos.value_counts().reindex(range(anos.min(), anos.max() + 1), fill_value=0)
    fig, ax = plt.subplots(figsize=(9, 3.8))
    ax.bar(serie.index, serie.values, color=[AZUL if a != ANO_PARCIAL else "#86b6ef" for a in serie.index],
           width=0.7, edgecolor="white")
    for a, v in serie.items():
        if v:
            ax.text(a, v + 0.2, str(v), ha="center", fontsize=9, color=TXT2)
    ax.set_xticks(serie.index); ax.set_xticklabels(serie.index, rotation=45)
    ax.set_ylabel("Número de estudos"); ax.set_ylim(0, serie.max() * 1.25)
    ax.yaxis.set_major_locator(mticker.MaxNLocator(integer=True))
    ax.grid(axis="y", color=GRADE); ax.set_axisbelow(True)
    titulo(ax, "Estudos publicados por ano",
           f"N = {len(dfall)}; {ANO_PARCIAL} em tom claro: ano ainda em curso (coleta parcial)")
    salvar(fig, P0, "01_publicacoes_por_ano", serie.to_frame("estudos").assign(acumulado=serie.cumsum()))

    st = art["status"].value_counts(); total = len(art)
    n_dup = int(st.get("Duplicated", 0)); n_inc = int(st.get("Accepted", 0))
    n_triados = total - n_dup
    n_exc_triagem = n_triados - N_BUSCADOS_TEXTO_COMPLETO
    n_nao_recup = N_BUSCADOS_TEXTO_COMPLETO - N_AVALIADOS_TEXTO_COMPLETO
    n_exc_tc = N_AVALIADOS_TEXTO_COMPLETO - n_inc
    assert n_exc_triagem + n_nao_recup + n_exc_tc == st.get("Rejected", 0), "contagens inconsistentes"
    etapas = [("Registros identificados", total, AZUL), ("Duplicatas removidas", n_dup, CINZA),
              ("Registros triados (título e resumo)", n_triados, "#6da7ec"),
              ("Excluídos na triagem", n_exc_triagem, CINZA),
              ("Buscados para texto completo", N_BUSCADOS_TEXTO_COMPLETO, "#6da7ec"),
              ("Texto completo não recuperado", n_nao_recup, CINZA),
              ("Avaliados por texto completo", N_AVALIADOS_TEXTO_COMPLETO, "#6da7ec"),
              ("Excluídos no texto completo", n_exc_tc, CINZA),
              ("Estudos incluídos", n_inc, "#184f95")]
    fig, ax = plt.subplots(figsize=(8, 4.2))
    nomes = [e[0] for e in etapas][::-1]; vals = [e[1] for e in etapas][::-1]
    ax.barh(nomes, vals, color=[e[2] for e in etapas][::-1], height=0.62, edgecolor="white")
    for i, v in enumerate(vals):
        ax.text(v + total * 0.01, i, str(v), va="center", fontsize=9.5, color=TXT1)
    ax.set_xlim(0, total * 1.12); ax.tick_params(axis="y", length=0); ax.spines["left"].set_visible(False)
    ax.set_xlabel("Número de registros"); ax.grid(axis="x", color=GRADE); ax.set_axisbelow(True)
    titulo(ax, "Fluxo de seleção dos estudos")
    salvar(fig, P0, "02_funil_selecao", pd.Series({e[0]: e[1] for e in etapas}, name="registros").to_frame())

    curtos = OrderedDict([
        ("não trata de recomendação de notícias", "CE1 – Não trata de recomendação de notícias"),
        ("não existe um sistema", "CE2 – Notícias sem processo de recomendação"),
        ("apenas mencionada", "CE3 – Diversidade só mencionada (sem método)"),
        ("apenas mede", "CE4 – Apenas mede/analisa diversidade"),
        ("refere-se a algo diferente", "CE5 – \"Diversidade\" com outro sentido"),
        ("beyond-accuracy", "CE6 – Apenas outros objetivos beyond-accuracy"),
        ("recomendação em geral", "CE7 – Método genérico não avaliado em notícias"),
        ("secundário", "CE8 – Estudo secundário/terciário"),
        ("informação suficiente", "CE9 – Informação insuficiente sobre a estratégia"),
        ("versão anterior", "CE11 – Sobreposição substancial de publicação"),
        ("texto completo não pôde ser obtido", "CE12 – Texto completo não recuperado"),
    ])
    def motivo(t):
        if pd.isna(t) or not str(t).strip():
            return "Sem critério registrado"
        return next((v for k, v in curtos.items() if k in str(t)), "Outro critério")
    rej = art[art["status"] == "Rejected"].copy()
    rej["etapa"] = np.where(rej["atualizado"] >= pd.Timestamp(INICIO_TEXTO_COMPLETO),
                            "Texto completo", "Triagem")
    for tit, base, etapa in ETAPA_DA_DECISAO:
        m = (rej["title"].map(chave) == chave(tit)) & (rej["source"] == base)
        assert m.sum() == 1, f"registro não encontrado: {tit}"
        rej.loc[m, "etapa"] = etapa
    rej["motivo"] = rej["selection_criteria"].map(motivo)
    tab = pd.crosstab(rej["motivo"], rej["etapa"]).reindex(columns=["Triagem", "Texto completo"], fill_value=0)
    ordem = [v for v in curtos.values() if v in tab.index] + [m for m in tab.index if m not in curtos.values()]
    tab = tab.loc[ordem]
    n_et = tab.sum()
    fig, axes = plt.subplots(1, 2, figsize=(8.2, 0.40 * len(tab) + 1.8), sharey=True,
                             gridspec_kw={"width_ratios": [2.6, 1.2]})
    for ax, et, cor in zip(axes, ["Triagem", "Texto completo"], [AZUL, "#eb6834"]):
        v = tab[et].values[::-1]
        ax.barh(tab.index[::-1], v, color=cor, height=0.66, edgecolor="white")
        for i, x in enumerate(v):
            if x:
                ax.text(x + max(v) * 0.02, i, str(x), va="center", fontsize=9, color=TXT2)
        ax.set_xlim(0, max(v) * 1.2); ax.set_xlabel("Número de registros")
        ax.xaxis.set_major_locator(mticker.MaxNLocator(integer=True))
        ax.grid(axis="x", color=GRADE); ax.set_axisbelow(True)
        ax.tick_params(axis="y", length=0); ax.spines["left"].set_visible(False)
        if True:  # rótulos dos painéis sempre exibidos
            rot = {"Triagem": "Triagem por título e resumo",
                   "Texto completo": "Texto completo (recuperação e avaliação)"}[et]
            ax.set_title(f"{rot}\nn = {n_et[et]}", loc="left", fontsize=10, color=TXT1)
    fig.tight_layout()
    salvar(fig, P0, "03_motivos_exclusao", tab)

    inc = set(dfall["k"])
    doi_inc = art.loc[art["k"].isin(inc) & (art["status"] == "Accepted"), ["k", "doi_n"]].dropna()
    mapa_doi = dict(zip(doi_inc["doi_n"], doi_inc["k"]))
    art["k_inc"] = art["k"].where(art["k"].isin(inc), art["doi_n"].map(mapa_doi))
    idx = art[art["k_inc"].notna()].groupby("source")["k_inc"].nunique().sort_values(ascending=False)
    origem = dfall["Base de origem"].value_counts()
    fig, ax = plt.subplots(figsize=(8, 3.2))
    bases = idx.index.tolist()[::-1]; y = np.arange(len(bases))
    ax.barh(y + 0.18, [idx[b] for b in bases], height=0.34, color=AZUL, label="Indexou o estudo")
    ax.barh(y - 0.18, [origem.get(b, 0) for b in bases], height=0.34, color="#eb6834",
            label="Registro mantido após deduplicação")
    for yi, b in zip(y, bases):
        ax.text(idx[b] + 0.6, yi + 0.18, str(idx[b]), va="center", fontsize=8.5, color=TXT2)
        ax.text(origem.get(b, 0) + 0.6, yi - 0.18, str(origem.get(b, 0)), va="center", fontsize=8.5, color=TXT2)
    ax.set_yticks(y); ax.set_yticklabels(bases); ax.set_xlim(0, len(dfall) * 1.08)
    ax.set_xlabel("Número de estudos incluídos"); ax.tick_params(axis="y", length=0)
    ax.spines["left"].set_visible(False); ax.legend(frameon=False, fontsize=8.5, loc="lower right")
    titulo(ax, "Cobertura das bases de dados sobre os estudos incluídos", f"N = {len(dfall)}")
    salvar(fig, P0, "04_bases_de_dados",
           pd.DataFrame({"indexou": idx, "mantido": origem}).fillna(0).astype(int))

    # ============================== RQ1 ==============================================
    P1 = "RQ1_estrategias"
    barras_h(contar_listas(lista(df, "Mecanismo de diversificação (principal)"), outros=False), N,
             "Mecanismo principal de diversificação", P1, "01_mecanismo_principal", sub_uni, largura=9)
    principal_complementar(df, N, P1)
    barras_h(contar_listas(lista(df, "Origem do método de diversificação"), outros=False), N,
             "Origem do método de diversificação", P1, "03_origem_metodo", sub_uni)
    lit = agrupar_raros(df["Algoritmo de diversificação"].map(metodo_literatura))
    s = contar_listas(lit, outros=False)
    barras_h(s, N, "Métodos da literatura aplicados ou usados como base", P1,
             "04_metodos_literatura_referenciados", f"Estudos que aplicam ou adaptam um método existente (de N = {N})")
    todos = df["Algoritmo de diversificação"].map(metodos_nomeados)
    barras_h(contar_listas(todos, outros=False), N, "Métodos da literatura aplicados ou usados como base (todos)",
             P1, "04b_metodos_literatura_todos",
             f"Cada método citado no campo 'Algoritmo de diversificação', sem agrupamento (de N = {N})", largura=9)
    barras_h(contar_listas(lista(df, "Paradigma do recomendador base"), outros=False), N,
             "Paradigma do recomendador base", P1, "05_paradigma_recomendador_base", sub_uni)
    barras_h(contar_listas(lista(df, "Técnica do recomendador base")), N,
             "Técnica do recomendador base", P1, "06_tecnica_recomendador_base", sub_multi)
    heatmap(L["mec"], L["tec"], "Mecanismo de diversificação × técnica do recomendador base", P1,
            "07_heatmap_mecanismo_x_tecnica", "Mecanismo principal", "Técnica do recomendador base",
            "Número de estudos")
    evolucao_por_periodo(L["mec"], periodo, ordem_p, P1)
    evolucao_linhas(df, L["mec"], P1, "09_evolucao_mecanismos_acumulada",
                    "Crescimento acumulado dos mecanismos de diversificação", "mecanismos")

    # ============================== RQ2 ==============================================
    P2 = "RQ2_etapa"
    s = contar_listas(lista(df, "Etapa da diversificação"), outros=False)
    barras_h(s.reindex([e for e in ORDEM_ETAPAS if e in s.index]).sort_values(ascending=False), N,
             "Etapa do processo de recomendação em que a diversificação é aplicada", P2,
             "01_etapa_diversificacao", sub_multi)
    n_et = L["etapa"].map(len).value_counts().sort_index()
    barras_h(n_et.rename(index=lambda k: f"{k} etapa" + ("s" if k > 1 else "")), N,
             "Número de etapas em que cada estudo diversifica", P2, "02_numero_etapas_por_estudo", f"N = {N}")
    ordem_et = [e for e in ORDEM_ETAPAS if any(e in l for l in L["etapa"])]
    heatmap(L["mec"], L["etapa"], "Mecanismo de diversificação × etapa de aplicação", P2,
            "03_heatmap_mecanismo_x_etapa", "Mecanismo principal", "Etapa (ordem do pipeline)",
            "Número de estudos", ordem_b=ordem_et)
    heatmap(L["etapa"], L["periodo"], "Etapa de diversificação ao longo do tempo", P2, "04_etapa_x_periodo",
            "Etapa", "Período", "Número de estudos", ordem_a=ordem_et, ordem_b=ordem_p, largura=8,
            quebra_b=30, rot=0)
    evolucao_linhas(df, L["etapa"], P2, "05_etapa_evolucao_acumulada",
                    "Crescimento acumulado das etapas de diversificação", "etapas")

    # ============================== RQ3 ==============================================
    P3 = "RQ3_representacao"
    barras_h(contar_listas(lista(df, "Aspecto diversificado")), N, "Aspectos diversificados", P3,
             "01_aspectos_diversificados", sub_multi)
    barras_h(contar_listas(lista(df, "Referência da diversidade"), outros=False), N,
             "Em relação a quê a diversidade é definida", P3, "02_referencia_diversidade", sub_multi)
    barras_h(contar_listas(lista(df, "Representação da notícia")), N, "Representação das notícias", P3,
             "03_representacao_noticia", sub_multi,
             destaque={"Texto (forma não especificada)"})
    barras_h(contar_listas(lista(df, "Representação do usuário")), N, "Representação dos usuários", P3,
             "04_representacao_usuario", sub_multi, destaque={"Não informado", "Sem perfil (não personalizado)"})
    barras_h(contar_listas(lista(df, "Medida de similaridade/distância")), N,
             "Medidas de similaridade/distância usadas na diversificação", P3, "05_medidas_similaridade",
             sub_multi, destaque={"Não especificada", "Não se aplica"})
    heatmap(L["mec"], L["aspecto"], "Mecanismo de diversificação × aspecto diversificado", P3,
            "06_heatmap_mecanismo_x_aspecto", "Mecanismo principal", "Aspecto diversificado", "Número de estudos")
    heatmap(L["aspecto"], L["ref"], "Aspecto diversificado × referência da diversidade", P3,
            "07_heatmap_aspecto_x_referencia", "Aspecto", "Referência da diversidade", "Número de estudos",
            quebra_b=20)
    heatmap(L["rn"], L["sim"], "Representação da notícia × medida de similaridade", P3,
            "08_heatmap_representacao_x_similaridade", "Representação da notícia", "Medida de similaridade/distância",
            "Número de estudos")
    heatmap(L["aspecto"], L["periodo"], "Aspectos diversificados ao longo do tempo", P3, "09_aspecto_x_periodo",
            "Aspecto diversificado", "Período", "Número de estudos", ordem_b=ordem_p, largura=8, quebra_b=30, rot=0)
    evolucao_linhas(df, L["aspecto"], P3, "10_aspecto_evolucao_acumulada",
                    "Crescimento acumulado dos aspectos diversificados", "aspectos")

    # ============================== RQ4 ==============================================
    P4 = "RQ4_datasets"
    base_ds = lambda t: "Conjunto próprio dos autores" if t.startswith("Próprio") else re.sub(r"\s*\(.*\)$", "", t)
    ds = df["Dataset"].map(lambda v: list(OrderedDict.fromkeys(base_ds(t) for t in dividir(v))))
    barras_h(contar_listas(ds), N, "Datasets utilizados", P4, "01_datasets",
             sub_multi + "; versões do mesmo dataset agrupadas", destaque={"Não informado"})
    mind = df["Dataset"].map(lambda v: [t for t in dividir(v) if t.startswith("MIND")])
    n_mind = int((mind.map(len) > 0).sum())
    barras_h(contar_listas(mind, outros=False), n_mind, "Versões do MIND utilizadas", P4, "02_versoes_mind",
             f"Estudos que usam o MIND = {n_mind}; percentuais sobre esses estudos")
    barras_h(contar_listas(lista(df, "Origem dos dados"), outros=False), N, "Origem dos dados de avaliação",
             P4, "03_origem_dos_dados", sub_multi)
    ordem = ["Sim", "Parcialmente", "Mediante solicitação", "Não", "Não informado"]
    pub = df["Dataset público?"].value_counts().reindex(ordem).dropna().astype(int)
    barras_h(pub, N, "Disponibilidade pública dos datasets", P4, "04_dataset_publico", sub_uni,
             cores=[{"Sim": "#1baf7a", "Parcialmente": "#eda100", "Mediante solicitação": "#86b6ef",
                     "Não": "#e34948"}.get(k, CINZA) for k in pub.index[::-1]])
    nd = ds.map(lambda l: len([x for x in l if x != "Não informado"])).value_counts().sort_index()
    barras_h(nd.rename(index=lambda k: f"{k} dataset" + ("s" if k != 1 else "")), N,
             "Número de datasets por estudo", P4, "05_datasets_por_estudo", f"N = {N}")
    heatmap(L["orig"], L["periodo"], "Origem dos dados ao longo do tempo", P4, "06_origem_dados_x_periodo",
            "Origem dos dados", "Período", "Número de estudos", ordem_b=ordem_p, largura=8, quebra_b=30, rot=0)
    evolucao_linhas(df, L["orig"], P4, "07_origem_dados_evolucao_acumulada",
                    "Crescimento acumulado por origem dos dados", "origens")

    # ============================== RQ5 ==============================================
    P5 = "RQ5_metricas"
    barras_h(contar_listas(lista(df, "Métricas de diversidade")), N, "Métricas de diversidade (por formulação)",
             P5, "01_metricas_diversidade", sub_multi,
             destaque={"Sem métrica de diversidade", "Métrica não especificada"})
    heatmap(L["aspecto"], L["metrica"], "Aspecto diversificado × métrica de diversidade", P5,
            "02_heatmap_aspecto_x_metrica", "Aspecto diversificado", "Métrica de diversidade", "Número de estudos")
    heatmap(L["mec"], L["metrica"], "Mecanismo de diversificação × métrica de diversidade", P5,
            "03_heatmap_mecanismo_x_metrica", "Mecanismo principal", "Métrica de diversidade", "Número de estudos")
    heatmap(L["ref"], L["metrica"], "Referência da diversidade × métrica usada", P5,
            "04_heatmap_referencia_x_metrica", "Referência definida pelo método", "Métrica de diversidade",
            "Número de estudos")
    heatmap(L["metrica"], L["periodo"], "Métricas de diversidade ao longo do tempo", P5, "05_metrica_x_periodo",
            "Métrica de diversidade", "Período", "Número de estudos", ordem_b=ordem_p, largura=8, quebra_b=30, rot=0)
    evolucao_linhas(df, L["metrica"], P5, "06_metrica_evolucao_acumulada",
                    "Crescimento acumulado das métricas de diversidade", "métricas")

    # ============================== RQ6 ==============================================
    P6 = "RQ6_avaliacao"
    barras_h(contar_listas(lista(df, "Cenário de avaliação"), outros=False), N, "Cenários de avaliação", P6,
             "01_cenarios_avaliacao", sub_multi)

    def paradigma(l):
        s = set(l); off = "Offline" in s; hum = bool(s & {"Estudo com usuários", "Experimento online/campo (A/B)"})
        sim = "Simulação" in s
        if off and hum: return "Offline + com usuários/online"
        if hum: return "Somente com usuários/online"
        if off and sim: return "Offline + simulação"
        if off: return "Somente offline"
        return "Somente simulação"
    par = L["cen"].map(paradigma)
    barras_h(par.value_counts(), N, "Combinação de cenários de avaliação", P6, "02_combinacao_cenarios",
             f"N = {N}; categorias mutuamente exclusivas")
    barras_h(contar_listas(lista(df, "Análises realizadas"), outros=False), N, "Análises realizadas", P6,
             "03_analises_realizadas", sub_multi)
    sig = df["Significância estatística"].value_counts()
    barras_h(sig, N, "Uso de testes de significância estatística", P6, "04_significancia_estatistica", sub_uni,
             cores=[{"Teste reportado": "#1baf7a", "Teste parcial": "#eda100",
                     "Teste não reportado": "#e34948"}.get(k, CINZA) for k in sig.index[::-1]])
    heatmap(L["mec"], L["cen"], "Mecanismo de diversificação × cenário de avaliação", P6,
            "05_heatmap_mecanismo_x_cenario", "Mecanismo principal", "Cenário de avaliação", "Número de estudos")
    linhas = OrderedDict((p, Counter(par[periodo == p])) for p in ordem_p)
    cats = [c for c in ["Somente offline", "Offline + simulação", "Offline + com usuários/online",
                        "Somente com usuários/online", "Somente simulação"] if any(c in v for v in linhas.values())]
    cor = {"Somente offline": AZUL, "Offline + simulação": "#86b6ef", "Offline + com usuários/online": "#1baf7a",
           "Somente com usuários/online": "#6250d6", "Somente simulação": "#eda100"}
    barras_empilhadas_100(linhas, cats, [cor[c] for c in cats], "Combinação de cenários ao longo do tempo", P6,
                          "06_cenarios_x_periodo", "% dos estudos de cada período", "Período")
    reprodutibilidade(df, N, P6)

    # ============================== RQ7 ==============================================
    P7 = "RQ7_tradeoffs"
    barras_h(contar_listas(lista(df, "Outras dimensões avaliadas")), N,
             "Outras dimensões avaliadas além da diversidade", P7, "01_outras_dimensoes", sub_multi)
    barras_h(contar_listas(lista(df, "Métricas das outras dimensões")), N,
             "Métricas usadas para as outras dimensões", P7, "02_metricas_outras_dimensoes", sub_multi, largura=9)
    ORD_REL = ["Ganho conjunto", "Sem perda significativa", "Dependente da configuração", "Trade-off",
               "Não analisada/inconclusiva"]
    COR_REL = ["#1baf7a", "#86b6ef", "#eda100", "#eb6834", CINZA]
    rel = df["Relação com a diversidade"].map(lambda v: CURTO.get(v, v)).value_counts().reindex(ORD_REL).dropna().astype(int)
    barras_h(rel, N, "Relação reportada entre diversidade e as demais dimensões", P7, "03_relacao_diversidade",
             sub_uni, cores=[dict(zip(ORD_REL, COR_REL))[k] for k in rel.index[::-1]])
    rel_s = df["Relação com a diversidade"].map(lambda v: CURTO.get(v, v))
    linhas = OrderedDict()
    for d in [a for a, _ in Counter(x for l in L["outras"] for x in set(l)).most_common() if a != "Outros"]:
        linhas[d] = Counter(rel_s[L["outras"].map(lambda l: d in l)])
    barras_empilhadas_100(linhas, ORD_REL, COR_REL, "Relação com a diversidade por dimensão avaliada", P7,
                          "04_relacao_por_dimensao", "% dos estudos que avaliam cada dimensão", legenda_cols=3)
    ORD_CTRL = ["Parâmetro com análise de sensibilidade", "Parâmetro sem análise sistemática",
                "Controle implícito", "Sem controle"]
    ctrl = df["Controle da relação/trade-off"].map(lambda v: CURTO.get(v, v)).value_counts().reindex(ORD_CTRL).dropna().astype(int)
    barras_h(ctrl, N, "Controle do compromisso diversidade × outras dimensões", P7, "05_controle_tradeoff", sub_uni)
    heatmap(L["rel"], L["ctrl"], "Relação reportada × forma de controle do trade-off", P7,
            "06_heatmap_relacao_x_controle", "Relação com a diversidade", "Controle do trade-off",
            "Número de estudos", ordem_a=[o for o in ORD_REL if any(o in l for l in L["rel"])],
            ordem_b=[o for o in ORD_CTRL if any(o in l for l in L["ctrl"])], quebra_b=22)
    linhas = OrderedDict()
    for m in [a for a, _ in Counter(x for l in L["mec"] for x in l).most_common() if a != "Outros"]:
        c = Counter(rel_s[L["mec"].map(lambda l: m in l)])
        if sum(c.values()) >= MIN_FREQ:
            linhas[m] = c
    barras_empilhadas_100(linhas, ORD_REL, COR_REL, "Relação com a diversidade por mecanismo", P7,
                          "07_relacao_por_mecanismo", "% dos estudos de cada mecanismo principal", legenda_cols=3)

    # ============================== RQ8 ==============================================
    P8 = "RQ8_limitacoes"
    l_rsl = lista(df, "Limitações identificadas na RSL (categorias)")
    l_dec = df["Limitações declaradas (categorias)"].map(lambda v: [x for x in rotulos(v) if x != "Não declaradas"])
    c_rsl = Counter(x for l in l_rsl for x in l); c_dec = Counter(x for l in l_dec for x in l)
    ordena = lambda c: pd.Series(c).sort_values(ascending=False).pipe(
        lambda s: pd.concat([s.drop("Outras", errors="ignore"), s.reindex(["Outras"]).dropna()]).astype(int))
    barras_h(ordena(c_rsl), N, "Limitações recorrentes identificadas na RSL", P8, "01_limitacoes_identificadas_rsl",
             f"N = {N}; categorias do livro de códigos", largura=9)
    sem_dec = int((l_dec.map(len) == 0).sum())
    barras_h(ordena(c_dec), N, "Limitações declaradas pelos autores", P8, "02_limitacoes_declaradas",
             f"N = {N}; {sem_dec} estudos não declaram limitações", largura=9)
    temas = list(ordena(c_rsl).index) + [t for t in ordena(c_dec).index if t not in c_rsl]
    fig, ax = plt.subplots(figsize=(9.5, 0.42 * len(temas) + 1.8))
    y = np.arange(len(temas))[::-1]
    v1 = np.array([100 * c_dec.get(t, 0) / N for t in temas]); v2 = np.array([100 * c_rsl.get(t, 0) / N for t in temas])
    ax.hlines(y, np.minimum(v1, v2), np.maximum(v1, v2), color="#c3c2b7", lw=2)
    ax.scatter(v1, y, color="#eb6834", s=55, zorder=3, label="Declarada pelos autores")
    ax.scatter(v2, y, color=AZUL, s=55, zorder=3, label="Identificada na RSL")
    ax.set_yticks(y); ax.set_yticklabels(temas); ax.xaxis.set_major_formatter(mticker.PercentFormatter())
    ax.set_xlabel("% dos estudos"); ax.grid(axis="x", color=GRADE); ax.set_axisbelow(True)
    ax.tick_params(axis="y", length=0); ax.spines["left"].set_visible(False)
    ax.legend(frameon=False, loc="lower right", fontsize=9)
    titulo(ax, "Limitações declaradas × identificadas na RSL", f"N = {N}; por categoria")
    salvar(fig, P8, "03_declaradas_x_identificadas",
           pd.DataFrame({"declarada": [c_dec.get(t, 0) for t in temas],
                         "identificada_rsl": [c_rsl.get(t, 0) for t in temas]}, index=temas))
    lacunas(df, L, N, P8)

    qualidade(df, L, N)


def evolucao_por_periodo(listas_familia, periodo, ordem_p, pasta):
    """Mapa de calor família × período, normalizado pelo nº de estudos do período."""
    fams = [k for k, _ in sorted(Counter(x for l in listas_familia for x in set(l)).items(),
                                 key=lambda x: (x[0] == "Outros", -x[1]))]
    n_p = periodo.value_counts()
    m = pd.DataFrame(0, index=fams, columns=ordem_p)
    for l, p in zip(listas_familia, periodo):
        if p:
            for f in set(l):
                m.loc[f, p] += 1
    pct = m.div(n_p.reindex(ordem_p), axis=1) * 100
    fig, ax = plt.subplots(figsize=(8, 0.42 * len(fams) + 2.2))
    im = ax.imshow(pct.values, cmap=SEQ, aspect="auto", vmin=0)
    vmax = pct.values.max()
    for i in range(len(fams)):
        for j in range(len(ordem_p)):
            if m.values[i, j]:
                ax.text(j, i, f"{pct.values[i, j]:.0f}%\n({m.values[i, j]})", ha="center", va="center",
                        fontsize=7.8, color="white" if pct.values[i, j] > vmax * 0.55 else TXT1)
    ax.set_xticks(range(len(ordem_p)))
    ax.set_xticklabels([f"{p}\n(n={n_p.get(p, 0)})" for p in ordem_p])
    ax.set_yticks(range(len(fams)))
    ax.set_yticklabels(fams)
    ax.set_xticks(np.arange(-.5, len(ordem_p), 1), minor=True)
    ax.set_yticks(np.arange(-.5, len(fams), 1), minor=True)
    ax.grid(which="minor", color="white", linewidth=1.5)
    ax.tick_params(which="both", length=0)
    for sp in ax.spines.values():
        sp.set_visible(False)
    cb = fig.colorbar(im, ax=ax, fraction=0.03, pad=0.02)
    cb.set_label("% dos estudos do período", color=TXT2)
    cb.outline.set_visible(False)
    titulo(ax, "Evolução dos mecanismos de diversificação",
           "% dos estudos de cada período que usam o mecanismo (nº de estudos entre parênteses)")
    salvar(fig, pasta, "08_evolucao_mecanismos_periodo", pct.round(1))


def evolucao_linhas(df, listas, pasta, nome, titulo_txt, rotulo_itens, top=8):
    """Contagem acumulada de estudos por ano para as categorias mais frequentes de um campo."""
    c = Counter(x for l in listas for x in set(l))
    cats = [k for k, _ in c.most_common() if k != "Outros"]
    top_f = cats[:min(top, len(CATEG))]
    anos = range(int(df["Ano"].min()), int(df["Ano"].max()) + 1)
    fig, ax = plt.subplots(figsize=(10, 4.8))
    tab = {}
    finais = []
    for f, cor in zip(top_f, CATEG):
        por_ano = Counter(a for a, l in zip(df["Ano"], listas) if f in l and pd.notna(a))
        acum = np.cumsum([por_ano.get(a, 0) for a in anos])
        tab[f] = acum
        ax.plot(list(anos), acum, color=cor, lw=2)
        ax.plot([max(anos)], [acum[-1]], "o", color=cor, ms=4)
        finais.append((acum[-1], f, cor))
    ymax = max(v for v, _, _ in finais)
    ax.set_ylim(0, ymax * 1.06)
    ys = _posicoes_rotulos([v for v, _, _ in finais], gap=ymax * 0.065)
    x0, x1 = max(anos), max(anos) + 0.9
    for (v, f, cor), y in zip(finais, ys):
        ax.plot([x0 + 0.08, x1 - 0.12], [v, y], color=cor, lw=0.8, clip_on=False)
        ax.text(x1, y, f"{f} ({v})", color=TXT1, fontsize=8.5, va="center", clip_on=False)
    ax.set_xlim(min(anos), max(anos) + 0.2)
    ax.spines["right"].set_visible(False)
    ax.set_ylabel("Estudos acumulados")
    ax.yaxis.set_major_locator(mticker.MaxNLocator(integer=True))
    ax.grid(axis="y", color=GRADE)
    ax.set_axisbelow(True)
    sub = (f"Todas as {len(top_f)} categorias" if len(top_f) == len(cats)
           else f"As {len(top_f)} {rotulo_itens} mais frequentes (de {len(cats)})")
    titulo(ax, titulo_txt, sub + f"; contagem acumulada de estudos por ano ({ANO_PARCIAL} parcial)")
    salvar(fig, pasta, nome, pd.DataFrame(tab, index=list(anos)))


def _posicoes_rotulos(valores, gap):
    """Posições verticais sem sobreposição, mantendo cada grupo centrado nos seus pontos."""
    ordem = np.argsort(valores)
    grupos = []  # cada grupo: lista de (índice, valor)
    for i in ordem:
        grupos.append([(i, valores[i])])
        while len(grupos) > 1:
            g1, g2 = grupos[-2], grupos[-1]
            c1 = np.mean([v for _, v in g1]) - gap * (len(g1) - 1) / 2
            topo1 = c1 + gap * (len(g1) - 1)
            c2 = np.mean([v for _, v in g2]) - gap * (len(g2) - 1) / 2
            if c2 < topo1 + gap:
                grupos[-2:] = [g1 + g2]
            else:
                break
    pos = [0.0] * len(valores)
    for g in grupos:
        base = max(np.mean([v for _, v in g]) - gap * (len(g) - 1) / 2, 0)
        for k, (i, _) in enumerate(g):
            pos[i] = base + k * gap
    return pos


def _desempilhar_rotulos(ax, minimo=None):
    textos = [t for t in ax.texts]
    if not textos:
        return
    ylim = ax.get_ylim()
    minimo = minimo or (ylim[1] - ylim[0]) * 0.055
    textos.sort(key=lambda t: t.get_position()[1])
    ultimo = -1e9
    for t in textos:
        x, y = t.get_position()
        y = max(y, ultimo + minimo)
        t.set_position((x, y))
        ultimo = y


def reprodutibilidade(df, N, pasta):
    def first(v):
        d = dividir(v)
        return d[0] if d else "Não informado/não verificável"
    ordem = ["Código público", "Dados públicos", "Disponibilidade parcial", "Disponível mediante solicitação",
             "Dados indisponíveis", "Não informado/não verificável"]
    cores = {"Código público": "#1baf7a", "Dados públicos": "#1baf7a", "Disponibilidade parcial": "#eda100",
             "Disponível mediante solicitação": "#86b6ef", "Dados indisponíveis": "#e34948",
             "Não informado/não verificável": CINZA}
    cod = df["Código disponível?"].map(first).value_counts()
    dad = df["Dados disponível?"].map(first).value_counts()
    rotulo = {"Código público": "Público", "Dados públicos": "Público"}
    cats = []
    for o in ordem:
        r = rotulo.get(o, o)
        if r not in cats and (o in cod.index or o in dad.index):
            cats.append(r)
    cod = cod.rename(index=rotulo).groupby(level=0).sum()
    dad = dad.rename(index=rotulo).groupby(level=0).sum()
    cor_cat = {rotulo.get(k, k): v for k, v in cores.items()}
    fig, ax = plt.subplots(figsize=(9, 2.6))
    for i, (nome, ser) in enumerate([("Dados", dad), ("Código", cod)]):
        esq = 0
        for c in cats:
            v = ser.get(c, 0)
            if v:
                ax.barh(i, v, left=esq, color=cor_cat[c], edgecolor="white", linewidth=1.5, height=0.6,
                        label=c if i == 0 else None)
                if v >= 2:
                    ax.text(esq + v / 2, i, str(v), ha="center", va="center", fontsize=8.5,
                            color=TXT1 if cor_cat[c] in ("#eda100", CINZA, "#86b6ef") else "white")
                esq += v
    for c in cats:  # garante legenda completa
        if c not in dad.index:
            ax.barh(0, 0, color=cor_cat[c], label=c)
    ax.set_yticks([0, 1])
    ax.set_yticklabels(["Dados", "Código"])
    ax.set_xlim(0, N)
    ax.set_xlabel("Número de estudos")
    ax.tick_params(axis="y", length=0)
    ax.spines["left"].set_visible(False)
    ax.legend(frameon=False, fontsize=8.5, ncol=3, loc="upper center", bbox_to_anchor=(0.5, -0.32))
    titulo(ax, "Disponibilidade de código e dados (reprodutibilidade)", f"N = {N}")
    salvar(fig, pasta, "06_disponibilidade_codigo_dados",
           pd.DataFrame({"codigo": cod, "dados": dad}).fillna(0).astype(int))


def lacunas(df, L, N, pasta):
    """Indicadores-síntese de lacunas metodológicas (% dos estudos)."""
    cen = L["cen"]
    ind = OrderedDict([
        ("Sem teste estatístico reportado", (df["Significância estatística"] == "Teste não reportado").sum()),
        ("Sem métrica de diversidade definida", L["metrica"].map(
            lambda l: bool(set(l) & {"Sem métrica de diversidade", "Métrica não especificada"})).sum()),
        ("Sem avaliação com usuários ou online", cen.map(
            lambda l: not (set(l) & {"Estudo com usuários", "Experimento online/campo (A/B)"})).sum()),
        ("Relação com outras dimensões não analisada", (df["Relação com a diversidade"] == "Não analisada/inconclusiva").sum()),
        ("Sem controle do trade-off", (df["Controle da relação/trade-off"] == "Sem controle").sum()),
        ("Código não disponível ou não verificável", (~df["Código disponível?"].astype(str).str.startswith("Código público")).sum()),
        ("Avaliação em um único dataset", df["Dataset"].map(lambda v: len([t for t in dividir(v) if t != "Não informado"]) == 1).sum()),
        ("Medida de similaridade não especificada", L["sim"].map(lambda l: "Não especificada" in l).sum()),
    ])
    s = pd.Series(ind).sort_values(ascending=False).astype(int)
    barras_h(s, N, "Indicadores de lacunas metodológicas", pasta, "04_indicadores_lacunas",
             f"N = {N}; % dos estudos que apresentam cada lacuna", cor="#e34948", largura=9)


def qualidade(df, L, N):
    P = "Qualidade"
    qa = [c for c in df.columns if re.fullmatch(r"QA\d+", str(c))]
    tot_col = [c for c in df.columns if str(c).startswith("Total")][0]
    # Distribuição das notas por critério
    linhas = OrderedDict()
    for q in qa:
        v = df[q].dropna()
        linhas[q] = Counter({"Atende (1)": int((v == 1).sum()), "Atende parcialmente (0,5)": int((v == 0.5).sum()),
                             "Não atende (0)": int((v == 0).sum())})
    barras_empilhadas_100(linhas, ["Atende (1)", "Atende parcialmente (0,5)", "Não atende (0)"],
                          ["#1baf7a", "#eda100", "#e34948"], "Avaliação de qualidade por critério",
                          P, "01_qualidade_por_criterio", "% dos estudos avaliados em cada critério",
                          "Critério de qualidade")
    # Histograma do escore total
    t = df[tot_col].dropna()
    fig, ax = plt.subplots(figsize=(8, 3.6))
    bins = np.arange(t.min() - 0.25, t.max() + 0.5, 0.5)
    ax.hist(t, bins=bins, color=AZUL, edgecolor="white", linewidth=1.5)
    ax.axvline(t.mean(), color="#eb6834", lw=2, ls="--")
    ax.text(t.mean() + 0.08, ax.get_ylim()[1] * 0.92, f"média = {t.mean():.2f}\nmediana = {t.median():.1f}",
            color="#c4501f", fontsize=8.5, va="top")
    ax.set_xlabel(f"Escore total de qualidade ({tot_col.replace('Total ', '')}, máximo 10)")
    ax.set_ylabel("Número de estudos")
    ax.yaxis.set_major_locator(mticker.MaxNLocator(integer=True))
    ax.grid(axis="y", color=GRADE)
    ax.set_axisbelow(True)
    titulo(ax, "Distribuição do escore de qualidade", f"Estudos avaliados = {len(t)}")
    salvar(fig, P, "02_distribuicao_escore_qualidade", t.value_counts().sort_index().to_frame("estudos"))
    # Qualidade por família de estratégia (pontos + mediana)
    fams = [k for k, _ in Counter(x for l in L["familia"] for x in set(l)).most_common()]
    fig, ax = plt.subplots(figsize=(9, 0.4 * len(fams) + 1.6))
    rng = np.random.default_rng(1)
    tab = {}
    for i, f in enumerate(fams[::-1]):
        vals = df.loc[L["familia"].map(lambda l: f in l), tot_col].dropna()
        ax.scatter(vals, i + rng.uniform(-0.15, 0.15, len(vals)), s=16, color=AZUL, alpha=0.55, lw=0)
        ax.scatter([vals.median()], [i], marker="|", s=320, color="#eb6834", lw=2.5, zorder=3)
        tab[f] = {"n": len(vals), "mediana": vals.median(), "media": round(vals.mean(), 2)}
    ax.set_yticks(range(len(fams)))
    ax.set_yticklabels([f"{f} (n={tab[f]['n']})" for f in fams[::-1]])
    ax.set_xlabel("Escore total de qualidade")
    ax.grid(axis="x", color=GRADE)
    ax.set_axisbelow(True)
    ax.tick_params(axis="y", length=0)
    ax.spines["left"].set_visible(False)
    titulo(ax, "Qualidade dos estudos por mecanismo de diversificação",
           "Pontos = estudos; traço laranja = mediana")
    salvar(fig, P, "03_qualidade_por_mecanismo", pd.DataFrame(tab).T)
    # Qualidade ao longo do tempo
    fig, ax = plt.subplots(figsize=(9, 3.8))
    d = df[["Ano", tot_col]].dropna()
    ax.scatter(d["Ano"].astype(int) + rng.uniform(-0.2, 0.2, len(d)), d[tot_col], s=20, color=AZUL,
               alpha=0.55, lw=0)
    med = d.groupby("Ano")[tot_col].median()
    ax.plot(med.index.astype(int), med.values, color="#eb6834", lw=2, marker="o", ms=4, label="Mediana anual")
    ax.set_xlabel("Ano de publicação")
    ax.set_ylabel("Escore total de qualidade")
    ax.grid(axis="y", color=GRADE)
    ax.set_axisbelow(True)
    ax.legend(frameon=False, fontsize=8.5)
    titulo(ax, "Qualidade dos estudos ao longo do tempo", "Pontos = estudos")
    salvar(fig, P, "04_qualidade_por_ano", med.to_frame("mediana"))
    # Qualidade × impacto (taxa de citações)
    tx = [c for c in df.columns if str(c).startswith("Taxa média")]
    if tx:
        d = df[[tot_col, tx[0]]].dropna()
        fig, ax = plt.subplots(figsize=(7.5, 4.2))
        ax.scatter(d[tot_col] + rng.uniform(-0.08, 0.08, len(d)), d[tx[0]], s=24, color=AZUL, alpha=0.6, lw=0)
        rho = d[tot_col].corr(d[tx[0]], method="spearman")
        ax.set_yscale("symlog", linthresh=1)
        ax.set_xlabel("Escore total de qualidade")
        ax.set_ylabel("Citações por ano (escala log)")
        ax.grid(color=GRADE)
        ax.set_axisbelow(True)
        titulo(ax, "Qualidade × impacto (citações por ano)", f"ρ de Spearman = {rho:.2f}; n = {len(d)}")
        salvar(fig, P, "05_qualidade_x_citacoes", d)


if __name__ == "__main__":
    main()
