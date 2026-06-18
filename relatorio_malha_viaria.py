# -*- coding: utf-8 -*-
"""
Relatório de Logradouros / Malha Viária - Versão Local
=======================================================

Gera, a partir de um CSV de trechos de logradouros:
  - Relatório em Excel (.xlsx)
  - Relatório em PDF (.pdf)
  - Gráficos individuais (.png) salvos em uma pasta separada

Diferença em relação à versão do Colab:
  - Roda 100% localmente (sem google.colab, sem !pip, sem display do Jupyter).
  - Todas as tabelas e gráficos que possuem TOTAIS são apresentados em
    PORCENTAGEM (participação na malha viária), com a linha de TOTAL = 100%,
    para não causar confusão em quem lê o relatório.

Uso:
  python relatorio_malha_viaria.py
  python relatorio_malha_viaria.py caminho/do/arquivo.csv

As colunas usam o padrão abaixo (altere em COLUNAS_PADRAO se necessário):
  RUA, TRECHO, SETOR, BAIRRO, STATUS, COMP_TRECH, TIPO
"""

import os
import sys
import glob
import shutil
import argparse
import warnings
from datetime import datetime

import matplotlib
matplotlib.use("Agg")  # backend sem tela - apenas salva arquivos

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.patches import Patch
from matplotlib.backends.backend_pdf import PdfPages

warnings.filterwarnings("ignore")

# ==========================================================================
# CONFIGURAÇÃO
# ==========================================================================

# Colunas padrão do CSV (nome exato conforme o arquivo)
COLUNAS_PADRAO = {
    "via": "RUA",
    "trecho": "TRECHO",
    "setor": "SETOR",
    "bairro": "BAIRRO",
    "status": "STATUS",
    "comprimento": "COMP_TRECH",
    "pavimentacao": "TIPO",
}

# Paleta principal (amarelo / preto / branco)
CORES_PRINCIPAIS = [
    "#FFD700", "#FFC107", "#FFEB3B", "#FFF59D",
    "#2C2C2C", "#424242", "#616161", "#9E9E9E",
    "#FFB300", "#FF8F00", "#F5F5F5", "#BDBDBD",
]

LOGO_PADRAO = "logomarca-eixo-cores-2.png"
PASTA_GRAFICOS = "graficos"  # subpasta dentro da pasta de saída


# ==========================================================================
# FORMATAÇÃO BRASILEIRA
# ==========================================================================

def formatar_numero_br(valor, decimais=2):
    """Formata número no padrão brasileiro (1.234,56)."""
    if valor is None or valor == "" or (isinstance(valor, float) and pd.isna(valor)):
        return ""
    try:
        texto = f"{float(valor):,.{decimais}f}"
        return texto.replace(",", "X").replace(".", ",").replace("X", ".")
    except (ValueError, TypeError):
        return str(valor)


def formatar_inteiro_br(valor):
    """Formata inteiro no padrão brasileiro (1.234)."""
    if valor is None or valor == "" or (isinstance(valor, float) and pd.isna(valor)):
        return ""
    try:
        return f"{int(valor):,}".replace(",", ".")
    except (ValueError, TypeError):
        return str(valor)


def formatar_pct(valor, decimais=2):
    """Formata porcentagem no padrão brasileiro (23,45%)."""
    if valor is None or valor == "" or (isinstance(valor, float) and pd.isna(valor)):
        return ""
    try:
        return formatar_numero_br(valor, decimais) + "%"
    except (ValueError, TypeError):
        return str(valor)


def get_cor_pavimentacao(nome):
    """Retorna a cor apropriada para o tipo/pavimentação/status."""
    nome = str(nome).upper().strip()
    if ("NÃO" in nome or "NAO" in nome or "NATURAL" in nome or "TERRA" in nome) \
            and ("PAVIMENT" in nome or "NATURAL" in nome or "TERRA" in nome or "LEITO" in nome):
        return "#D32F2F"  # vermelho
    if "ASFALT" in nome or "CBUQ" in nome or "REVESTIMENTO" in nome:
        return "#757575"  # cinza
    if "PARALEL" in nome or "PEDRA" in nome or "PAVIMENT" in nome:
        return "#388E3C"  # verde
    return "#9E9E9E"      # neutro


# ==========================================================================
# LEITURA E PREPARAÇÃO DOS DADOS
# ==========================================================================

def ler_csv(caminho):
    """Lê o CSV tentando diferentes encodings e separadores."""
    tentativas = [
        {"encoding": "utf-8", "sep": ";"},
        {"encoding": "utf-8", "sep": ","},
        {"encoding": "latin-1", "sep": ";"},
        {"encoding": "latin-1", "sep": ","},
    ]
    ultimo_erro = None
    for opc in tentativas:
        try:
            df = pd.read_csv(caminho, **opc)
            if len(df.columns) > 1:
                return df
        except Exception as e:  # noqa: BLE001
            ultimo_erro = e
    # última tentativa: deixar o pandas decidir
    try:
        return pd.read_csv(caminho, encoding="latin-1")
    except Exception as e:  # noqa: BLE001
        raise RuntimeError(f"Não foi possível ler o CSV '{caminho}': {ultimo_erro or e}")


def converter_para_float(valor):
    """Converte valores numéricos em formato BR/US para float."""
    if pd.isna(valor) or valor == "NÃO INFORMADO":
        return 0.0
    s = str(valor).strip().replace(",", ".")
    if s.count(".") > 1:  # remove pontos de milhar (ex.: 1.000.000 -> 1000000.0)
        partes = s.split(".")
        s = "".join(partes[:-1]) + "." + partes[-1]
    try:
        return float(s)
    except ValueError:
        return 0.0


def preparar_dados(df, col_map):
    """Limpa nulos e converte a coluna de comprimento para float."""
    df = df.copy()
    for col in col_map.values():
        if col in df.columns and col != col_map.get("comprimento"):
            df[col] = df[col].fillna("NÃO INFORMADO")

    comp = col_map.get("comprimento")
    if comp and comp in df.columns:
        df[comp] = df[comp].apply(converter_para_float)
    return df


def filtrar_trecho_1(df, col_map):
    """Retorna apenas os trechos = 1 (usado para contar vias únicas)."""
    trecho = col_map.get("trecho")
    if trecho and trecho in df.columns:
        return df[df[trecho].astype(str).str.strip() == "1"].copy()
    return df.copy()


def contar_vias_unicas(df, col_map):
    """Conta vias únicas (considerando apenas trecho = 1)."""
    if not col_map.get("via"):
        return None
    df_t1 = filtrar_trecho_1(df, col_map)
    return df_t1[col_map["via"]].nunique()


# ==========================================================================
# CÁLCULO DOS QUANTITATIVOS (EM PORCENTAGEM)
# ==========================================================================

def distribuicao_percentual(df, col_map, coluna_grupo):
    """
    Monta uma tabela de distribuição com TOTAIS EM PORCENTAGEM.

    Colunas geradas (todas somam 100%):
      - % da Extensão  -> participação na extensão total da malha (km)
      - % dos Trechos  -> participação na quantidade de trechos
    O comprimento usa TODOS os trechos; a contagem de vias é informativa.
    """
    comp = col_map.get("comprimento")
    via = col_map.get("via")

    # Extensão total (todos os trechos) e total de trechos por grupo
    agg = {}
    if comp and comp in df.columns:
        agg[comp] = "sum"
    grupo = df.groupby(coluna_grupo)

    tabela = pd.DataFrame(index=sorted(grupo.groups.keys()))
    tabela.index.name = coluna_grupo

    # Trechos
    tabela["_trechos"] = grupo.size()

    # Extensão (m) -> usada só para calcular %
    if comp and comp in df.columns:
        tabela["_extensao_m"] = grupo[comp].sum()
    else:
        tabela["_extensao_m"] = tabela["_trechos"]  # fallback

    # Vias únicas (trecho = 1) - informativo
    if via and via in df.columns:
        df_t1 = filtrar_trecho_1(df, col_map)
        vias_grupo = df_t1.groupby(coluna_grupo)[via].nunique()
        tabela["_vias"] = vias_grupo
        tabela["_vias"] = tabela["_vias"].fillna(0)

    # Percentuais (somam 100%)
    total_ext = tabela["_extensao_m"].sum() or 1
    total_tre = tabela["_trechos"].sum() or 1
    tabela["% da Extensão"] = (tabela["_extensao_m"] / total_ext * 100).round(2)
    tabela["% dos Trechos"] = (tabela["_trechos"] / total_tre * 100).round(2)

    tabela = tabela.sort_values("% da Extensão", ascending=False)
    return tabela


def gerar_relatorio(df, col_map, nome_municipio):
    """Calcula todos os blocos do relatório e devolve um dicionário."""
    rel = {
        "municipio": nome_municipio,
        "data": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "total_registros": len(df),
    }

    if col_map.get("via"):
        rel["total_vias"] = contar_vias_unicas(df, col_map)
    if col_map.get("trecho"):
        rel["total_trechos"] = len(df)

    comp = col_map.get("comprimento")
    if comp and comp in df.columns:
        total_m = df[comp].sum()
        rel["total_metros"] = total_m
        rel["total_km"] = total_m / 1000

    for chave, campo in [
        ("por_setor", "setor"),
        ("por_bairro", "bairro"),
        ("por_status", "status"),
        ("por_pavimentacao", "pavimentacao"),
    ]:
        coluna = col_map.get(campo)
        if coluna and coluna in df.columns:
            rel[chave] = distribuicao_percentual(df, col_map, coluna)

    # Top 10 vias mais longas (em % da extensão total)
    if col_map.get("via") and comp and comp in df.columns:
        total_m = df[comp].sum() or 1
        top = df.groupby(col_map["via"])[comp].sum().sort_values(ascending=False).head(10)
        top_df = pd.DataFrame({
            "Via": top.index,
            "% da Extensão": (top.values / total_m * 100).round(2),
        })
        top_df.index = range(1, len(top_df) + 1)
        rel["top_vias"] = top_df

    return rel


# ==========================================================================
# TABELAS "AMIGÁVEIS" (somente % + total = 100%)
# ==========================================================================

def tabela_exibicao(tabela_interna, incluir_vias=True):
    """
    Converte a tabela interna em uma tabela limpa para exibição,
    contendo apenas porcentagens e uma linha de TOTAL = 100%.
    """
    cols = ["% da Extensão", "% dos Trechos"]
    out = tabela_interna[cols].copy()

    if incluir_vias and "_vias" in tabela_interna.columns:
        out.insert(0, "Vias (qtd)", tabela_interna["_vias"].astype(int))

    # Linha de total: porcentagens somam 100% por definição (evita
    # artefatos de arredondamento tipo 100,01%).
    total = {}
    if "Vias (qtd)" in out.columns:
        total["Vias (qtd)"] = int(out["Vias (qtd)"].sum())
    total["% da Extensão"] = 100.0
    total["% dos Trechos"] = 100.0
    out.loc["TOTAL"] = total
    return out


def formatar_tabela_br(tabela):
    """Aplica formatação BR (porcentagens com vírgula) para exibição/export."""
    out = tabela.copy()
    for col in out.columns:
        if col.startswith("%"):
            out[col] = out[col].apply(lambda v: formatar_pct(v, 2))
        elif "qtd" in col.lower() or "Vias" in col:
            out[col] = out[col].apply(formatar_inteiro_br)
    return out


# ==========================================================================
# GRÁFICOS (EM PORCENTAGEM) - SALVOS INDIVIDUALMENTE
# ==========================================================================

def _novo_fig(figsize=(12, 7)):
    fig = plt.figure(figsize=figsize)
    fig.patch.set_facecolor("white")
    return fig


def grafico_barra_pizza(tabela, titulo, nome_arquivo, pasta, cores=None):
    """Gera barras horizontais (%) + pizza (%) lado a lado."""
    dados = tabela["% da Extensão"].sort_values(ascending=True)
    if dados.empty:
        return None

    if cores is None:
        cores = [get_cor_pavimentacao(i) for i in dados.index] \
            if any(get_cor_pavimentacao(i) != "#9E9E9E" for i in dados.index) \
            else (CORES_PRINCIPAIS * (len(dados) // len(CORES_PRINCIPAIS) + 1))[:len(dados)]

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 6))
    fig.patch.set_facecolor("white")

    bars = ax1.barh(dados.index.astype(str), dados.values, color=cores)
    ax1.set_xlabel("Participação na extensão (%)", fontsize=12, fontweight="bold")
    ax1.set_title(f"{titulo} (% da malha)", fontsize=14, fontweight="bold", color="#2C2C2C")
    ax1.grid(axis="x", alpha=0.3, color="#9E9E9E")
    ax1.set_facecolor("#FAFAFA")
    ax1.set_xlim(0, max(dados.values) * 1.18 if len(dados) else 1)
    ax1.bar_label(bars, padding=4,
                  labels=[formatar_pct(v, 1) for v in dados.values],
                  fontweight="bold", fontsize=9)

    dados_pizza = tabela["% da Extensão"].sort_values(ascending=False)
    cores_pizza = cores[::-1] if isinstance(cores, list) else cores
    wedges, _texts, autotexts = ax2.pie(
        dados_pizza.values,
        labels=dados_pizza.index.astype(str),
        autopct=lambda pct: formatar_pct(pct, 1),
        colors=cores_pizza,
        startangle=90,
        textprops={"fontweight": "bold", "fontsize": 9},
    )
    ax2.set_title(f"{titulo} - Distribuição %", fontsize=14, fontweight="bold", color="#2C2C2C")
    for at in autotexts:
        at.set_color("white")
        at.set_fontweight("bold")

    plt.tight_layout()
    destino = os.path.join(pasta, nome_arquivo)
    plt.savefig(destino, dpi=300, bbox_inches="tight", facecolor="white")
    plt.close()
    return destino


def grafico_barra_simples(tabela, titulo, nome_arquivo, pasta, top=15):
    """Barras horizontais (%) - usado para bairros (Top N)."""
    dados = tabela["% da Extensão"].head(top).sort_values(ascending=True)
    if dados.empty:
        return None

    n = len(dados)
    cores = (CORES_PRINCIPAIS * (n // len(CORES_PRINCIPAIS) + 1))[:n]

    fig, ax = plt.subplots(figsize=(14, 8))
    fig.patch.set_facecolor("white")
    bars = ax.barh(dados.index.astype(str), dados.values, color=cores)
    ax.set_xlabel("Participação na extensão (%)", fontsize=12, fontweight="bold")
    ax.set_title(f"{titulo} (% da malha)", fontsize=14, fontweight="bold", color="#2C2C2C")
    ax.grid(axis="x", alpha=0.3, color="#9E9E9E")
    ax.set_facecolor("#FAFAFA")
    ax.set_xlim(0, max(dados.values) * 1.18)
    ax.bar_label(bars, padding=4,
                 labels=[formatar_pct(v, 2) for v in dados.values],
                 fontweight="bold", fontsize=9, color="#2C2C2C")
    plt.tight_layout()
    destino = os.path.join(pasta, nome_arquivo)
    plt.savefig(destino, dpi=300, bbox_inches="tight", facecolor="white")
    plt.close()
    return destino


def grafico_empilhado_pct(df, col_map, linha, coluna, titulo, nome_arquivo, pasta):
    """Barras empilhadas em % do total da malha (soma geral = 100%)."""
    comp = col_map.get("comprimento")
    if not (col_map.get(linha) and col_map.get(coluna) and comp and comp in df.columns):
        return None

    pivot = df.groupby([col_map[linha], col_map[coluna]])[comp].sum().unstack(fill_value=0)
    total = pivot.values.sum() or 1
    pivot_pct = pivot / total * 100

    fig, ax = plt.subplots(figsize=(14, 8))
    fig.patch.set_facecolor("white")
    cores = (CORES_PRINCIPAIS * (len(pivot_pct.columns) // len(CORES_PRINCIPAIS) + 1))[:len(pivot_pct.columns)]
    pivot_pct.plot(kind="bar", stacked=True, ax=ax, color=cores)
    ax.set_xlabel(col_map[linha], fontsize=12, fontweight="bold", color="#2C2C2C")
    ax.set_ylabel("Participação na extensão total (%)", fontsize=12, fontweight="bold", color="#2C2C2C")
    ax.set_title(titulo, fontsize=14, fontweight="bold", color="#2C2C2C")
    ax.legend(title=col_map[coluna], bbox_to_anchor=(1.02, 1), loc="upper left",
              frameon=True, facecolor="white")
    ax.grid(axis="y", alpha=0.3, color="#9E9E9E")
    ax.set_facecolor("#FAFAFA")
    plt.xticks(rotation=45, ha="right")
    plt.tight_layout()
    destino = os.path.join(pasta, nome_arquivo)
    plt.savefig(destino, dpi=300, bbox_inches="tight", facecolor="white")
    plt.close()
    return destino


def grafico_heatmap_pct(df, col_map, linha, coluna, titulo, nome_arquivo, pasta):
    """Mapa de calor em % do total da malha."""
    comp = col_map.get("comprimento")
    if not (col_map.get(linha) and col_map.get(coluna) and comp and comp in df.columns):
        return None
    try:
        import seaborn as sns
    except ImportError:
        return None

    pivot = df.groupby([col_map[linha], col_map[coluna]])[comp].sum().unstack(fill_value=0)
    total = pivot.values.sum() or 1
    pivot_pct = (pivot / total * 100).round(2)

    fig, ax = plt.subplots(figsize=(12, max(6, len(pivot_pct) * 0.4)))
    fig.patch.set_facecolor("white")
    sns.heatmap(pivot_pct, annot=True, fmt=".2f", cmap="YlOrBr", ax=ax,
                cbar_kws={"label": "% da extensão total"},
                linewidths=0.5, linecolor="white")
    ax.set_title(titulo, fontsize=14, fontweight="bold", color="#2C2C2C")
    ax.set_xlabel(col_map[coluna], fontsize=12, fontweight="bold", color="#2C2C2C")
    ax.set_ylabel(col_map[linha], fontsize=12, fontweight="bold", color="#2C2C2C")
    plt.xticks(rotation=45, ha="right")
    plt.yticks(rotation=0)
    plt.tight_layout()
    destino = os.path.join(pasta, nome_arquivo)
    plt.savefig(destino, dpi=300, bbox_inches="tight", facecolor="white")
    plt.close()
    return destino


def gerar_graficos(df, rel, col_map, pasta_graficos):
    """Gera todos os gráficos individuais e devolve a lista de arquivos."""
    os.makedirs(pasta_graficos, exist_ok=True)
    gerados = []

    if "por_setor" in rel:
        f = grafico_barra_pizza(rel["por_setor"], "Vias por Setor",
                                "fig_01_setores.png", pasta_graficos)
        if f:
            gerados.append(f)

    if "por_bairro" in rel:
        f = grafico_barra_simples(rel["por_bairro"], "Top 15 Bairros",
                                  "fig_02_bairros.png", pasta_graficos, top=15)
        if f:
            gerados.append(f)

    if "por_status" in rel:
        cores = [get_cor_pavimentacao(i) for i in
                 rel["por_status"]["% da Extensão"].sort_values(ascending=True).index]
        f = grafico_barra_pizza(rel["por_status"], "Status das Vias",
                                "fig_03_status.png", pasta_graficos, cores=cores)
        if f:
            gerados.append(f)

    if "por_pavimentacao" in rel:
        cores = [get_cor_pavimentacao(i) for i in
                 rel["por_pavimentacao"]["% da Extensão"].sort_values(ascending=True).index]
        f = grafico_barra_pizza(rel["por_pavimentacao"], "Tipo de Pavimentação",
                                "fig_04_pavimentacao.png", pasta_graficos, cores=cores)
        if f:
            gerados.append(f)

    f = grafico_empilhado_pct(df, col_map, "setor", "status",
                              "Setor x Status (% da malha)",
                              "fig_05_setor_status.png", pasta_graficos)
    if f:
        gerados.append(f)

    f = grafico_heatmap_pct(df, col_map, "setor", "pavimentacao",
                            "Mapa de Calor - Setor x Pavimentação (% da malha)",
                            "fig_06_heatmap_setor_pavimentacao.png", pasta_graficos)
    if f:
        gerados.append(f)

    # Treemap interativo (opcional, requer plotly)
    if col_map.get("setor") and col_map.get("bairro") and col_map.get("comprimento"):
        try:
            import plotly.express as px
            comp = col_map["comprimento"]
            total_m = df[comp].sum() or 1
            cross = df.groupby([col_map["setor"], col_map["bairro"]])[comp].sum().reset_index()
            cross.columns = ["Setor", "Bairro", "Extensao_m"]
            cross["% da Extensão"] = (cross["Extensao_m"] / total_m * 100).round(2)
            fig = px.treemap(
                cross, path=["Setor", "Bairro"], values="% da Extensão",
                title="Distribuição da malha - Setor e Bairro (% da extensão total)",
                color="% da Extensão",
                color_continuous_scale=[
                    [0, "#FFF59D"], [0.25, "#FFEB3B"], [0.5, "#FFC107"],
                    [0.75, "#FFB300"], [1, "#FF8F00"]],
            )
            fig.update_layout(height=600, paper_bgcolor="white", plot_bgcolor="white")
            fig.update_traces(marker=dict(line=dict(color="white", width=2)))
            destino = os.path.join(pasta_graficos, "treemap_setor_bairro.html")
            fig.write_html(destino)
            gerados.append(destino)
        except Exception as e:  # noqa: BLE001
            print(f"   (treemap interativo não gerado: {e})")

    return gerados


# ==========================================================================
# EXPORTAÇÃO PARA EXCEL
# ==========================================================================

def exportar_excel(df, rel, col_map, caminho):
    """Gera o relatório Excel com abas em porcentagem."""
    with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
        # Resumo executivo
        resumo = {"Indicador": [], "Valor": []}

        def add(ind, val):
            resumo["Indicador"].append(ind)
            resumo["Valor"].append(val)

        add("Município", rel["municipio"])
        add("Data do Relatório", rel["data"])
        add("Total de Registros (Trechos)", formatar_inteiro_br(rel["total_registros"]))
        if "total_vias" in rel:
            add("Total de Vias", formatar_inteiro_br(rel["total_vias"]))
        if "total_trechos" in rel:
            add("Total de Trechos", formatar_inteiro_br(rel["total_trechos"]))
        if "total_km" in rel:
            add("Extensão Total (km)", formatar_numero_br(rel["total_km"], 3))
            add("Extensão Total (m)", formatar_numero_br(rel["total_metros"], 2))
        pd.DataFrame(resumo).to_excel(writer, sheet_name="Resumo", index=False)

        # Distribuições (% + total 100%)
        abas = [
            ("por_setor", "Por Setor (%)"),
            ("por_bairro", "Por Bairro (%)"),
            ("por_status", "Por Status (%)"),
            ("por_pavimentacao", "Por Pavimentação (%)"),
        ]
        for chave, aba in abas:
            if chave in rel:
                tab = formatar_tabela_br(tabela_exibicao(rel[chave]))
                tab.to_excel(writer, sheet_name=aba[:31])

        if "top_vias" in rel:
            top = rel["top_vias"].copy()
            top["% da Extensão"] = top["% da Extensão"].apply(lambda v: formatar_pct(v, 2))
            top.to_excel(writer, sheet_name="Top 10 Vias (%)", index=False)


# ==========================================================================
# EXPORTAÇÃO PARA PDF
# ==========================================================================

def _pagina_tabela(pdf, titulo, df_show, municipio, rodape_extra=None):
    fig = plt.figure(figsize=(11, 8.5))
    fig.patch.set_facecolor("white")
    ax = fig.add_subplot(111)
    ax.axis("off")

    fig.text(0.5, 0.95, titulo, ha="center", fontsize=16, fontweight="bold", color="#2C2C2C")
    fig.text(0.5, 0.925, "─" * 70, ha="center", fontsize=10, color="#FFD700")

    data = [df_show.columns.tolist()] + df_show.astype(str).values.tolist()
    n_rows, n_cols = len(data), len(df_show.columns)

    table_ax = fig.add_axes([0.08, 0.12, 0.84, 0.78])
    table_ax.axis("off")
    tbl = table_ax.table(cellText=data, cellLoc="center", loc="center")
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(8 if n_cols > 5 else 9)
    tbl.scale(1, 1.7)

    for j in range(n_cols):
        tbl[(0, j)].set_facecolor("#FFD700")
        tbl[(0, j)].set_text_props(weight="bold", color="#2C2C2C")
    for i in range(1, n_rows):
        is_total = str(data[i][0]).upper() == "TOTAL"
        for j in range(n_cols):
            tbl[(i, j)].set_facecolor("#FFF3C4" if is_total else ("#FAFAFA" if i % 2 == 0 else "white"))
            tbl[(i, j)].set_edgecolor("#D0D0D0")
            if is_total:
                tbl[(i, j)].set_text_props(weight="bold", color="#2C2C2C")

    if rodape_extra:
        fig.text(0.5, 0.07, rodape_extra, ha="center", fontsize=8, color="#757575", style="italic")
    fig.text(0.5, 0.03, f'{municipio} - {datetime.now().strftime("%Y")}',
             ha="center", fontsize=9, color="#9E9E9E", style="italic")
    pdf.savefig(fig, bbox_inches="tight", facecolor="white")
    plt.close()


def exportar_pdf(rel, graficos, caminho, logo_path=None):
    """Gera o relatório PDF (capa + tabelas em % + gráficos)."""
    municipio = rel["municipio"]
    with PdfPages(caminho) as pdf:
        # Capa
        fig = plt.figure(figsize=(8.5, 11))
        fig.patch.set_facecolor("white")
        ax = fig.add_subplot(111)
        ax.axis("off")
        if logo_path and os.path.exists(logo_path):
            try:
                from PIL import Image
                img = Image.open(logo_path)
                if img.mode == "RGBA":
                    bg = Image.new("RGB", img.size, (255, 255, 255))
                    bg.paste(img, mask=img.split()[3])
                    img = bg
                elif img.mode != "RGB":
                    img = img.convert("RGB")
                h = 0.15
                w = h * (img.width / img.height) * (11 / 8.5)
                ax_logo = fig.add_axes([(1 - w) / 2, 0.75, w, h])
                ax_logo.imshow(img)
                ax_logo.axis("off")
            except Exception as e:  # noqa: BLE001
                print(f"   (logomarca não inserida: {e})")

        fig.text(0.5, 0.62, "RELATÓRIO DE LOGRADOUROS", ha="center",
                 fontsize=26, fontweight="bold", color="#2C2C2C")
        fig.text(0.5, 0.57, "Análise Quantitativa da Malha Viária Municipal",
                 ha="center", fontsize=14, color="#424242")
        fig.text(0.5, 0.53, "─" * 40, ha="center", fontsize=12, color="#FFD700")
        fig.text(0.5, 0.45, municipio, ha="center", fontsize=20,
                 fontweight="bold", color="#2C2C2C")
        fig.text(0.5, 0.40, rel["data"], ha="center", fontsize=12,
                 style="italic", color="#616161")
        fig.text(0.5, 0.06, "EIXO - Soluções em Gestão Pública", ha="center",
                 fontsize=11, color="#757575", style="italic")
        pdf.savefig(fig, bbox_inches="tight", facecolor="white")
        plt.close()

        # Quadro resumo
        resumo_rows = [
            ("Total de Registros (Trechos)", formatar_inteiro_br(rel["total_registros"])),
        ]
        if "total_vias" in rel:
            resumo_rows.append(("Total de Vias", formatar_inteiro_br(rel["total_vias"])))
        if "total_trechos" in rel:
            resumo_rows.append(("Total de Trechos", formatar_inteiro_br(rel["total_trechos"])))
        if "total_km" in rel:
            resumo_rows.append(("Extensão Total (km)", formatar_numero_br(rel["total_km"], 3)))
        df_resumo = pd.DataFrame(resumo_rows, columns=["Indicador", "Valor"])
        _pagina_tabela(pdf, "QUADRO RESUMO", df_resumo, municipio)

        # Tabelas de distribuição (em %)
        for chave, titulo in [
            ("por_setor", "DISTRIBUIÇÃO POR SETOR (%)"),
            ("por_bairro", "DISTRIBUIÇÃO POR BAIRRO (%)"),
            ("por_status", "DISTRIBUIÇÃO POR STATUS (%)"),
            ("por_pavimentacao", "DISTRIBUIÇÃO POR PAVIMENTAÇÃO (%)"),
        ]:
            if chave in rel:
                tab = tabela_exibicao(rel[chave])
                total = len(tab) - 1  # tira a linha TOTAL
                rodape = None
                if total > 24:
                    # mantém TOTAL no fim
                    corpo = formatar_tabela_br(tab.iloc[:-1]).head(24)
                    total_fmt = formatar_tabela_br(tab.iloc[[-1]])
                    tab_fmt = pd.concat([corpo, total_fmt])
                    rodape = f"* Exibindo as 24 maiores de {total} categorias (TOTAL considera todas)"
                else:
                    tab_fmt = formatar_tabela_br(tab)
                _pagina_tabela(pdf, titulo, tab_fmt.reset_index(), municipio, rodape)

        # Top 10 vias
        if "top_vias" in rel:
            top = rel["top_vias"].copy()
            top["% da Extensão"] = top["% da Extensão"].apply(lambda v: formatar_pct(v, 2))
            top = top.reset_index().rename(columns={"index": "#"})
            _pagina_tabela(pdf, "TOP 10 VIAS MAIS LONGAS (% da malha)", top, municipio)

        # Gráficos (apenas PNG)
        for img_file in sorted(g for g in graficos if g.lower().endswith(".png")):
            try:
                img = plt.imread(img_file)
                fig = plt.figure(figsize=(11, 8.5))
                ax = fig.add_subplot(111)
                ax.imshow(img)
                ax.axis("off")
                pdf.savefig(fig, bbox_inches="tight")
                plt.close()
            except Exception as e:  # noqa: BLE001
                print(f"   (gráfico não incluído no PDF: {img_file}: {e})")


# ==========================================================================
# UTILIDADES DE ENTRADA
# ==========================================================================

def descobrir_csv(arg_csv):
    """Determina o caminho do CSV (argumento, autodetecção ou pergunta)."""
    if arg_csv:
        if not os.path.exists(arg_csv):
            sys.exit(f"❌ Arquivo não encontrado: {arg_csv}")
        return arg_csv

    csvs = sorted(glob.glob("*.csv"))
    if len(csvs) == 1:
        print(f"📁 CSV encontrado automaticamente: {csvs[0]}")
        return csvs[0]
    if len(csvs) > 1:
        print("📁 Vários CSVs encontrados:")
        for i, c in enumerate(csvs, 1):
            print(f"   {i}. {c}")
        escolha = input("Digite o número do arquivo desejado: ").strip()
        try:
            return csvs[int(escolha) - 1]
        except (ValueError, IndexError):
            sys.exit("❌ Escolha inválida.")
    # nenhum encontrado
    caminho = input("Digite o caminho do arquivo CSV: ").strip().strip('"')
    if not os.path.exists(caminho):
        sys.exit(f"❌ Arquivo não encontrado: {caminho}")
    return caminho


def validar_colunas(df, col_map):
    """Remove do mapa as colunas que não existem no CSV e avisa."""
    presentes = {}
    faltando = []
    for chave, nome in col_map.items():
        if nome in df.columns:
            presentes[chave] = nome
        else:
            faltando.append(nome)
    if faltando:
        print("⚠️  Colunas não encontradas no CSV (serão ignoradas): "
              + ", ".join(faltando))
        print(f"   Colunas disponíveis: {', '.join(df.columns)}")
    return presentes


# ==========================================================================
# MAIN
# ==========================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Gera relatório da malha viária (Excel, PDF e gráficos) em porcentagem.")
    parser.add_argument("csv", nargs="?", help="Caminho do arquivo CSV (opcional).")
    parser.add_argument("--municipio", help="Nome do município (se omitido, será perguntado).")
    parser.add_argument("--saida", default="relatorio_saida",
                        help="Pasta de saída (padrão: relatorio_saida).")
    parser.add_argument("--logo", default=LOGO_PADRAO,
                        help=f"Caminho da logomarca (padrão: {LOGO_PADRAO}).")
    args = parser.parse_args()

    print("=" * 60)
    print("📊 RELATÓRIO DE MALHA VIÁRIA - VERSÃO LOCAL")
    print("=" * 60)

    caminho_csv = descobrir_csv(args.csv)
    nome_municipio = args.municipio or input("\n🏙️  Digite o nome do município: ").strip() or "Município"

    print("\n🔄 Lendo dados...")
    df = ler_csv(caminho_csv)
    print(f"   ✓ {formatar_inteiro_br(len(df))} registros, {len(df.columns)} colunas")

    col_map = validar_colunas(df, COLUNAS_PADRAO)
    df = preparar_dados(df, col_map)

    print("📐 Calculando quantitativos (em porcentagem)...")
    rel = gerar_relatorio(df, col_map, nome_municipio)

    # Pastas de saída
    os.makedirs(args.saida, exist_ok=True)
    pasta_graficos = os.path.join(args.saida, PASTA_GRAFICOS)

    print("📊 Gerando gráficos individuais...")
    graficos = gerar_graficos(df, rel, col_map, pasta_graficos)
    for g in graficos:
        print(f"   ✓ {g}")

    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    base = f"{nome_municipio.replace(' ', '_')}_{stamp}"
    arquivo_excel = os.path.join(args.saida, f"Relatorio_Logradouros_{base}.xlsx")
    arquivo_pdf = os.path.join(args.saida, f"Relatorio_Completo_{base}.pdf")

    print("📑 Gerando Excel...")
    exportar_excel(df, rel, col_map, arquivo_excel)
    print(f"   ✓ {arquivo_excel}")

    print("📄 Gerando PDF...")
    logo = args.logo if args.logo and os.path.exists(args.logo) else None
    exportar_pdf(rel, graficos, arquivo_pdf, logo_path=logo)
    print(f"   ✓ {arquivo_pdf}")

    # Compacta os gráficos em zip (entregues separadamente)
    zip_graficos = os.path.join(args.saida, f"graficos_{base}")
    shutil.make_archive(zip_graficos, "zip", pasta_graficos)
    print(f"   ✓ {zip_graficos}.zip")

    print("\n" + "=" * 60)
    print("✅ RELATÓRIO CONCLUÍDO!")
    print("=" * 60)
    print(f"   • Município: {nome_municipio}")
    print(f"   • Registros: {formatar_inteiro_br(rel['total_registros'])}")
    if "total_vias" in rel:
        print(f"   • Vias: {formatar_inteiro_br(rel['total_vias'])}")
    if "total_km" in rel:
        print(f"   • Extensão total: {formatar_numero_br(rel['total_km'], 3)} km")
    print(f"\n📁 Arquivos em: {os.path.abspath(args.saida)}")
    print(f"   • Excel:    {os.path.basename(arquivo_excel)}")
    print(f"   • PDF:      {os.path.basename(arquivo_pdf)}")
    print(f"   • Gráficos: {PASTA_GRAFICOS}/ (e .zip)")
    print("=" * 60)


if __name__ == "__main__":
    main()
