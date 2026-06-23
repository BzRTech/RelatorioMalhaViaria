# Relatório de Malha Viária (versão local)

Gera relatórios da malha viária municipal a partir de um **CSV** ou de um
**shapefile** (`.shp`) de trechos de logradouros, produzindo:

- 📑 **Excel** (`.xlsx`) com todas as tabelas
- 📄 **PDF** (capa, quadro resumo, tabelas e gráficos)
- 📊 **Gráficos individuais** em PNG (entregues também compactados em `.zip`)
- 🌳 **Treemap interativo** (HTML)

> Esta é a versão **local** do antigo notebook do Google Colab. Roda no seu
> computador, sem precisar de internet e sem o Colab.

## Diferença importante: tudo em porcentagem

Todas as tabelas e gráficos que possuem **totais** são apresentados em
**porcentagem** (participação na malha viária), com a linha de **TOTAL = 100%**.
Isso evita confusão para quem lê o relatório. As tabelas mostram, lado a lado, o **comprimento (km)** e a porcentagem:

- **Extensão (km)** — comprimento total da categoria, em quilômetros
- **% da Extensão** — quanto cada categoria representa da extensão total
- **Trechos (qtd)** e **% dos Trechos** — quantidade e participação dos trechos
- **Vias (qtd)** — quantidade de vias
- A linha de **TOTAL** traz a extensão somada e **100%** nas colunas de porcentagem

Os gráficos de barras exibem a **% e o km** juntos em cada barra
(ex.: `53,6% (60,99 km)`).

A extensão total da malha (em km) continua aparecendo apenas no **Quadro Resumo**,
como número geral de referência.

## Novos insights (denominação das vias)

Além das distribuições, o relatório identifica as **vias sem denominação**
(logradouros cujo nome contém "SEM NOME" ou está em branco) e traz:

- **Denominação geral** — % de vias com nome x sem denominação (vias, trechos e extensão)
- **Bairros com mais vias nominadas** — ranking estilo Centro / Juliana Pires
- **Bairros com mais logradouros sem denominação** — para priorizar ações de nomeação
- **Sem denominação por setor e por pavimentação** — onde se concentram as vias sem nome

> Exemplo real (Tabira/PE): **67,4% das vias** estão sem denominação, e elas se
> concentram no **leito natural (60,7%)** — ou seja, ruas ainda não pavimentadas.

## Formatos de entrada aceitos

| Formato            | Como informar                                  |
|--------------------|------------------------------------------------|
| CSV                | `arquivo.csv`                                  |
| Shapefile          | `arquivo.shp` (use os arquivos `.shp/.dbf/...` juntos na mesma pasta) |
| Pasta              | pasta contendo um `.shp` ou `.csv`             |
| ZIP                | `.zip` com o shapefile/CSV dentro              |

## Como usar

### 1. Instalar o Python e as dependências

Instale o [Python 3.10+](https://www.python.org/downloads/) e depois rode:

```bash
pip install -r requirements.txt
```

### 2. Colocar o CSV na pasta

Copie o arquivo CSV dos logradouros para a mesma pasta do script. As colunas
devem seguir o padrão abaixo (já é o formato usado normalmente):

| Função        | Coluna no CSV |
|---------------|---------------|
| Logradouro    | `RUA`         |
| Trecho        | `TRECHO`      |
| Setor         | `SETOR`       |
| Bairro        | `BAIRRO`      |
| Status        | `STATUS`      |
| Comprimento   | `COMP_TRECH`  |
| Pavimentação  | `TIPO`        |

> Se algum nome de coluna for diferente, edite o dicionário `COLUNAS_PADRAO`
> no início do arquivo `relatorio_malha_viaria.py`.

### 3. Rodar

```bash
python relatorio_malha_viaria.py
```

O programa **encontra o CSV automaticamente** (se houver só um na pasta) e
pergunta apenas o **nome do município**. Pronto!

### Opções avançadas (opcionais)

```bash
# Informar o arquivo e o município direto (CSV, .shp ou .zip)
python relatorio_malha_viaria.py dados.csv --municipio "Minha Cidade"
python relatorio_malha_viaria.py LOGRADOUROS_OFICIAL.shp --municipio "Tabira"
python relatorio_malha_viaria.py malha.zip --municipio "Tabira"

# Escolher a pasta de saída e a logomarca
python relatorio_malha_viaria.py --saida resultados --logo logomarca-eixo-cores-2.png
```

## Onde ficam os resultados

Tudo é salvo na pasta `relatorio_saida/` (ou na que você indicar):

```
relatorio_saida/
├── Relatorio_Logradouros_<municipio>_<data>.xlsx
├── Relatorio_Completo_<municipio>_<data>.pdf
├── graficos_<municipio>_<data>.zip
└── graficos/
    ├── fig_01_setores.png
    ├── fig_02_bairros.png
    ├── fig_03_status.png
    ├── fig_04_pavimentacao.png
    ├── fig_05_setor_status.png
    ├── fig_06_heatmap_setor_pavimentacao.png
    ├── fig_07_denominacao.png
    ├── fig_08_bairros_sem_denominacao.png
    ├── fig_09_sem_nome_pavimentacao.png
    ├── fig_10_sem_nome_setor.png
    └── treemap_setor_bairro.html
```

## Logomarca (opcional)

Coloque o arquivo `logomarca-eixo-cores-2.png` na pasta do script para que ele
apareça na capa do PDF. Se não houver, o relatório é gerado normalmente sem ela.
