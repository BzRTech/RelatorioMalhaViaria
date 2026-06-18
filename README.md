# Relatório de Malha Viária (versão local)

Gera relatórios da malha viária municipal a partir de um arquivo **CSV** de
trechos de logradouros, produzindo:

- 📑 **Excel** (`.xlsx`) com todas as tabelas
- 📄 **PDF** (capa, quadro resumo, tabelas e gráficos)
- 📊 **Gráficos individuais** em PNG (entregues também compactados em `.zip`)
- 🌳 **Treemap interativo** (HTML)

> Esta é a versão **local** do antigo notebook do Google Colab. Roda no seu
> computador, sem precisar de internet e sem o Colab.

## Diferença importante: tudo em porcentagem

Todas as tabelas e gráficos que possuem **totais** são apresentados em
**porcentagem** (participação na malha viária), com a linha de **TOTAL = 100%**.
Isso evita confusão para quem lê o relatório. As tabelas mostram:

- **% da Extensão** — quanto cada categoria representa da extensão total (km)
- **% dos Trechos** — quanto representa da quantidade de trechos
- **Vias (qtd)** — quantidade de vias (valor de contexto)

A extensão total da malha (em km) continua aparecendo apenas no **Quadro Resumo**,
como número geral de referência.

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
# Informar o CSV e o município direto
python relatorio_malha_viaria.py dados.csv --municipio "Minha Cidade"

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
    └── treemap_setor_bairro.html
```

## Logomarca (opcional)

Coloque o arquivo `logomarca-eixo-cores-2.png` na pasta do script para que ele
apareça na capa do PDF. Se não houver, o relatório é gerado normalmente sem ela.
