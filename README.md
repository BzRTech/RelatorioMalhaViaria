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
- **Logradouros (qtd)** — *apenas em Status e Pavimentação*: número de ruas
  distintas (cada via classificada pelo seu trecho principal)
- A linha de **TOTAL** traz a extensão somada e **100%** nas colunas de porcentagem

> **Sobre a "quantidade de vias":** o **total geral de vias** do município
> aparece no **Quadro Resumo** (indicador importante). Já **nas distribuições
> por setor/bairro** essa contagem foi removida de propósito: uma mesma via
> pode atravessar vários setores/bairros, então contar "quantas vias" por grupo
> é ambíguo (um bairro pode mostrar "0 vias" e ainda ter trechos, porque o
> trecho 1 da rua está em outro bairro). Por grupo, as medidas **exatas** são a
> **extensão (km)** e a **quantidade de trechos**.

Os gráficos de barras exibem a **% e o km** juntos em cada barra
(ex.: `53,6% (60,99 km)`).

### Unidade de cada gráfico

A unidade foi escolhida pelo que cada gráfico realmente mede:

| Gráfico | Unidade | Por quê |
|---------|---------|---------|
| Setor, Bairro, Setor×Status, Heatmap, Treemap | **% da extensão (km)** | distribuição territorial — cada metro pertence a um só local; contar vias é ambíguo (a via cruza fronteiras) |
| Status, Pavimentação | **% da extensão (km)** | variam ao longo da via (uma rua pode ser meio pavimentada) |
| Denominação geral | **% das vias distintas** | "ter nome" é propriedade do logradouro inteiro |
| Bairros sem denominação | **qtd de logradouros** | ruas distintas sem nome por bairro (cada rua conta 1 vez) |
| Sem denominação por pavimentação | **% de trechos** | fração de trechos sem nome em cada tipo |

Cada gráfico traz uma **legenda discreta no rodapé** explicando a base do número
(ex.: *"Participação na extensão da malha · km do setor ÷ km total"*). A mesma
explicação, em forma de tabela, fica na aba **Notas** do Excel.

> **Bairros sem delimitação:** trechos cujo bairro é um placeholder
> (`SEM DELIMITACAO`, `NÃO INFORMADO` etc.) são **excluídos das análises por
> bairro** (Top Bairros, sem denominação por bairro e treemap), para não competir
> com os bairros reais. Eles continuam contando nos totais gerais e por setor.

### Limpeza automática de categorias

Setor, bairro, status e pavimentação são **unificados automaticamente** quando
aparecem com grafias diferentes só por **acento, maiúscula ou espaços**
(ex.: `ESPIRITO SANTO` e `ESPÍRITO SANTO`, ou `ANTONIO CRISTÓVÃO` e
`ANTÔNIO CRISTÓVÃO` viram uma única categoria). Isso evita que o mesmo bairro
apareça duplicado nos gráficos e nas tabelas. O rótulo exibido é a grafia mais
frequente no arquivo.

A extensão total da malha (em km) continua aparecendo apenas no **Quadro Resumo**,
como número geral de referência. A data do relatório é exibida apenas como
**mês/ano** (ex.: `Junho/2026`).

## Novos insights (denominação das vias)

Além das distribuições, o relatório identifica as **vias sem denominação**
(logradouros cujo nome contém "SEM NOME" ou está em branco) e traz:

- **Denominação geral** — por **vias distintas** (% dos logradouros com nome x sem
  nome), pois "ter nome" é propriedade do logradouro inteiro; a extensão (km)
  aparece ao lado como referência
- **Bairros com mais logradouros sem denominação** — contagem de ruas distintas
  sem nome por bairro (cada rua conta 1 vez, mesmo com vários trechos); mostra
  quantos logradouros precisam de denominação em cada bairro
- **Sem denominação por pavimentação** — % de trechos sem nome em cada tipo

> Exemplo real (Tabira/PE): **62,9% dos logradouros** estão sem denominação
> (equivalente a **44,7% da extensão**, ≈57 km), e essas vias se concentram no
> **leito natural** — ou seja, ruas ainda não pavimentadas.

## Formatos de entrada aceitos

| Formato            | Como informar                                  |
|--------------------|------------------------------------------------|
| CSV                | `arquivo.csv`                                  |
| Shapefile          | `arquivo.shp` (use os arquivos `.shp/.dbf/...` juntos na mesma pasta) |
| Só a tabela (.dbf) | `arquivo.dbf` (lê apenas os atributos, sem geometria) |
| Pasta              | pasta contendo um `.shp`, `.dbf` ou `.csv`     |
| ZIP / RAR          | `.zip` ou `.rar` com o shapefile/CSV dentro    |

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
    └── treemap_setor_bairro.html
```

## Logomarca (opcional)

Coloque o arquivo `logomarca-eixo-cores-2.png` na pasta do script para que ele
apareça na capa do PDF. Se não houver, o relatório é gerado normalmente sem ela.
