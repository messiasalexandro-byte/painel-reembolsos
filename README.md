# Painel de Reembolsos

Painel executivo para acompanhar **estornos, cupons e chargebacks** a partir da planilha de controle de reembolsos.

| Quero… | Vá para |
|---|---|
| ver o painel com os dados atuais | [Acessar o painel](#acessar-o-painel) |
| saber onde está cada informação | [Navegação](#navegação) |
| preparar ou corrigir a planilha | [Formato esperado da planilha](#formato-esperado-da-planilha) |
| mudar o painel e publicar | [Para quem mantém](#para-quem-mantém) |

---

## Acessar o painel

### 1. Painel online (recomendado)

**https://script.google.com/macros/s/AKfycbxgwqnN5kGArZmYM-zbGyV_h4AjfkzqITj2ZTfVP1PfpZMDMu2Ag8FDnlDpEGiWbXVy/exec**

Abre já com os dados da planilha que está na pasta do Google Drive. Não é preciso importar nada.

- **Quem acessa:** só quem está logado com uma conta **@talgui.com.br**.
- **Atualização:** com o painel aberto, ele verifica a pasta **a cada 5 minutos** e só redesenha se a planilha mudou, mantendo o mês, os filtros e a página abertos. **Atualizar agora**, no menu lateral, força a verificação.
- **Qual arquivo:** a Planilha Google editada por último na pasta (exportada como `.xlsx`). Se não houver nenhuma Planilha Google, usa o `.xlsx` modificado por último. Subpastas são ignoradas, então dá para guardar versões antigas numa subpasta "Arquivo".
- **Se a versão nova tiver erro de estrutura**, o painel mantém os dados anteriores e avisa no menu lateral.
- **Várias contas Google no mesmo navegador:** se o painel não carregar, abra numa janela anônima só com a conta da Talgui. É uma limitação conhecida do Google Apps Script.

### 2. Importação manual (GitHub Pages)

**https://messiasalexandro-byte.github.io/painel-reembolsos/** (ou o arquivo `index.html` aberto no navegador)

1. Arraste a planilha `.xlsx` para a área de importação ou clique para selecioná-la.
2. Revise o **diagnóstico da importação**: abas encontradas, linhas lidas, válidas e excluídas, erros e avisos.
3. Clique em **Continuar para o Dashboard**.

Na próxima visita, o painel oferece **"Continuar com <arquivo>"**. **"Esquecer arquivo"** apaga essa cópia do navegador. A tela de importação também mostra o link do painel online.

## Navegação

O menu lateral agrupa as páginas pelo que a pessoa quer fazer. Cada página tem endereço próprio, então dá para salvar nos favoritos ou mandar o link de uma página específica.

| Grupo | Página | Endereço | Para quê | Filtros que valem |
|---|---|---|---|---|
| **Análise** | Visão geral | `#/visao-geral` | Situação do mês em segundos: KPIs, resumo por mês (clique para trocar o mês), indicadores-chave e maiores exposições diárias | mês, métrica, categoria |
| | Diário e semanal | `#/diario-semanal` | Comportamento dia a dia, comparativo semanal e últimos 7 dias | mês, métrica, semana, categoria |
| | Mês a mês | `#/mensal` | Trajetória do ano, acumulado e composição por categoria | métrica, categoria |
| | Canais e status | `#/canais` | Distribuição por canal (SAC, TD Nuvem, PIX Direto) e por status | mês, métrica, categoria |
| **Rotina** | Relatório diário | `#/relatorio` | Totais de um dia escolhido, comparados à média e ao dia anterior | categoria |
| | Metas e projeções | `#/metas` | Teto do mês, previsão de fechamento e projeções de estornos × realizado | mês |
| **Dados** | Lançamentos | `#/lancamentos` | Tabela linha a linha, com busca, filtros próprios e CSV | mês, semana, categoria |
| | Diagnóstico da planilha | `#/diagnostico` | Resultado da leitura e da validação da planilha | — |

- A barra de filtros mostra **só os filtros que mudam a página aberta**. Os outros ficam guardados e voltam a valer quando você abre uma página que os usa.
- O subtítulo abaixo do nome da página diz sempre qual recorte está na tela (mês, semana, categoria).
- Alguns quadros têm recorte fixo, avisado no canto do quadro: **Últimos 7 dias** (até o último lançamento), **Composição por categoria** (ano inteiro, todas as categorias) e **Metas** (em R$, todas as categorias).
- No celular, o menu abre pelo botão ☰. **Imprimir / PDF** imprime a página aberta. **Exportar Excel** fica no topo de todas as páginas.

## Funcionalidades

- **KPIs**: volume, impacto financeiro, ticket médio e participação no ano, com variação sobre o mês anterior e minigráfico anual.
- **Resumo por mês**: cartões clicáveis para trocar o mês. Mostram "▲ acima da meta" quando o mês passa do teto definido.
- **Comportamento diário**: barras por dia, média e acumulado. Opção **Comparar com mês anterior** (linha tracejada). A semana escolhida no filtro fica destacada.
- **Comparativo semanal**: semanas fixas (dias 1–7, 8–14, 15–21, 22–28 e 29–31), com o intervalo de dias no próprio filtro.
- **Trajetória mensal**, **acumulado no ano** e **composição por categoria** (tabela do ano inteiro).
- **Relatório diário**: totais do dia por natureza, representatividade no mês, desvio da média e variação sobre o dia anterior.
- **Metas e projeções**:
  - teto mensal em R$ com barra de progresso, previsão de fechamento pelo ritmo diário e alerta quando a meta é ou deve ser ultrapassada;
  - **projeções de estornos**: registre o valor projetado de um dia (até 3 casas decimais, ex.: `R$ 123,456`) e compare com o realizado. O restante do painel usa centavos.
- **Lançamentos**: busca; filtros de canal, status e "somente com avisos"; ordenação por coluna, paginação e total filtrado.
- **Exportar Excel**: 4 abas (lançamentos filtrados, resumo mensal, diário do mês, ocorrências do diagnóstico).
- **Baixar CSV**: dos lançamentos, no padrão brasileiro (`;` e vírgula decimal).
- **Imprimir / PDF**: sempre em tema claro, sem botões e com as tabelas abertas.
- **Tema claro/escuro**: a escolha fica salva.

A previsão de fechamento e os "últimos 7 dias" usam a data do último lançamento da planilha como "hoje".

## Formato esperado da planilha

| Item | Regra |
|---|---|
| Abas mensais | Uma aba por mês, reconhecida pelo nome (`Janeiro`…`Dezembro`, ou abreviado: `Jan`, `Fev`…). Maiúsculas e acentos são ignorados. |
| Colunas obrigatórias | `Data`, `Tipo de Ocorrência`, `Quantidade`, `Valor`. O cabeçalho pode estar em qualquer uma das primeiras 150 linhas. |
| Colunas opcionais | `Pedido/Referência`, `Status`, `Observação` |
| Tipos reconhecidos | `Estorno`, `Cupom`, `Chargeback` |
| Ano de referência | Célula `B2` da aba `Consolidado Anual`. Sem ela, usa a mediana dos anos das datas (com 5 datas ou mais) ou, em último caso, o ano atual. |
| Canal | Detectado no texto da `Observação`: `SAC`, `TD Nuvem` ou `PIX Direto` |

A leitura termina após duas linhas totalmente vazias. Linhas de "Total" são ignoradas.

### Validações do diagnóstico

Nada é descartado sem aviso. Cada ocorrência mostra aba, linha, campo e o valor encontrado.

- **Erro (linha excluída dos cálculos):** data inválida ou vazia; tipo, quantidade ou valor ausentes ou não numéricos.
- **Aviso (linha mantida):** tipo fora de Estorno/Cupom/Chargeback, quantidade ou valor negativos, data fora do ano de referência, data de outro mês que não o da aba, possível lançamento duplicado.
- **Bloqueio:** nenhuma aba mensal reconhecida, nenhuma tabela válida ou nenhuma linha válida.

## Privacidade e armazenamento

- Na importação manual, a planilha **não é enviada** para nenhum servidor: tudo acontece no navegador.
- No painel online, a planilha vai do Google Drive direto para o navegador, pelo Apps Script, sem passar por outros servidores. Só contas @talgui.com.br conseguem abrir.
- Algumas informações ficam salvas **só no navegador de cada pessoa**, sem compartilhamento entre computadores:
  - último arquivo importado ou baixado do Drive (IndexedDB `reembolsos_db`);
  - metas mensais (`localStorage`: `reembolsos_metas_v1`);
  - projeções (`localStorage`: `reembolsos_projecoes_v1`);
  - tema (`localStorage`: `reembolsos_tema_v1`).
- Limpar os dados do site no navegador apaga essas informações.

---

## Para quem mantém

### Estrutura dos arquivos

```
painel-reembolsos/
├── index.html            ← o painel inteiro (fonte única: edite só este)
├── README.md
├── .claude/agents/       ← agentes do Claude Code (arquiteto, auditor, redator, revisor)
└── apps-script/          ← projeto do Google Apps Script que serve o painel online
    ├── Codigo.gs         ← doGet() entrega a página; obterPlanilha(versao) devolve a planilha do Drive
    ├── appsscript.json   ← escopos e acesso restrito ao domínio (DOMAIN)
    ├── .clasp.json       ← scriptId usado pelo clasp
    └── publicar.sh       ← publica o index.html no Apps Script, mesma URL
```

- `apps-script/Index.html` **não é mantido à mão**: o `publicar.sh` o gera copiando o `index.html` a cada publicação (e o `.gitignore` o exclui do Git). Não edite esse arquivo.
- O mesmo `index.html` serve os dois lugares: ele detecta se está dentro do Apps Script (`google.script.run`) e, nesse caso, busca a planilha no Drive e verifica mudanças a cada 5 minutos.
- A pasta do Drive é definida em `FOLDER_ID`, no `Codigo.gs`. Ela continua privada: o script lê o arquivo em nome da conta dona.

### Publicar mudanças

São dois destinos, publicados separadamente:

| Destino | Como publicar |
|---|---|
| Painel online (Apps Script) | `sh apps-script/publicar.sh` — copia `index.html` para `apps-script/Index.html` e roda `clasp push` + `clasp update-deployment`, mantendo a URL. Antes, faça `npx @google/clasp login` com a conta dona do script, marcando **todas** as permissões. O login expira de tempos em tempos (erro `invalid_rapt`); aí é só repetir o login. |
| GitHub Pages | `git push` na branch `main`. O site atualiza em cerca de 1 minuto. |

### Tecnologia

- Um único arquivo, `index.html`, com HTML, CSS e JavaScript. Sem build e sem dependências para instalar.
- [SheetJS](https://sheetjs.com/) 0.18.5 para ler e gerar `.xlsx` e [Chart.js](https://www.chartjs.org/) 4.5.1 para os gráficos, carregados do cdnjs. Por isso, **é preciso internet** para abrir o painel.

O JavaScript está dividido em três blocos dentro do `index.html`:

| Bloco | Responsabilidade |
|---|---|
| `RefundsApp.core` | Leitura da planilha, validação e diagnóstico (`processWorkbook`), formatação |
| `RefundsApp.agg` | Agregações: por mês, dia, semana, categoria, canal/status, previsão, filtros |
| Renderização | Telas, rotas (`ROUTES`, com os filtros de cada página), gráficos, exportação, tema e eventos |

Para criar uma página nova: acrescente a entrada em `ROUTES` (com a lista `filtros`), o link no menu lateral dentro do grupo certo e a `<section class="page" data-page="…">`. Não renomeie os ids das rotas existentes: eles estão nos favoritos das pessoas.

```bash
git clone https://github.com/messiasalexandro-byte/painel-reembolsos.git
cd painel-reembolsos
# abra o index.html no navegador
```
