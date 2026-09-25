# Painel de Reembolsos

Painel executivo para acompanhar **estornos, cupons e chargebacks** a partir da planilha de controle de reembolsos.

| Quero… | Vá para |
|---|---|
| ver o painel com os dados atuais | [Acessar o painel](#acessar-o-painel) |
| saber onde está cada informação | [Navegação](#navegação) |
| preparar ou corrigir a planilha | [Formato esperado da planilha](#formato-esperado-da-planilha) |
| mudar o painel e publicar | [Para quem mantém](#para-quem-mantém) |

### Novidades (25/09/2026)

- **Indicadores revisados:** o canal passou a ser dividido pela abertura da Observação, as linhas de exemplo ficam fora das contas, e as médias e comparações usam só dias já decorridos. Os indicadores têm nomes técnicos (MTD, YTD, MoM, run-rate, σ) e um ícone "i" com a fórmula.
- **Nova estrutura:** Início com alertas, Lançamentos, Análises, Acompanhamento e Planilha. Também chegam busca global (Ctrl+K), clique no gráfico para abrir os lançamentos filtrados, link da visão com os filtros e Exportar ▾.
- **Visual e interação:** seis temas (Automático, Claro, Ameixa, Corporativo, Grafite e Alto contraste), gestos no celular, animações suaves e impressão compacta.

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

A estrutura segue o padrão do painel administrativo da Nuvemshop: **Início** e **Lançamentos** no topo do menu, as análises agrupadas por pergunta e a **Planilha** (configuração e diagnóstico) no rodapé. A proposta completa está em `../relatorios/arquitetura-informacao-2026-09-25.md`.

| Grupo | Página | Endereço | Para quê | Filtros que valem |
|---|---|---|---|---|
| — | Início | `#/inicio` | Situação do mês em segundos: alertas (meta, dias atípicos, avisos da planilha, o que mudou), KPIs, meta do mês, últimos 7 dias, top 5 dias e resumo por mês | mês, métrica, categoria |
| — | Lançamentos | `#/lancamentos` | Lista linha a linha: busca, período (mês ou ano todo), canal, status, avisos, filtros aplicados como chips, detalhe do lançamento e CSV | mês, semana, categoria |
| **Análises** | Dia a dia | `#/dia-a-dia` | Evolução diária, **Resumo do dia** (antigo Relatório diário), comparativo semanal e indicadores-chave do mês | mês, métrica, semana, categoria |
| | Mês a mês | `#/mensal` | Evolução do ano, acumulado e composição por tipo de ocorrência | métrica, categoria |
| | Canais e status | `#/canais` | Distribuição por canal (SAC, TD Nuvem, PIX Direto) e por status | mês, métrica, categoria |
| **Acompanhamento** | Metas e projeções | `#/metas` | Teto do mês, previsão de fechamento e projeções de estornos × realizado | mês |
| rodapé | Planilha | `#/planilha` | Origem dos dados, Atualizar agora, Importar outro arquivo e diagnóstico da leitura | — |

**Endereços antigos continuam funcionando** e passam para os novos: `#/visao-geral` → `#/inicio`, `#/diario-semanal` → `#/dia-a-dia`, `#/relatorio` → `#/dia-a-dia` (rolando até o Resumo do dia), `#/diagnostico` → `#/planilha`.

- **O endereço guarda a página e os filtros** (ex.: `#/lancamentos?mes=set&cat=Estorno&q=123`). **Copiar link** (no topo) copia a visão exata para mandar a alguém. No painel online, o link aponta para a URL `/exec`, e voltar/avançar do navegador funciona.
- **Busca global: Ctrl+K** (⌘K no Mac, ou `/`, ou a caixa de busca no topo) procura no ano inteiro por número de pedido, texto da observação, valor exato (`189,90`) ou data (`23/09`), e também leva a qualquer página.
- **Drill-down:** clicar num dia do Top 5, numa barra da evolução diária ou num dia da tabela abre o Resumo do dia; **Ver lançamentos do dia** abre a lista já filtrada. O mês na Composição e as linhas de canal e status (e as fatias da rosca) abrem Lançamentos filtrados. Clicar numa semana do comparativo seleciona a semana.
- **Alertas do Início:** meta ultrapassada ou prevista acima do teto (mesma regra de Metas), dias atípicos pelo mesmo critério do indicador (valor acima de média + 2σ), avisos da planilha, o que mudou desde a última visita (guardado só neste navegador) e até que data vão os dados.
- A barra de filtros mostra **só os filtros que mudam a página aberta**. Os outros ficam guardados e voltam a valer quando você abre uma página que os usa. Em Lançamentos, **Ano todo** dispensa mês e semana.
- O subtítulo abaixo do nome da página diz sempre qual recorte está na tela (mês, semana, dia, categoria).
- Alguns quadros têm recorte fixo, avisado no canto do quadro: **Últimos 7 dias** (até o último lançamento), **Composição por tipo** (ano inteiro, todas as categorias) e **Metas** (em R$, todas as categorias).
- **Exportar ▾** (no topo de todas as páginas): Excel com 4 abas, CSV dos lançamentos filtrados e Imprimir / PDF da página aberta.
- No celular, o menu abre pelo botão ☰, os filtros ficam no botão **Filtrar** e os lançamentos viram cartões; o detalhe do lançamento abre em tela cheia.

### Toque e gestos

- **Resumo do dia**: deslize para a esquerda (dia seguinte) ou para a direita (dia anterior). O conteúdo acompanha o dedo e o novo dia entra pelo lado certo. Se a tabela do dia tiver rolagem lateral, ela rola primeiro; o gesto troca o dia quando a tabela chega à borda. No computador, as setas ← → do teclado fazem o mesmo quando o foco está no quadro.
- **Detalhe do lançamento**: arraste para a direita para fechar.
- **Menu lateral** (celular): arraste para a esquerda para fechar.
- **Resumo por mês** (celular): os 12 meses viram uma faixa deslizante, já centralizada no mês escolhido.
- Todos os controles têm área de toque de pelo menos 44 px em telas de toque; os efeitos que dependem de passar o mouse (elevar cartões, brilho que segue o cursor) só aparecem com mouse.

## Funcionalidades

Cada indicador tem um ícone **i** com o nome técnico e a fórmula. O glossário completo, com o que foi removido e por quê, está em `../relatorios/glossario-indicadores-2026-09-25.md`.

- **KPIs do mês**: quantidade de ocorrências (volume), valor total reembolsado (R$), ticket médio de reembolso e acumulado no ano (YTD). Os três primeiros trazem a **variação MoM**: mês encerrado contra o mês anterior inteiro; mês em andamento contra os mesmos dias do mês anterior (like-for-like). Minigráfico anual em cada um.
- **Resumo por mês**: cartões compactos e clicáveis para trocar o mês, com valor, ocorrências e ticket médio (TM) numa linha, e uma barra com o valor do mês em relação ao maior mês do ano. O mês em andamento leva o selo "Parcial · dd/mm"; o selo "▲ meta" aparece quando o mês passa do teto definido (o valor da meta fica na dica do selo).
- **Indicadores-chave do mês**: pico diário (valor e quantidade), média diária por dia corrido, desvio-padrão (σ) e coeficiente de variação (CV), dias atípicos (acima de média + 2σ), concentração nos 5 maiores dias, dias com lançamento, efeito volume × ticket da variação MoM e, com todas as categorias, o mix por tipo e o peso de chargeback (% do valor).
- **Top 5 dias do mês**: maiores exposições, com a participação de cada dia no mês.
- **Evolução diária**: barras por dia, média diária (dias corridos) e acumulado no mês (MTD). Opção **Comparar com mês anterior** (linha tracejada). A semana escolhida no filtro fica destacada.
- **Comparativo semanal**: semanas fixas (dias 1–7, 8–14, 15–21, 22–28 e 29–fim do mês), com participação de cada semana. A média semanal usa só semanas completas de 7 dias já encerradas; a semana 5 e a semana em curso aparecem marcadas.
- **Evolução mensal**: linha do ano com média mensal (só meses encerrados), média móvel de 3 meses (MM3) e acumulado no ano (YTD). A tabela traz ticket médio e variação MoM.
- **Composição por tipo de ocorrência, mês a mês**: quantidade, valor, ticket médio, MoM, valor de estornos, cupons e chargebacks, peso de chargeback e participação no ano, com linha de total do ano.
- **Canais e status**: rosca e tabela por canal (quantidade, valor, ticket médio, participação). O canal vem da abertura escrita na Observação (`SAC: 2 (R$ 301,70) | TD Nuvem: 1 (R$ 269,90)`): cada canal recebe a sua parte da linha. Status só aparece quando a coluna está preenchida.
- **Resumo do dia** (em Dia a dia; antes Relatório diário): setas para dia anterior/seguinte, totais do dia por tipo (com ticket médio e participação), participação no mês até a data, desvio vs. média diária em R$ e em σ (z-score) e variação DoD contra o último dia com lançamento.
- **Metas e projeções**:
  - teto mensal em R$ com barra de consumo da meta, previsão de fechamento por run-rate (realizado ÷ dias decorridos × dias do mês), folga vs. meta e alerta quando a meta é ou deve ser ultrapassada;
  - **projeções de estornos**: registre o valor projetado de um dia (até 3 casas decimais, ex.: `R$ 123,456`) e compare com o realizado (desvio em R$ e em %). O restante do painel usa centavos.
- **Lançamentos**: busca; período (mês do filtro ou ano todo); filtros de canal, status e "somente com avisos"; chips com os filtros aplicados; ordenação por coluna, paginação e total filtrado. Clicar numa linha abre o **detalhe**: aba e linha da planilha, avisos, divisão por canal e outros lançamentos do mesmo pedido.
- **Exportar Excel**: 4 abas (lançamentos filtrados; resumo mensal com quantidade e valor por tipo, ticket médio, MoM, peso de chargeback, participação no ano e teto; diário do mês; ocorrências do diagnóstico).
- **Baixar CSV**: dos lançamentos, no padrão brasileiro (`;` e vírgula decimal).
- **Imprimir / PDF**: sempre em tema claro, sem botões e com as tabelas abertas; gráficos com altura fixa (180 px, os menores 100 px) para não ocuparem a página inteira.
- **Temas** (botão **Tema** no rodapé do menu lateral e na tela de importação; a escolha fica salva):
  - **Automático**: segue o claro ou escuro do sistema e acompanha a troca;
  - **Claro** (lilás suave) e **Ameixa** (escuro violeta), os originais;
  - **Corporativo**: claro e neutro, com o azul institucional do Nimbus (Nuvemshop);
  - **Grafite**: escuro neutro, sem matiz, para leitura longa de números;
  - **Alto contraste**: preto, branco e amarelo, bordas nítidas, links sublinhados, foco de 3 px e sem efeitos decorativos.

  Todos passam no WCAG AA (texto 4,5:1, cores de gráfico 3:1), e as cores dos gráficos acompanham o tema. A troca usa uma revelação circular a partir do botão, quando o navegador suporta.

### Movimento e interação

- Durações e curvas únicas para o painel inteiro (`--dur-1` a `--dur-5`, `--ease-out`, `--ease-emph`, `--ease-spring`) e animações só com transform/opacity.
- Elevação em quatro níveis por tema: cartões em repouso (`--shadow-sm`), hover e barra fixa (`--shadow-md`), menus e avisos (`--shadow-lg`), diálogos (`--shadow-xl`).
- Contagem animada nos KPIs, cartões entrando em sequência ao trocar de página, dica dos gráficos deslizando entre os pontos, menus que crescem a partir do botão e resposta ao pressionar em todos os controles.
- Com **reduzir movimento** ativado no sistema, tudo isso é desligado (números e gráficos aparecem prontos).

A data de referência ("hoje" do painel) é a do **último lançamento até a data atual**: lançamentos com data futura não empurram a referência. Ela define o mês aberto por padrão, o mês em andamento, a previsão de fechamento, as médias por dia corrido e os "últimos 7 dias".

Não há taxa de chargeback de mercado (chargeback ratio = chargebacks ÷ transações de venda): a planilha não traz o total de vendas. O painel mostra o **peso de chargeback no valor reembolsado**, que é outra coisa.

## Formato esperado da planilha

| Item | Regra |
|---|---|
| Abas mensais | Uma aba por mês, reconhecida pelo nome (`Janeiro`…`Dezembro`, ou abreviado: `Jan`, `Fev`…). Maiúsculas e acentos são ignorados. |
| Colunas obrigatórias | `Data`, `Tipo de Ocorrência`, `Quantidade`, `Valor`. O cabeçalho pode estar em qualquer uma das primeiras 150 linhas. |
| Colunas opcionais | `Pedido/Referência`, `Status`, `Observação` |
| Tipos reconhecidos | `Estorno`, `Cupom`, `Chargeback` |
| Ano de referência | Célula `B2` da aba `Consolidado Anual`. Sem ela, usa a mediana dos anos das datas (com 5 datas ou mais) ou, em último caso, o ano atual. |
| Canal | Lido da `Observação`. Com abertura (`SAC: 2 (R$ 301,70) \| TD Nuvem: 1 (R$ 269,90)`), cada canal recebe a sua quantidade e o seu valor. Sem números, um único canal citado recebe a linha inteira. Canais reconhecidos: `SAC`, `TD Nuvem`, `PIX Direto`. |

A leitura termina após duas linhas totalmente vazias. Linhas de "Total" são ignoradas. `-`, `—` ou `n/a` em Status e Pedido/Referência contam como não preenchido.

### Validações do diagnóstico

Nada é descartado sem aviso. Cada ocorrência mostra aba, linha, campo e o valor encontrado.

- **Erro (linha excluída dos cálculos):** data inválida ou vazia; tipo, quantidade ou valor ausentes ou não numéricos.
- **Aviso (linha mantida):** tipo fora de Estorno/Cupom/Chargeback, quantidade ou valor negativos, data fora do ano de referência, data de outro mês que não o da aba, possível lançamento duplicado, abertura por canal na Observação que não bate com Quantidade/Valor.
- **Aviso (linha excluída dos cálculos):** linha de exemplo do modelo da planilha (Observação começando com "Exemplo"). Apague-a da planilha para o aviso sumir.
- **Bloqueio:** nenhuma aba mensal reconhecida, nenhuma tabela válida ou nenhuma linha válida.

## Privacidade e armazenamento

- Na importação manual, a planilha **não é enviada** para nenhum servidor: tudo acontece no navegador.
- No painel online, a planilha vai do Google Drive direto para o navegador, pelo Apps Script, sem passar por outros servidores. Só contas @talgui.com.br conseguem abrir.
- Algumas informações ficam salvas **só no navegador de cada pessoa**, sem compartilhamento entre computadores:
  - último arquivo importado ou baixado do Drive (IndexedDB `reembolsos_db`);
  - metas mensais (`localStorage`: `reembolsos_metas_v1`);
  - projeções (`localStorage`: `reembolsos_projecoes_v1`);
  - tema (`localStorage`: `reembolsos_tema_v1`; valores `auto`, `light`, `dark`, `nimbus`, `grafite`, `contraste`).
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

Para criar uma página nova: acrescente a entrada em `ROUTES` (com a lista `filtros`), o link no menu lateral dentro do grupo certo e a `<section class="page" data-page="…">`. Não renomeie os ids das rotas existentes: eles estão nos favoritos das pessoas. Se precisar trocar um id, deixe o antigo em `ROUTE_ALIASES`. Os parâmetros do endereço são montados em `buildHash` e lidos em `applyHashParams`.

```bash
git clone https://github.com/messiasalexandro-byte/painel-reembolsos.git
cd painel-reembolsos
# abra o index.html no navegador
```
