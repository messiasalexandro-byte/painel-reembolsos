# Painel de Reembolsos

Painel executivo para acompanhar **estornos, cupons e chargebacks** a partir da planilha de controle de reembolsos.

| Quero… | Vá para |
|---|---|
| ver o painel com os dados atuais | [Acessar o painel](#acessar-o-painel) |
| saber onde está cada informação | [Navegação](#navegação) |
| preparar ou corrigir a planilha | [Formato esperado da planilha](#formato-esperado-da-planilha) |
| ver as vendas da Nuvemshop | [Vendas (Nuvemshop)](#vendas-nuvemshop) |
| mudar o painel e publicar | [Para quem mantém](#para-quem-mantém) |

### Novidades (21/10/2026)

- **Devoluções reestruturadas:** cada solicitação vira um caso lido em quatro eixos (enquadramento, etapa, desfecho e motivo), com ciclo de vida de 7 etapas, filas de trabalho por prioridade, lista de casos com filtros por eixo, taxa de devolução por coorte sem trocas, tempo de ciclo e uma aba de regras e prazos. Veja [Devoluções (AfterClick)](#devoluções-afterclick).

### Novidades (01/10/2026)

- **Vendas da Nuvemshop:** novas páginas **Vendas** (abas Mês, Semana e Dia) e **Pedidos**, alimentadas pelas exportações de pedidos salvas na pasta `nuvemshop` do Drive. Mostram faturamento, pedidos, ticket médio, líquido recebido, formas de pagamento, cupons e o que ficou fora das vendas. Veja [Vendas (Nuvemshop)](#vendas-nuvemshop).
- **Cadastro com senha no painel online:** só entra quem foi liberado na página **Acesso**, e cada pessoa cria a própria senha no primeiro acesso. Veja [Quem acessa e senha](#quem-acessa-e-senha).
- **Semana já marcada:** ao abrir o painel, o filtro de semana vem na semana do último lançamento.

### Novidades (25/09/2026)

- **Marca Talgui:** o ícone do menu lateral agora é o logotipo Talgui (letras brancas em fundo preto, igual em todos os temas) e também aparece como ícone da aba do navegador.
- **Início reorganizado:** os cartões de mês subiram para o topo e escolhem o mês; a categoria fica no cabeçalho deles e a ordenação (Valor ou Qtde) no Top 5. A barra de filtros solta saiu do Início.
- **Dia a dia e metas:** Metas e projeções virou uma aba da página Dia a dia (`#/dia-a-dia?aba=metas`); o endereço antigo `#/metas` abre direto essa aba.
- **Indicadores revisados:** o canal passou a ser dividido pela abertura da Observação, as linhas de exemplo ficam fora das contas, e as médias e comparações usam só dias já decorridos. Os indicadores têm nomes técnicos (MTD, YTD, MoM, run-rate, σ) e um ícone "i" com a fórmula.
- **Nova estrutura:** Início com alertas, Lançamentos, Análises e Planilha. Também chegam busca global (Ctrl+K), clique no gráfico para abrir os lançamentos filtrados, link da visão com os filtros e Exportar ▾.
- **Correções da revisão:** voltar do navegador no painel online sem loop (uma entrada por página), KPIs do mês em andamento só até a data de referência, alerta que separa linhas de exemplo (fora das contas) dos demais avisos, canal por palavra inteira e `R$ 1.500` lido como mil e quinhentos, "Ano todo" em Lançamentos limpa o dia, balão acessível no ícone "i" e quantidade nos cartões do celular.
- **Visual e interação:** seis temas (Automático, Claro, Ameixa, Corporativo, Grafite e Alto contraste), gestos no celular, animações suaves e impressão compacta.

---

## Acessar o painel

### 1. Painel online (recomendado)

**https://script.google.com/macros/s/AKfycbxgwqnN5kGArZmYM-zbGyV_h4AjfkzqITj2ZTfVP1PfpZMDMu2Ag8FDnlDpEGiWbXVy/exec**

Abre já com os dados da planilha que está na pasta do Google Drive. Não é preciso importar nada.

- **Quem acessa:** só contas **@talgui.com.br** liberadas na página Acesso, e com senha (veja abaixo).
- **Atualização:** com o painel aberto, ele verifica a pasta **a cada 5 minutos** e só redesenha se a planilha mudou, mantendo o mês, os filtros e a página abertos. **Atualizar agora**, no menu lateral, força a verificação.
- **Qual arquivo:** a Planilha Google editada por último na pasta (exportada como `.xlsx`). Se não houver nenhuma Planilha Google, usa o `.xlsx` modificado por último. Subpastas são ignoradas, então dá para guardar versões antigas numa subpasta "Arquivo".
- **Se a versão nova tiver erro de estrutura**, o painel mantém os dados anteriores e avisa no menu lateral.
- **Várias contas Google no mesmo navegador:** se o painel não carregar, abra numa janela anônima só com a conta da Talgui. É uma limitação conhecida do Google Apps Script.

#### Quem acessa e senha

- **Primeiro acesso:** depois que o administrador libera o e-mail, a pessoa abre o link e cria uma senha de pelo menos 8 caracteres.
- **Nas outras vezes:** a senha fica lembrada no navegador até a pessoa clicar em **Sair**, no menu lateral.
- **Senha errada:** depois de 5 tentativas erradas, a conta fica bloqueada por 15 minutos.
- **Esqueceu a senha:** o administrador clica em **Liberar novo cadastro** na página Acesso. A senha antiga deixa de valer, e a pessoa cria outra no próximo acesso.
- **Administradores:** o dono do script e quem estiver na propriedade `PAINEL_ADMINS`. Só eles veem a página **Acesso**, onde adicionam e removem e-mails. Quem é removido perde o acesso na hora. Se estiver com o painel aberto, ele trava na próxima verificação.
- **Contas fora da lista:** veem só o aviso "Você não tem acesso a este painel".

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
| — | Início | `#/inicio` | Situação do mês em segundos: alertas (meta, dias atípicos, avisos da planilha, o que mudou), cartões de mês (escolhem o mês) com a categoria no cabeçalho, KPIs, meta do mês, últimos 7 dias e top 5 dias (com "Ordenar por Valor / Qtde") | mês, métrica, categoria (no próprio conteúdo) |
| — | Lançamentos | `#/lancamentos` | Lista linha a linha: busca, período (mês ou ano todo), canal, status, avisos, filtros aplicados como chips, detalhe do lançamento e CSV | mês, semana, categoria |
| **Análises** | Dia a dia e metas | `#/dia-a-dia` | Duas abas. **Dia a dia:** evolução diária, **Resumo do dia** (antigo Relatório diário), comparativo semanal e indicadores-chave do mês. **Metas e projeções** (`#/dia-a-dia?aba=metas`): teto do mês, previsão de fechamento e projeções de estornos × realizado | aba Dia a dia: mês, métrica, semana, categoria · aba Metas: mês |
| | Mês a mês | `#/mensal` | Evolução do ano, acumulado e composição por tipo de ocorrência | métrica, categoria |
| | Canais e status | `#/canais` | Distribuição por canal (SAC, TD Nuvem, PIX Direto) e por status | mês, métrica, categoria |
| rodapé | Minha conta | `#/minha-conta` | Só no painel online, para quem entrou com senha: foto (aparece no menu; sem foto, as iniciais do e-mail), e-mail e perfil (Administrador ou Usuário), trocar senha e **Sair** (o botão saiu do rodapé do menu e fica aqui). No GitHub Pages o item não aparece e o endereço abre o Início | — |
| rodapé | Acesso | `#/acesso` | Só administradores, no painel online: quem pode abrir o painel, adicionar e remover e-mails e liberar novo cadastro | — |
| rodapé | Planilha | `#/planilha` | Origem dos dados, Atualizar agora, Importar outro arquivo e diagnóstico da leitura | — |

**Endereços antigos continuam funcionando** e passam para os novos: `#/visao-geral` → `#/inicio`, `#/diario-semanal` → `#/dia-a-dia`, `#/relatorio` → `#/dia-a-dia` (rolando até o Resumo do dia), `#/metas` → `#/dia-a-dia?aba=metas`, `#/diagnostico` → `#/planilha`.

- **Menu lateral retrátil:** no computador, a alça « na borda direita do menu recolhe o menu para uma coluna só de ícones (passe o mouse para ver o nome da página) e o » o expande de novo. A escolha fica guardada neste navegador. No celular o menu continua abrindo como gaveta pelo botão ☰, e na impressão nada muda.
- **Minha conta:** a foto é recortada no centro e reduzida no próprio navegador (64×64, JPEG, no máximo 4000 caracteres) antes de ir para o servidor. Cada cadastro com foto ocupa cerca de 4 KB, e as propriedades do script somam no máximo 500 KB: cabem por volta de 100 pessoas com foto (sem foto, cada cadastro ocupa menos de 200 bytes). Trocar a senha desconecta os outros aparelhos e navegadores, que passam a pedir a senha nova.

- **Abas de Dia a dia e metas:** clique ou setas ← → (Home/End) quando o foco está nas abas. A aba vai para o endereço (`aba=metas`; sem o parâmetro, abre Dia a dia). Trocar de aba, como trocar filtro, não cria entrada no histórico. Abrir o Resumo de um dia (Top 5, alerta, Ctrl+K) sempre leva à aba Dia a dia.

- **O endereço guarda a página e os filtros** (ex.: `#/lancamentos?mes=set&cat=Estorno&q=123`). **Copiar link** (no topo) copia a visão exata para mandar a alguém. No painel online, o link aponta para a URL `/exec`, e voltar/avançar do navegador funciona: cada troca de página cria **uma** entrada no histórico (a da janela de cima, pelo `google.script.history`); filtros só atualizam a entrada atual. Início → Lançamentos → Voltar volta ao Início, e mais um Voltar sai do painel.
- **Busca global: Ctrl+K** (⌘K no Mac, ou `/`, ou a caixa de busca no topo) procura no ano inteiro por número de pedido, texto da observação, valor exato (`189,90`) ou data (`23/09`), e também leva a qualquer página.
- **Drill-down:** clicar num dia do Top 5, numa barra da evolução diária ou num dia da tabela abre o Resumo do dia; **Ver lançamentos do dia** abre a lista já filtrada. O mês na Composição e as linhas de canal e status (e as fatias da rosca) abrem Lançamentos filtrados. Clicar numa semana do comparativo seleciona a semana.
- **Alertas do Início:** meta ultrapassada ou prevista acima do teto (mesma regra da aba Metas; **Ver metas e projeções** abre a aba), dias atípicos pelo mesmo critério do indicador (valor acima de média + 2σ), avisos da planilha (contados à parte: "12 linhas de exemplo fora dos cálculos" e os avisos que não tiram linhas das contas), linhas do mês em andamento com dia posterior à data de referência (ficam fora dos KPIs), o que mudou desde a última visita (guardado só neste navegador) e até que data vão os dados.
- **No Início não há barra de filtros:** o mês se escolhe nos cartões de mês (no topo), a categoria no cabeçalho desses cartões e a métrica em **Ordenar por** no Top 5 (é o único bloco do Início que ela muda). São os mesmos filtros das outras páginas: o que você escolhe no Início continua valendo nelas, e vice-versa.
- Nas outras páginas, a barra de filtros mostra **só os filtros que mudam a página aberta**. Os outros ficam guardados e voltam a valer quando você abre uma página que os usa. Em Lançamentos, **Ano todo** dispensa mês e semana.
- O subtítulo abaixo do nome da página diz sempre qual recorte está na tela (mês, semana, dia, categoria).
- Alguns quadros têm recorte fixo, avisado no canto do quadro: **Últimos 7 dias** (até o último lançamento), **Composição por tipo** (ano inteiro, todas as categorias) e a aba **Metas** (em R$, todas as categorias).
- **Exportar ▾** (no topo de todas as páginas): Excel com 4 abas, CSV dos lançamentos filtrados e Imprimir / PDF da página aberta.
- No celular, o menu abre pelo botão ☰, os filtros ficam no botão **Filtrar** (no Início, a categoria fica logo acima da faixa de meses) e os lançamentos viram cartões (com a quantidade junto ao tipo, ex.: "Estorno · 37 ocorr."); o detalhe do lançamento abre em tela cheia.

### Toque e gestos

- **Resumo do dia**: deslize para a esquerda (dia seguinte) ou para a direita (dia anterior). O conteúdo acompanha o dedo e o novo dia entra pelo lado certo. Se a tabela do dia tiver rolagem lateral, ela rola primeiro; o gesto troca o dia quando a tabela chega à borda. No computador, as setas ← → do teclado fazem o mesmo quando o foco está no quadro.
- **Detalhe do lançamento**: arraste para a direita para fechar.
- **Menu lateral** (celular): arraste para a esquerda para fechar.
- **Resumo por mês** (celular): os 12 meses viram uma faixa deslizante, já centralizada no mês escolhido.
- Todos os controles têm área de toque de pelo menos 44 px em telas de toque; os efeitos que dependem de passar o mouse (elevar cartões, brilho que segue o cursor) só aparecem com mouse.

## Funcionalidades

Cada indicador tem um ícone **i** com o nome técnico e a fórmula, num balão que abre ao passar o mouse, ao focar pelo teclado (Tab) ou ao tocar (toque de novo, toque fora ou Esc fecham). O texto também vai para o leitor de tela e o balão não sai na impressão. O glossário completo, com o que foi removido e por quê, está em `../relatorios/glossario-indicadores-2026-09-25.md`.

- **KPIs do mês**: quantidade de ocorrências (volume), valor total reembolsado (R$), ticket médio de reembolso e acumulado no ano (YTD). Os três primeiros trazem a **variação MoM**: mês encerrado contra o mês anterior inteiro; mês em andamento contra os mesmos dias do mês anterior (like-for-like). No mês em andamento, quantidade, valor, ticket e mix contam só os dias 1 até a data de referência, como a previsão; linhas da aba com dia posterior aparecem numa nota no Início. Minigráfico anual em cada um.
- **Resumo por mês** (no topo do Início): cartões compactos e clicáveis para trocar o mês, com valor, ocorrências e ticket médio (TM) numa linha, e uma barra com o valor do mês em relação ao maior mês do ano. O mês em andamento leva o selo "Parcial · dd/mm"; o selo "▲ meta" aparece quando o mês passa do teto definido (o valor da meta fica na dica do selo).
- **Indicadores-chave do mês**: pico diário (valor e quantidade), média diária por dia corrido, desvio-padrão (σ) e coeficiente de variação (CV), dias atípicos (acima de média + 2σ), concentração nos 5 maiores dias, dias com lançamento, efeito volume × ticket da variação MoM e, com todas as categorias, o mix por tipo e o peso de chargeback (% do valor).
- **Top 5 dias do mês**: maiores exposições, com a participação de cada dia no mês.
- **Evolução diária**: barras por dia, média diária (dias corridos) e acumulado no mês (MTD). Opção **Comparar com mês anterior** (linha tracejada). A semana escolhida no filtro fica destacada.
- **Comparativo semanal**: semanas fixas (dias 1–7, 8–14, 15–21, 22–28 e 29–fim do mês), com participação de cada semana. A média semanal usa só semanas completas de 7 dias já encerradas; a semana 5 e a semana em curso aparecem marcadas.
- **Evolução mensal**: linha do ano com média mensal (só meses encerrados), média móvel de 3 meses (MM3) e acumulado no ano (YTD). A tabela traz ticket médio e variação MoM.
- **Composição por tipo de ocorrência, mês a mês**: quantidade, valor, ticket médio, MoM, valor de estornos, cupons e chargebacks, peso de chargeback e participação no ano, com linha de total do ano.
- **Canais e status**: rosca e tabela por canal (quantidade, valor, ticket médio, participação). O canal vem da abertura escrita na Observação (`SAC: 2 (R$ 301,70) | TD Nuvem: 1 (R$ 269,90)`): cada canal recebe a sua parte da linha. Status só aparece quando a coluna está preenchida.
- **Resumo do dia** (em Dia a dia; antes Relatório diário): setas para dia anterior/seguinte, totais do dia por tipo (com ticket médio e participação), participação no mês até a data, desvio vs. média diária em R$ e em σ (z-score) e variação DoD contra o último dia com lançamento.
- **Metas e projeções** (aba de Dia a dia e metas):
  - teto mensal em R$ com barra de consumo da meta, previsão de fechamento por run-rate (realizado ÷ dias decorridos × dias do mês), folga vs. meta e alerta quando a meta é ou deve ser ultrapassada;
  - **projeções de estornos**: registre o valor projetado de um dia (até 3 casas decimais, ex.: `R$ 123,456`) e compare com o realizado (desvio em R$ e em %). O restante do painel usa centavos.
- **Lançamentos**: busca; período (mês do filtro ou ano todo; escolher ano todo limpa o dia vindo de um drill-down); filtros de canal, status e "somente com avisos"; chips com os filtros aplicados; ordenação por coluna, paginação e total filtrado. Clicar numa linha abre o **detalhe**: aba e linha da planilha, avisos, divisão por canal e outros lançamentos do mesmo pedido.
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

## Vendas (Nuvemshop)

### Como alimentar

1. No admin da Nuvemshop, vá em **Vendas → Exportar** e baixe o CSV do período (de preferência com **todos os status, inclusive cancelados**).
2. Salve o arquivo na pasta **`nuvemshop`** do Google Drive, sem converter para Planilha Google. Pode ficar na raiz ou numa subpasta, por exemplo `nuvemshop/historico` para as exportações antigas.
3. O painel online lê **todos** os CSVs da pasta e das subpastas e junta os pedidos pelo **número do pedido**. Se o mesmo pedido aparece em duas exportações, vale a do arquivo modificado por último. Assim, as exportações podem se sobrepor em datas, e o status mais novo prevalece.

**Histórico (desde fev/2026):** exporte em pedaços de 2 ou 3 meses, por exemplo fev–abr e mai–jul. Cada CSV é lido **uma vez só** e guardado numa base compacta por ano, em `nuvemshop/_base/vendas-AAAA.json`, que o painel cria e mantém sozinho. Não mexa nessa pasta.
- Na primeira abertura depois de um arquivo novo, a página mostra "Organizando as exportações novas…" até terminar. Isso pode levar alguns minutos com o histórico.
- Se um CSV for **apagado** da pasta, a base é refeita sem ele.
- Na página Vendas, **Origem das vendas** mostra o período de cada exportação e o total de pedidos sem repetição.
- Com mais de um ano na base, aparece o seletor de **Ano** acima dos meses.

Fora do painel online (GitHub Pages), use **Importar CSV da Nuvemshop** na página Vendas. As vendas ficam só naquele navegador.

### Regras das métricas

- **Data base:** sempre a **data do pedido**. A data de pagamento aparece só no detalhe do pedido.
- **Situação do pedido**, nesta ordem:
  - **Cancelado:** status do pedido ou do pagamento cancelado, ou com data de cancelamento preenchida.
  - **Pagamento recusado:** status do pagamento "Recusado".
  - **Cortesia:** "Pedido com 100% de desconto" ou total zero.
  - **Venda:** pagamento confirmado, inclusive os estornados depois.
  - **Aguardando pagamento:** os demais.
- Só a **venda** entra no faturamento. As outras situações aparecem em **Fora das vendas**.
- **Faturamento** = soma do Total dos pedidos de venda, ou seja, subtotal − desconto + frete.
- **Líquido recebido** = Total líquido da Nuvemshop, isto é, faturamento − taxas − juros do parcelamento − estornos + devolução dos custos.
- **Formas de pagamento:**
  - Pix;
  - Cartão à vista (1 parcela);
  - Cartão parcelado (2 ou mais parcelas);
  - Boleto;
  - Outros (ex.: "A combinar").
- **Comparações:**
  - O **mês** em andamento é comparado com os mesmos dias do mês anterior (MoM).
  - A **semana** (de segunda a domingo) é comparada com a anterior (WoW), nos mesmos dias se ela estiver em curso.
  - O **dia** é comparado com o mesmo dia da semana anterior (D-7). No dia do último pedido, a comparação vai até o mesmo horário.
  - Sem base completa, ou seja, quando a base começa antes da primeira exportação, não há comparação.

### Privacidade

O CSV da Nuvemshop traz nome, e-mail, CPF, telefone e endereço. Esses dados **não saem do servidor**: o `obterVendas` devolve só número, datas, status, valores, pagamento, cupom, estado, canal e envio.

## Devoluções (AfterClick)

### Como alimentar

1. Na AfterClick (Trocas e devoluções), exporte as solicitações com **todos os status** dos **últimos 60 dias**, de preferência todo dia.
2. Salve o arquivo (.xlsx ou .csv) na subpasta **`nuvemshop/AJD`** do Google Drive.
3. O painel lê cada exportação **uma vez** e junta as solicitações pelo **Nº do Pedido** da AfterClick (ex.: D-TLG-53088-1). Vale o arquivo modificado por último. A base fica em `nuvemshop/_base/devolucoes.json` e o controle dos arquivos em `devolucoes-estado.json`. Não mexa nesses arquivos.

A primeira carga foi o export completo desde mar/2026 (`Pedidos.xlsx`).

### Como a página está organizada (desde 21/10/2026)

Cada solicitação da AfterClick é um **caso** (Nº da solicitação, ex.: D-TLG-53088-1). A página vai do resumo ao detalhe: **Filas de trabalho** (resumo, ciclo de vida e filas) → **Casos** (lista segmentada) → **detalhe do caso** (drawer). O endereço guarda a aba e os filtros (ex.: `#/devolucoes?aba=casos&et=recebida&al=fora_sla`).

Todo caso é lido em quatro eixos independentes, com os mesmos nomes e cores em todas as telas:

| Eixo | Valores | De onde vem |
|---|---|---|
| **Enquadramento** (sugerido) | Arrependimento · Vício do produto · Erro da loja · Não entregue · Troca comercial · Fora de prazo · A classificar | Motivo + data de entrega do pedido de origem, nesta ordem: vício → erro da loja → não entregue → até 7 dias da entrega → política |
| **Etapa** | 1 Solicitada · 2 Autorizada · 3 Em trânsito · 4 Recebida · 5 Conferida · 6 Encerrada · 7 Cancelada | Status AfterClick + rastreio da reversa |
| **Desfecho** | Forma (estorno no cartão, Pix, boleto/TED, crédito em loja, troca) × situação (não devido, devido, solicitado, concluído na loja, divergente) | Solução pretendida + meio de pagamento do pedido + card FINA |
| **Motivo** | Tamanho · Caimento · Aparência · Arrependimento · Defeito · Erro da loja · Logística · Sinais de uso (achado de conferência) · Outro | Motivo da AfterClick (um por solicitação) |

**Alertas** (marcam o caso, não mudam a etapa): fora do SLA, prazo vencendo, divergente, chargeback, sem movimento e desatualizado.

### Telas

- **Filas de trabalho:**
  - quatro indicadores: casos em andamento (com quantos estão fora do SLA ou sem movimento), peça na loja sem desfecho, **a devolver em dinheiro** e **crédito em loja a emitir** (a troca aparece à parte, sem valor a devolver);
  - **ciclo de vida:** um cartão por etapa com quantidade, valor, mediana de dias na etapa e quantos passaram do SLA. Clicar abre a lista de casos da etapa. A tabela por etapa fica em "Ver etapas em tabela";
  - **filas**, em ordem de prioridade: prazo legal ou da plataforma vencendo, conferida sem desfecho, recebida sem conferência, NF pendente ou fora do prazo, aguardando triagem, postagem expirada ou sem movimento, desfecho divergente, chargeback ligado a devolução, dados a revisar e devido acima do saldo Nuvem Pago (esta sem dado). Cada fila diz o critério e a ordem, e todas têm as mesmas colunas;
  - rastreio da reversa (autorizados e em trânsito) e notas fiscais de devolução.
- **Casos:** lista com busca (caso, pedido, SKU ou produto) e um filtro por eixo (etapa, enquadramento, desfecho previsto, motivo, alerta, período de abertura). Cada opção mostra quantos casos sobram. Chips com os filtros, paginação e **Baixar CSV**.
- **Ritmo e taxas:** casos abertos (com cancelados depois e válidos), peças, recebidas e encerradas por dia, semana ou mês; **taxa de devolução por mês da venda** (pedidos, efetivada, troca, peças e valor, sobre pedidos de venda ou entregues); **tempo de ciclo** (mediana, P90 e % no SLA de cada trecho); **enquadramento × desfecho previsto**.
- **Peças e motivos:** produtos mais devolvidos com taxa de devolução, contados pela data da venda (coorte, padrão) ou da solicitação, com ou sem trocas; tamanhos; motivos com grupo e enquadramento pelo motivo, para conferir o mapeamento.
- **Regras e prazos:** ciclo de vida, cadeia de SLA, prazos da lei e da plataforma, ordem do enquadramento, desfecho, vocabulário e dados que ainda não chegam ao painel. É gerada das mesmas constantes que as telas usam.
- **Detalhe do caso:** pílulas dos quatro eixos e dos alertas, resumo, linha do tempo das etapas, SLA, enquadramento e prazos (arrependimento, vício, Nuvem Pago), desfecho (valor a devolver, valor devolvido quando a Nuvemshop ou a planilha mostram, divergência), conferência e peça, NF e Financeiro, peças, lançamentos da planilha de reembolsos e dados técnicos (status AfterClick, rastreio e visto por último).
- No **drawer do pedido** (Pedidos) aparecem os casos daquele pedido com as mesmas pílulas. Os indicadores de **Vendas** têm a linha "Pedidos com devolução" com a mesma regra da taxa (sem cancelados; trocas à parte).

### Regras

- **Etapas:**
  - `replacement` = Solicitada;
  - `replacement-authorized` = Autorizada (peça ainda não postada), **Em trânsito** (rastreio de postagem, coleta ou transporte) ou **Recebida** (rastreio "Entregue");
  - `replacement-approved` = Conferida, com resultado Aprovada (reprovada e parcial não vêm no export);
  - `replacement-completed` = Encerrada;
  - `canceled` = Cancelada.
  - Status ou evento de rastreio que o painel não reconhece não vira etapa nem "aguardando": vai para a fila **Dados a revisar**.
- **Cadeia única de SLA** (`DV_SLA` no `index.html`, dias corridos): triagem 2, postagem pela cliente 30, transporte sem novo evento 10, conferência 2, desfecho 3, NF de devolução 7, Financeiro após a NF 7, ciclo interno total 14. Passar do SLA da loja = **fora do SLA**; passar do da cliente ou da transportadora = **sem movimento** (não é mais "parada" nem sai das contas).
- **Válidos** = casos sem os cancelados. É a base de peças, valores, taxa, aba Peças e da linha de Vendas. O ritmo mostra também os cancelados depois, à parte.
- **Taxa de devolução:** pela data do pedido de origem (coorte), sem trocas (a troca tem taxa própria). Denominador: pedidos de venda ou entregues, de todos os anos da base de vendas.
- **Data de entrega:** a API da Nuvemshop não informa. O servidor guarda em `_base/entregas.json` a primeira vez em que viu o pedido como entregue; o painel trata a entrega como um intervalo (data do pedido ou último registro sem entrega → primeira vez visto entregue). Com o intervalo incerto, o enquadramento fica "A classificar".
- **Visto por último:** dia do export mais novo que trouxe o caso. Caso em andamento que ficou para trás saiu da janela de 60 dias: status congelado, alerta "desatualizado".
- **Datas de status:** o export não traz. O painel registra quando viu cada caso mudar de status, de uma exportação para a seguinte. Autorização, conferência, encerramento e cancelamento só têm data a partir do início do controle (05/10/2026); antes disso, os dias na etapa aparecem com "≈".
- **Taxa por produto:** usa as peças vendidas (SKU e Quantidade Comprada dos CSVs da Nuvemshop) nos meses do período escolhido. Taxa acima de 100% aparece como "—", porque as vendas do recorte não cobrem as devoluções.

### NF de devolução (desde 21/10/2026)

- **De onde vem:** da planilha mais nova "… - Exportação do Painel - Financeiro" (cards FINA), na pasta do Drive `1Cy5a2oVeecFEo6_Z8na99gZUv5rP6v2w`. A propriedade `NF_PASTA` troca essa pasta.
- **Como o painel acha a NF:** o número e a data saem do texto "NF emitida: Nº … - Emitida em …" (ou "NF emitida em dd/mm"), na Descrição ou nas Observações. Se não houver texto, usa a coluna "Data NF emitida".
- **Como liga à solicitação:** o card entra na solicitação pelo **Pedido Nuvem**, que pode ser o nº D-TLG da AfterClick ou o pedido da Nuvemshop. No segundo caso, só liga se o pedido tiver uma única solicitação.
- **Cache:** o resultado fica em `_base/nf-financeiro.json` e só é refeito quando a planilha muda.
- **Dois prazos, os dois de 7 dias corridos:**
  - **recebimento → NF**. O recebimento é a data da entrega no rastreio ou, sem ela, o dia em que o painel viu a solicitação aprovada.
  - **NF → estorno/cupom**. Conta enquanto o card não está Concluído, ou até a "Data de processamento", quando ela está preenchida.
- **Onde aparece:** na fila **NF pendente ou fora do prazo** e no painel **Notas fiscais de devolução** da aba Filas de trabalho, e no detalhe do caso. Os dois prazos vêm da cadeia única de SLA.
- **Privacidade:** do card, só entram código, fase, pedido, ação, número e data da NF, criação e processamento.

### Privacidade

O export da AfterClick traz nome, CPF e e-mail do cliente e a nota fiscal. Esses dados **não entram na base nem saem do servidor**: o `obterDevolucoes` devolve só número, datas, status, motivo, solução, valores e as peças (SKU, preço e quantidade). Dos pedidos de origem, devolve só data, status do pagamento e do envio, meio de pagamento, parcelas, estorno, total e as datas de entrega registradas pelo painel.

## Formato esperado da planilha

| Item | Regra |
|---|---|
| Abas mensais | Uma aba por mês, reconhecida pelo nome (`Janeiro`…`Dezembro`, ou abreviado: `Jan`, `Fev`…). Maiúsculas e acentos são ignorados. |
| Colunas obrigatórias | `Data`, `Tipo de Ocorrência`, `Quantidade`, `Valor`. O cabeçalho pode estar em qualquer uma das primeiras 150 linhas. |
| Colunas opcionais | `Pedido/Referência`, `Status`, `Observação` |
| Tipos reconhecidos | `Estorno`, `Cupom`, `Chargeback` |
| Ano de referência | Célula `B2` da aba `Consolidado Anual`. Sem ela, usa a mediana dos anos das datas (com 5 datas ou mais) ou, em último caso, o ano atual. |
| Canal | Lido da `Observação`. Com abertura (`SAC: 2 (R$ 301,70) \| TD Nuvem: 1 (R$ 269,90)`), cada canal recebe a sua quantidade e o seu valor. Sem números, um único canal citado recebe a linha inteira. Canais reconhecidos: `SAC`, `TD Nuvem`, `PIX Direto`, só como palavra ou expressão inteira, sem diferenciar acentos e maiúsculas ("Transação", "Isaac" e "sacola" não contam como SAC). Valores com ponto de milhar e sem centavos (`R$ 1.500`) valem mil e quinhentos. |

A leitura termina após duas linhas totalmente vazias. Linhas de "Total" são ignoradas. `-`, `—` ou `n/a` em Status e Pedido/Referência contam como não preenchido.

### Validações do diagnóstico

Nada é descartado sem aviso. Cada ocorrência mostra aba, linha, campo e o valor encontrado.

- **Erro (linha excluída dos cálculos):** data inválida ou vazia; tipo, quantidade ou valor ausentes ou não numéricos.
- **Aviso (linha mantida):** tipo fora de Estorno/Cupom/Chargeback, quantidade ou valor negativos, data fora do ano de referência, data de outro mês que não o da aba, possível lançamento duplicado, abertura por canal na Observação que não bate com Quantidade/Valor.
- **Aviso (linha excluída dos cálculos):** linha de exemplo do modelo da planilha (Observação começando com "Exemplo"). Apague-a da planilha para o aviso sumir.
- **Bloqueio:** nenhuma aba mensal reconhecida, nenhuma tabela válida ou nenhuma linha válida.

## Privacidade e armazenamento

- Na importação manual, a planilha **não é enviada** para nenhum servidor: tudo acontece no navegador.
- No painel online, a planilha vai do Google Drive direto para o navegador, pelo Apps Script, sem passar por outros servidores. Só contas @talgui.com.br liberadas e com a senha certa conseguem abrir. A senha é conferida no servidor e guardada só como hash, então ninguém consegue lê-la, nem o administrador.
- Algumas informações ficam salvas **só no navegador de cada pessoa**, sem compartilhamento entre computadores:
  - último arquivo importado ou baixado do Drive (IndexedDB `reembolsos_db`), apagado ao clicar em **Sair**;
  - acesso ao painel online (`localStorage`: `reembolsos_acesso_v1`), apagado ao clicar em **Sair**;
  - metas mensais (`localStorage`: `reembolsos_metas_v1`);
  - projeções (`localStorage`: `reembolsos_projecoes_v1`);
  - tema (`localStorage`: `reembolsos_tema_v1`; valores `auto`, `light`, `dark`, `nimbus`, `grafite`, `contraste`);
  - vendas importadas à mão fora do painel online (IndexedDB `vendas_db`), apagadas em **Esquecer as vendas deste navegador**.
- No painel online, as vendas não ficam guardadas no navegador: o ano aberto é buscado a cada abertura. No Drive, a base `nuvemshop/_base` guarda só as colunas sem dados pessoais.
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
    ├── Codigo.gs         ← doGet() entrega a página; cadastro e senha; obterPlanilha(versao, token) devolve a planilha do Drive;
    │                       obterVendas(token, versao, ano) consolida os CSVs da pasta nuvemshop na base por ano
    ├── appsscript.json   ← escopos (Drive com escrita, para a base de vendas) e acesso restrito ao domínio (DOMAIN)
    ├── .clasp.json       ← scriptId usado pelo clasp
    ├── testar.sh         ← envia o index.html para o /dev (sem mexer no /exec)
    └── publicar.sh       ← publica o index.html no Apps Script, mesma URL
```

- A pasta das vendas é achada pelo nome `nuvemshop`. O id fica guardado na propriedade `NUVEMSHOP_PASTA`, e também pode ser posto lá à mão, para usar outra pasta.
- `VENDAS_BASE` (propriedade): quais CSVs já entraram na base (id → data de modificação), o período de cada um e o id do arquivo de cada ano em `_base`.
  - Apagar essa propriedade faz a base ser refeita do zero na próxima abertura.
  - Os arquivos antigos de `_base` podem ser jogados fora.

- `apps-script/Index.html` **não é mantido à mão**: o `publicar.sh` o gera copiando o `index.html` a cada publicação (e o `.gitignore` o exclui do Git). Não edite esse arquivo.
- O mesmo `index.html` serve os dois lugares: ele detecta se está dentro do Apps Script (`google.script.run`) e, nesse caso, busca a planilha no Drive e verifica mudanças a cada 5 minutos.
- A pasta do Drive é definida em `FOLDER_ID`, no `Codigo.gs`. Ela continua privada: o script lê o arquivo em nome da conta dona.
- O acesso fica nas **Propriedades do script**, no editor do Apps Script, em Configurações do projeto. Elas valem para o `/dev` e o `/exec` ao mesmo tempo:
  - `usuario:<e-mail>`: um cadastro por pessoa, com o hash e o sal da senha, ou só a data, enquanto aguarda o primeiro acesso;
  - `PAINEL_SEGREDO`: chave que assina os tokens de acesso, criada sozinha. Se for apagada, todo mundo precisa digitar a senha de novo;
  - `PAINEL_ADMINS` (opcional): administradores além do dono, separados por vírgula;
  - `PAINEL_MIGRADO`: marca que a lista antiga (`PAINEL_EMAILS`, da época da senha única) já virou cadastros. `PAINEL_EMAILS` e `PAINEL_SENHA` não são mais usadas e podem ser apagadas.
- Se o dono esquecer a própria senha, apague a propriedade `usuario:<e-mail do dono>`. No próximo acesso ele cria outra senha.

### Conferir no Apps Script antes de publicar

O histórico (voltar/avançar) e o foco inicial do teclado dependem do iframe do Apps Script e não dá para testar no GitHub Pages. Antes de publicar, abra a **implantação de teste** (Editor do Apps Script → Implantar → Testar implantações → URL terminada em `/dev`) e confira:

1. Início → Lançamentos (pelo menu) → **Voltar** volta ao Início; **Voltar** de novo sai do painel, sem alternar entre Início e Lançamentos.
2. Trocar o mês não cria entrada no histórico; **Avançar** refaz a navegação.
3. Um link com filtros (`/dev#/lancamentos?mes=ago`) abre a página e o filtro certos.
4. Logo ao abrir, **Ctrl+K** abre a busca sem precisar clicar no painel (se não abrir, é o navegador bloqueando o foco; um clique resolve).

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

Para criar uma página nova: acrescente a entrada em `ROUTES` (com a lista `filtros`), o link no menu lateral dentro do grupo certo e a `<section class="page" data-page="…">`. Não renomeie os ids das rotas existentes: eles estão nos favoritos das pessoas. Se precisar trocar um id, deixe o antigo em `ROUTE_ALIASES`. Os parâmetros do endereço são montados em `buildHash` e lidos em `applyHashParams`. Uma rota pode ter `abas` (cada aba com os seus `filtros`, como em Dia a dia e metas) e `filtrosNaPagina: true` quando os controles ficam no próprio conteúdo e a barra de filtros deve sumir (Início).

```bash
git clone https://github.com/messiasalexandro-byte/painel-reembolsos.git
cd painel-reembolsos
# abra o index.html no navegador
```
