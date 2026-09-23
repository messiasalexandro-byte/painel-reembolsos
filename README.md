# Painel de Reembolsos

Painel executivo para acompanhar **estornos, cupons e chargebacks** a partir da planilha de controle de reembolsos (`.xlsx`).
É uma página única, sem servidor: a planilha é lida no navegador e nenhum dado sai do computador de quem usa.

**Acesse:** https://messiasalexandro-byte.github.io/painel-reembolsos/

---

## Como usar

1. Abra o link acima (ou o arquivo `index.html` no navegador).
2. Arraste a planilha `.xlsx` para a área de importação ou clique para selecioná-la.
3. Revise o **diagnóstico da importação**: ele lista abas encontradas, linhas lidas, válidas e excluídas, além de erros e avisos.
4. Clique em **Continuar para o Dashboard**.

Na próxima visita, o painel oferece **"Continuar com <arquivo>"** para reabrir o último arquivo sem novo upload. **"Esquecer arquivo"** apaga essa cópia do navegador.

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

## Navegação

O painel funciona como um site com **menu lateral**. Cada página tem endereço próprio, então dá para salvar nos favoritos ou mandar o link para uma página específica:

| Página | Endereço | Conteúdo |
|---|---|---|
| Visão geral | `#/visao-geral` | KPIs, resumo por mês, indicadores-chave e Top 5 |
| Diário e semanal | `#/diario-semanal` | Comportamento diário, comparativo semanal e últimos 7 dias |
| Mensal e composição | `#/mensal` | Trajetória mensal e composição por categoria |
| Relatório diário | `#/relatorio` | Totais de um dia e projeções de estornos |
| Canais e status | `#/canais` | Distribuição por canal e por status |
| Metas e projeções | `#/metas` | Teto mensal, previsão de fechamento e projetado × realizado |
| Lançamentos | `#/lancamentos` | Tabela linha a linha |
| Diagnóstico | `#/diagnostico` | Resultado da validação da planilha |

No celular, o menu fica escondido e abre pelo botão ☰. Os filtros do topo valem para todas as páginas, menos o Diagnóstico. **Imprimir / PDF** imprime a página aberta.

## Funcionalidades

**Filtros globais:** mês, métrica (quantidade ou valor), semana e categoria.

- **KPIs**: volume, impacto financeiro, ticket médio e representatividade no ano, com variação sobre o mês anterior e minigráfico anual.
- **Resumo por mês**: cartões clicáveis para trocar o mês. Mostram "▲ acima da meta" quando o mês passa do teto definido.
- **Comportamento diário**: barras por dia, média e acumulado. Opção **Comparar com mês anterior** (linha tracejada).
- **Trajetória mensal** e **acumulado no ano**.
- **Comparativo semanal**: semanas fixas (dias 1–7, 8–14, 15–21, 22–28 e 29–31).
- **Composição por categoria**: tabela do ano inteiro.
- **Top 5** dias de maior exposição e **indicadores-chave** do período.
- **Tendência dos últimos 7 dias**.
- **Relatório diário**: totais do dia por natureza, comparação com a média e com o dia anterior.
- **Canais e status**: gráfico de rosca por canal e tabela por status.
- **Metas e projeções**:
  - teto mensal em R$ com barra de progresso;
  - previsão de fechamento pelo ritmo diário;
  - alerta quando a meta é ou deve ser ultrapassada;
  - tabela de projeções de estornos (projetado × realizado).
- **Lançamentos**: tabela linha a linha com:
  - busca;
  - filtros de canal, status e "somente com avisos";
  - ordenação por coluna, paginação e total filtrado.
- **Exportar Excel**: 4 abas (lançamentos filtrados, resumo mensal, diário do mês, ocorrências do diagnóstico).
- **Baixar CSV**: dos lançamentos, no padrão brasileiro (`;` e vírgula decimal).
- **Imprimir / PDF**: sempre em tema claro, sem botões e com as tabelas abertas.
- **Tema claro/escuro**: a escolha fica salva.

A previsão de fechamento usa a data do último lançamento da planilha como "hoje".

## Privacidade e armazenamento

- A planilha **não é enviada** para nenhum servidor. Todo o processamento acontece no navegador.
- Algumas informações ficam salvas **só no navegador de cada pessoa**, sem compartilhamento entre computadores:
  - último arquivo importado (IndexedDB `reembolsos_db`);
  - metas mensais (`localStorage`: `reembolsos_metas_v1`);
  - projeções (`localStorage`: `reembolsos_projecoes_v1`);
  - tema (`localStorage`: `reembolsos_tema_v1`).
- Limpar os dados do site no navegador apaga essas informações.

## Tecnologia

- Um único arquivo: `index.html`, com HTML, CSS e JavaScript.
- [SheetJS](https://sheetjs.com/) 0.18.5 para ler e gerar `.xlsx`.
- [Chart.js](https://www.chartjs.org/) 4.5.1 para os gráficos.
- As duas bibliotecas são carregadas via CDN (cdnjs). Por isso, **é preciso internet** para abrir o painel.

O JavaScript está dividido em três blocos dentro do `index.html`:

| Bloco | Responsabilidade |
|---|---|
| `RefundsApp.core` | Leitura da planilha, validação e diagnóstico (`processWorkbook`), formatação |
| `RefundsApp.agg` | Agregações: por mês, dia, semana, categoria, canal/status, previsão, filtros |
| Renderização | Telas, gráficos, painéis, exportação, tema e eventos |

## Desenvolvimento

Não há etapa de build nem dependências para instalar:

```bash
git clone https://github.com/messiasalexandro-byte/painel-reembolsos.git
cd painel-reembolsos
# abra o index.html no navegador
```

A publicação é feita pelo GitHub Pages a partir da branch `main`. Cada push atualiza o site em cerca de 1 minuto.
