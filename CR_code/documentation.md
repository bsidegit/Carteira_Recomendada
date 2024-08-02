# Documentação do Projeto de Simulação de Carteira

## Visão Geral

Este projeto consiste no desenvolvimento de scripts em Python para simulação de carteiras de investimentos, utilizando dados de fundos, benchmarks e ativos específicos. A simulação foi feita para integrar-se com planilhas Excel, dando uma análise detalhada de retornos, volatilidade, drawdown, correlações e outras métricas de performance.

## Estrutura dos Códigos

### 1. `simulacao_carteira.py`

O código `simulacao_carteira.py` foi o ponto de partida do projeto, projetado por Eduardo Scheffer em 12 de Setembro de 2022. Ele foi desenvolvido inicialmente para realizar todas as operações de simulação em uma única função `main_code()`. Durante o desenvolvimento, identificamos a necessidade de modularizar o código para melhorar a manutenção e a compreensão. Algumas correções importantes realizadas nesta versão incluem ajustes em manipulações de datas, como a conversão de `pd.to_datetime(benchmark_Returns.index).date > date_24M` para `benchmark_Returns.index > pd.Timestamp(date_24M)`, entre outras correções bastante similares.

### 2. `simulacao_carteira_(new).py`

Esta versão é uma refatoração modularizada do `simulacao_carteira.py`. O objetivo principal foi dividir a lógica da simulação em funções menores e específicas, facilitando a visualização e o entendimento do código. As funções foram separadas para leitura de dados, cálculo de retornos, manipulação de colunas categóricas, entre outras tarefas. Esta modularização também possibilitou a reutilização de código e a adição de novas funcionalidades de forma mais estruturada.

### 3. `simulacao_carteira_(bill).py`

Esta versão foi desenvolvida para integrar o código com a planilha 'Carteira Recomendada - Bill'. Baseando-nos na estrutura modular do `simulacao_carteira_(new).py`, ajustamos o código para utilizar dados específicos da nova planilha e garantir que os cálculos e resultados fossem consistentes com os fornecidos anteriormente pela 'Carteira Recomendada - MFO'. É importante destacar que, para o funcionamento do código, o usuário preencha a coluna 'BENCH' com o Benchmark desejado para todos os ativos. 

## Implementações Futuras

1. **Integração com Macro no Excel**:
   - Adicionar um botão na 'Carteira Recomendada - Bill' que, ao ser clicado, execute um macro que rode o executável do `simulacao_carteira_(bill).py`. Isso permitirá que os cálculos sejam realizados diretamente no Excel, facilitando a utilização pelos usuários finais.

2. **Campos de Input para Data e Taxa de Gestão**:
   - Adicionar campos na 'Carteira Recomendada - Bill' para que os usuários possam inserir a data de início e fim da simulação, além da taxa de gestão. Isso permitirá uma personalização maior das simulações.

3. **Seção 8: Exportação de Resultados para Excel**:
   - Implementar uma seção no `simulacao_carteira_(bill).py` para que os resultados gerados pelos cálculos sejam exportados diretamente para o Excel. Isso incluirá métricas de performance, retornos mensais e outros dados relevantes, como nos códigos anteriores a ele. 

4. **Geração de Gráficos no Excel**:
   - Utilizar os resultados exportados para gerar gráficos diretamente no Excel, baseando-se nas fórmulas e estilos de gráficos usados na 'Carteira Recomendada - MFO'. Isso facilitará a visualização dos dados e a análise de performance.

5. **Atualização do Banco de Dados**:
   - Realizar atualizações no banco de dados para garantir a integridade dos dados, evitando dados faltantes ou inconsistências, como cotas de fundos que foram trocadas em períodos que não estivessem mais na janela de reciclagem da atualização do Banco de Dados. Isso é necessário para garantir que os cálculos de simulação sejam precisos.

## Conclusão

Este projeto proporcionou uma visão detalhada de como integrar cálculos de simulação de carteiras com ferramentas de análise de dados como o Excel, além de garantir a precisão e a modularidade do código. As futuras implementações visam melhorar a experiência do usuário e a precisão dos dados, tornando a ferramenta mais robusta, fácil de usar e confiável. O projeto teve seu início no dia 01/07/2024 e foi finalizado no dia 02/08/2024. 

