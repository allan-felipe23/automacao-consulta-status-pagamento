# Programa de Automação de Consulta de Status de Pagamento

Este projeto é um programa de automação que consulta o status de pagamento de clientes com base em seus CPFs. Ele lê os dados dos clientes de uma planilha Excel (`dados_clientes.xlsx`), acessa um site de consulta de CPF, verifica o status de pagamento de cada cliente e, dependendo do status, atualiza uma planilha de fechamento (`planilha fechamento.xlsx`) com as informações relevantes. Se o pagamento estiver em dia, o programa captura a data e o método de pagamento. Caso contrário, marca o status como "pendente".

## Funcionalidades

- **Leitura de Dados**: O programa lê os dados dos clientes de uma planilha Excel.
- **Consulta de CPF**: Acessa um site de consulta de CPF para verificar o status de pagamento.
- **Atualização de Planilha**: Atualiza uma planilha de fechamento com o status de pagamento de cada cliente.
- **Automação Completa**: O processo é totalmente automatizado, desde a leitura dos dados até a atualização da planilha.

## Pré-requisitos

Antes de executar o programa, certifique-se de que você tem os seguintes requisitos instalados:

- **Python 3.x**: O programa foi desenvolvido em Python. Certifique-se de ter o Python instalado em sua máquina.
- **Bibliotecas Python**: As bibliotecas `openpyxl` e `selenium` são necessárias para executar o programa. Você pode instalá-las usando o pip:

  ```bash
  pip install openpyxl selenium
  ```

- **WebDriver do Chrome**: O programa utiliza o WebDriver do Chrome para automatizar o navegador. Certifique-se de ter o ChromeDriver instalado e configurado. Você pode baixar o ChromeDriver [aqui](https://sites.google.com/chromium.org/driver/).

## Como Usar

1. **Prepare as Planilhas**:
   - Crie uma planilha chamada `dados_clientes.xlsx` com os dados dos clientes. A planilha deve ter as seguintes colunas:
     - `Nome`
     - `Valor`
     - `CPF`
     - `Vencimento`
   - Crie uma planilha chamada `planilha fechamento.xlsx` para armazenar os resultados da consulta. Esta planilha será atualizada automaticamente pelo programa.

2. **Execute o Programa**:
   - Coloque o script Python no mesmo diretório das planilhas.
   - Execute o script Python:

     ```bash
     python nome_do_script.py
     ```

3. **Resultados**:
   - O programa irá atualizar a planilha `planilha fechamento.xlsx` com o status de pagamento de cada cliente. Se o pagamento estiver em dia, a data e o método de pagamento serão incluídos. Caso contrário, o status será marcado como "pendente".

## Estrutura do Código

O código está dividido em várias etapas, cada uma com uma função específica:

1. **Carregar a Planilha de Clientes**: O programa abre a planilha `dados_clientes.xlsx` e acessa a aba `Sheet1`.
2. **Inicializar o Navegador**: O navegador Chrome é aberto e direcionado para o site de consulta de CPF.
3. **Iterar sobre os Clientes**: O programa lê cada linha da planilha, começando da segunda linha (para ignorar o cabeçalho).
4. **Preencher o Campo de Pesquisa**: O CPF do cliente é inserido no campo de pesquisa do site.
5. **Clicar no Botão de Pesquisar**: O botão de pesquisa é clicado para consultar o status do pagamento.
6. **Verificar o Status**: O programa verifica se o pagamento está "em dia" ou "atrasado".
7. **Capturar Dados Adicionais (se estiver em dia)**: Se o pagamento estiver em dia, a data e o método de pagamento são extraídos.
8. **Atualizar a Planilha de Fechamento**: Dependendo do status, o programa atualiza a planilha `planilha fechamento.xlsx` com as informações relevantes.
9. **Intervalo entre Consultas**: Um intervalo de 2 segundos é adicionado entre cada consulta para evitar sobrecarregar o servidor.
10. **Fechar o Navegador**: Após processar todos os clientes, o navegador é fechado.

## Considerações Finais

- **Tempos de Espera**: Os tempos de espera (`sleep`) podem precisar de ajustes dependendo da velocidade da sua conexão com a internet e do tempo de resposta do site.
- **Erros e Exceções**: O programa inclui um bloco `try-finally` para garantir que o navegador seja fechado corretamente, mesmo que ocorra um erro durante a execução.
- **Uso em Produção**: Este programa foi desenvolvido para fins educacionais e de portfólio. Certifique-se de testá-lo adequadamente antes de usá-lo em um ambiente de produção.

## Contribuições

Contribuições são bem-vindas! Se você encontrar algum problema ou tiver sugestões de melhorias, sinta-se à vontade para abrir uma issue ou enviar um pull request.

## Licença

Este projeto está licenciado sob a licença MIT. Consulte o arquivo [LICENSE](LICENSE) para mais detalhes.
