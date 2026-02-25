# corretor-ortografico-simples-python
Este código corrige palavras escritas erradas em português.

Para rodar, é necessário baixar a biblioteca language_tool_python:
!pip install language_tool_python

## Fluxo do código:
1. Verifica se o arquivo existe, se sim, continua para o próximo passo, senão, mostra um erro de 'Arquivo não encontrado'.

2. Leitura do arquivo

3. Verifica se a coluna que será analisada existe. Se não existir, mostra um erro e imprime na tela as colunas disponíveis no arquivo lido.

4. Inicializa o Language tool

5. Para verificar se está processando o código, foi usado a biblioteca tqdm. 

6. Verifica erros e corrige o texto

7. Fecha o Language Tool e salva os resultados em um arquivo de excel.

8. Arquivo analise_ortografica_gramatical.xlsx criado

9. Imprime os 5 erros mais frequentes no Terminal

