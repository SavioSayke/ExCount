# Programa de Contagem de Elementos em Colunas de Excel

Este é um programa em Python criado para ler um arquivo Excel e contabilizar quantas vezes um determinado elemento aparece em uma coluna específica. O programa permite ao usuário selecionar as colunas e os elementos a serem verificados e gera um arquivo de resultados.

---

## Como Usar

Este programa requer que você tenha o [Python](https://www.python.org/downloads/) instalado em sua máquina, bem como a biblioteca `pandas` para manipulação do arquivo Excel. Você pode instalar o `pandas` executando o seguinte comando:

```bash
pip install pandas
```

### Passo 1: Preparando o Arquivo Excel
1. Copie seu arquivo Excel (com a extensão `.xlsx`) para a mesma pasta onde o programa está localizado.
2. Renomeie o arquivo Excel para **`table.xlsx`**.

### Passo 2: Executando o Programa
1. Execute o script Python.
2. O programa irá apresentar um menu com as seguintes opções:
    - **1**: Iniciar a contagem de elementos.
    - **2**: Exibir o tutorial de como usar o programa.
    - **3**: Exibir os créditos de criação do programa.
    - **4**: Sair do programa.

### Passo 3: Contando Elementos
1. Selecione a opção `1` para iniciar o processo.
2. O programa irá solicitar que você insira:
    - O número de colunas que deseja verificar.
    - O nome exato das colunas e os elementos que deseja contar.
3. O programa irá gerar o resultado da contagem e salvar no arquivo `Resultado.txt`.

### Exemplo de Resultado:
O resultado será salvo no arquivo `Resultado.txt` com a seguinte formatação:

## Funcionalidades

- **Contagem de Elementos**: O programa permite que você escolha um ou mais nomes de colunas em um arquivo Excel e conte quantas vezes um elemento específico aparece nessas colunas.
- **Tutorial e Instruções**: O programa fornece um tutorial para guiar os usuários na utilização.
- **Créditos**: Exibe quem criou o programa.
- **Arquivo de Resultados**: Os resultados são salvos em um arquivo de texto chamado `Resultado.txt` dentro da mesma pasta.

## Funções do Programa

### Funções de Carregamento
O programa apresenta barras de progresso para indicar o andamento de diferentes etapas, como:
- **Carregando**: Durante a inicialização do programa.
- **Buscando Elementos**: Enquanto o programa busca o elemento nas colunas selecionadas.
- **Obtendo Resultados**: Quando o programa gera o arquivo de resultados.

### Funções de Menu
- **Tutorial**: Explica como preparar o arquivo e iniciar o processo.
- **Créditos**: Exibe os créditos do desenvolvedor do programa.
- **Sair**: Encerra o programa.

## Estrutura do Código

- **Imports**: Utiliza as bibliotecas `pandas`, `os` e `time` para manipulação de arquivos Excel, controle de arquivos e exibição de carregamento, respectivamente.
- **Arquivo Excel**: O programa espera um arquivo Excel com o nome exato `table.xlsx` no mesmo diretório.
- **Resultado**: A contagem de elementos é salva em um arquivo de texto chamado `Resultado.txt`.