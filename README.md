<h1 align="center">Planilha de Funcionários e Dependentes - Formatter</h1>

<p align="center">
  <img src="https://github.com/user-attachments/assets/e9335630-b228-43a8-b47f-289b8c9f1e53" alt="projeto_excel" />
</p>

![Python](https://img.shields.io/badge/python-3.10%2B-blue)
![License](https://img.shields.io/badge/license-MIT-green)
![Status](https://img.shields.io/badge/status-active-brightgreen)

## 📑 Descrição do Projeto

Este projeto surgiu de uma necessidade emergencial na empresa onde trabalho. Fomos surpreendidos com a demanda de transformar uma planilha complexa de funcionários com dependentes em um novo formato, exigido por sistemas externos — e o prazo era curtíssimo.

Sem experiência prévia com a biblioteca Pandas, mergulhei de cabeça no desafio e desenvolvi esse script para automatizar a tarefa que antes estava sendo feita manualmente. O projeto é simples, mas funcional, e representa meu primeiro contato prático com manipulação de dados usando Python.

A entrada é uma planilha onde cada funcionário pode ter até 6 dependentes em uma única linha. A saída é um novo arquivo Excel formatado, onde cada pessoa (funcionário ou dependente) é representada em uma linha separada e pronta para importação em outros sistemas.

## 🛠️ Funcionalidades Principais

- **Transformação de dados tabulares**
  - Cada funcionário e dependente é convertido para uma linha única.
  - Conversão de datas para formato brasileiro (dd/mm/aaaa).
  - Organização de colunas com repetições intencionais para manter integridade do vínculo.

- **Formatação de planilha**
  - Estilização automática do cabeçalho: fundo azul com texto branco em negrito.
  - Alinhamento centralizado de todas as células.
  - Exportação para novo arquivo Excel.

- **Configuração com segurança**
  - Caminho da planilha é lido de um arquivo `.env` para segurança e flexibilidade.
  
## 🧰 Tecnologias Utilizadas

- **Linguagem:** Python 3.10+
- **Bibliotecas:**
  - `pandas` - manipulação de dados tabulares
  - `openpyxl` - leitura e escrita de arquivos Excel com formatação
  - `dotenv` - leitura de variáveis de ambiente
  - `os` - manipulação de caminhos de arquivos
  - `datetime` - tratamento de datas
- **Gestão de Configuração:** Dotenv (.env)

## 📁 Estrutura Esperada do Arquivo de Entrada

| Nome Funcionario | Cpf Funcionario | Id Plano Escolhido | Nome Dep 1 | Cpf Dep 1 | Sexo Dep 1 | Data Nascimento Dep 1 | Nome Mae Dep 1 | Parentesco Dep 1 | ... | Data Cadastro Update |
|------------------|------------------|---------------------|-------------|-----------|-------------|------------------------|----------------|-------------------|-----|-----------------------|

- A planilha pode conter até 6 dependentes por funcionário, com colunas nomeadas sequencialmente (`Dep 1`, `Dep 2`, ..., `Dep 6`).

## 🧪 Exemplo de Uso

Para executar o projeto:

1. Instale as dependências:

```bash
pip install pandas openpyxl python-dotenv
