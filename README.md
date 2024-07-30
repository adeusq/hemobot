# Sistema de Automatização de Planilhas

Este projeto é um sistema de automatização de tarefas para planilhas do Excel, desenvolvido em Python usando a biblioteca Tkinter para a interface gráfica. O sistema permite preencher planilhas e exportar dados para arquivos de texto (.txt).

## Funcionalidades

- **Preencher Planilha**: Permite iniciar o preenchimento de uma planilha do Excel a partir de uma linha específica.
- **Exportar Dados para TXT**: Exporta dados de uma planilha do Excel para um arquivo de texto a partir de uma linha específica.
- **Menu Principal**: Interface principal do sistema que permite acessar as funcionalidades de preenchimento e exportação.
- **Tela de Sobre**: Mostra informações sobre o desenvolvedor do sistema.
- **Login e Autorização**: Tela de login com autenticação básica e persistência de login por um período específico (ou indefinidamente, conforme configuração).

## Tecnologias Utilizadas

- **Python 3.x**: Linguagem de programação utilizada para desenvolver o sistema.
- **Tkinter**: Biblioteca para a criação da interface gráfica do usuário.
- **openpyxl**: Biblioteca para manipulação de arquivos Excel (se necessário).
- **os**: Biblioteca para operações com o sistema operacional, como a verificação de arquivos.

## Requisitos

Antes de executar o projeto, certifique-se de que você possui o Python 3.x instalado e que as bibliotecas necessárias estão disponíveis. Você pode instalar as bibliotecas necessárias usando o `pip`:

```bash
pip install openpyxl
