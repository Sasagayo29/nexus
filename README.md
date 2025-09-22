# NEXUS - Automação de Planilhas (Rateio & Atualização)

Aplicação desktop desenvolvida em Python com PySide6 para automatizar processos de planilhas que eram originalmente executados por scripts VBA. O objetivo é fornecer uma solução mais rápida, robusta e de fácil manutenção.

## Funcionalidades Principais

-   **Gerar Rateio:** Consolida dados das abas "Controle de Radios", "Controle de Componentes" e "Controle Solicitações" para gerar um relatório de rateio final, salvo em um novo arquivo Excel.
-   **Atualização Massiva:** Atualiza as abas "Controle de Radios" e "Controle de Componentes" com informações da aba "Base de Dados", como Centro de Custo e Gestor.
-   **Interface Gráfica:** Painel de controle amigável para selecionar o arquivo de origem e executar as automações, com feedback de progresso em tempo real.
-   **Segurança:** Gera novos arquivos de resultado com data e hora no nome, prevenindo que dados sejam sobrescritos acidentalmente.
-   **Validação:** Verifica a estrutura da planilha de entrada antes de iniciar o processamento para evitar erros inesperados.

## Screenshot

*(Aqui você pode adicionar um print da tela principal da aplicação para referência visual.)*
![Screenshot da Aplicação](assets/screenshot.png)

## Pré-requisitos

-   Python 3.8 ou superior

## Instalação

Siga os passos abaixo para configurar o ambiente e executar a aplicação.

1.  **Clone ou baixe** este projeto para sua máquina local.

2.  **Navegue** até a pasta raiz do projeto (`automacao_rateio/`) pelo seu terminal.

3.  **(Recomendado)** Crie e ative um ambiente virtual para isolar as dependências do projeto:
    ```sh
    # Criar o ambiente virtual
    python -m venv venv

    # Ativar no Windows
    .\venv\Scripts\activate

    # Ativar no macOS/Linux
    source venv/bin/activate
    ```

4.  **Instale as dependências** necessárias a partir do arquivo `requirements.txt`:
    ```sh
    pip install -r requirements.txt
    ```

## Como Usar

1.  Certifique-se de que sua planilha de controle (ex: `Motorola - Planilha de Controle.xlsm`) esteja disponível.

2.  Execute a aplicação a partir da pasta raiz do projeto (`automacao_rateio/`):
    ```sh
    python main.py
    ```

3.  Na janela da aplicação, clique em **"Selecionar Arquivo"** e escolha sua planilha.

4.  Clique no botão da ação desejada: **"Gerar Rateio"** ou **"Executar Atualização Massiva"**.

5.  Acompanhe o progresso na caixa de log na parte inferior da janela.

6.  Ao final, um pop-up informará o sucesso da operação. Os arquivos de resultado (ex: `Rateio_Final_2025-09-22_10-23-00.xlsx`) serão salvos na mesma pasta onde a aplicação foi executada.

## Estrutura do Projeto
Com certeza. Um bom projeto precisa de uma boa documentação. O arquivo README.md é a porta de entrada para qualquer pessoa (incluindo você no futuro) que precise entender e utilizar a aplicação.

Aqui está uma proposta de README.md completo, seguindo as melhores práticas.

Como Usar
Crie um arquivo de texto na pasta principal do projeto (automacao_rateio/).

Nomeie o arquivo como README.md (a extensão .md é importante).

Copie e cole o conteúdo abaixo dentro dele.

Conteúdo para o README.md
Markdown

# Automação de Planilhas (Projeto Rateio & Atualização)

Aplicação desktop desenvolvida em Python com PySide6 para automatizar processos de planilhas que eram originalmente executados por scripts VBA. O objetivo é fornecer uma solução mais rápida, robusta e de fácil manutenção.

## Funcionalidades Principais

-   **Gerar Rateio:** Consolida dados das abas "Controle de Radios", "Controle de Componentes" e "Controle Solicitações" para gerar um relatório de rateio final, salvo em um novo arquivo Excel.
-   **Atualização Massiva:** Atualiza as abas "Controle de Radios" e "Controle de Componentes" com informações da aba "Base de Dados", como Centro de Custo e Gestor.
-   **Interface Gráfica:** Painel de controle amigável para selecionar o arquivo de origem e executar as automações, com feedback de progresso em tempo real.
-   **Segurança:** Gera novos arquivos de resultado com data e hora no nome, prevenindo que dados sejam sobrescritos acidentalmente.
-   **Validação:** Verifica a estrutura da planilha de entrada antes de iniciar o processamento para evitar erros inesperados.

## Screenshot

*(Aqui você pode adicionar um print da tela principal da aplicação para referência visual.)*
![Screenshot da Aplicação](assets/screenshot.png)

## Pré-requisitos

-   Python 3.8 ou superior

## Instalação

Siga os passos abaixo para configurar o ambiente e executar a aplicação.

1.  **Clone ou baixe** este projeto para sua máquina local.

2.  **Navegue** até a pasta raiz do projeto (`automacao_rateio/`) pelo seu terminal.

3.  **(Recomendado)** Crie e ative um ambiente virtual para isolar as dependências do projeto:
    ```sh
    # Criar o ambiente virtual
    python -m venv venv

    # Ativar no Windows
    .\venv\Scripts\activate

    # Ativar no macOS/Linux
    source venv/bin/activate
    ```

4.  **Instale as dependências** necessárias a partir do arquivo `requirements.txt`:
    ```sh
    pip install -r requirements.txt
    ```

## Como Usar

1.  Certifique-se de que sua planilha de controle (ex: `Motorola - Planilha de Controle.xlsm`) esteja disponível.

2.  Execute a aplicação a partir da pasta raiz do projeto (`automacao_rateio/`):
    ```sh
    python main.py
    ```

3.  Na janela da aplicação, clique em **"Selecionar Arquivo"** e escolha sua planilha.

4.  Clique no botão da ação desejada: **"Gerar Rateio"** ou **"Executar Atualização Massiva"**.

5.  Acompanhe o progresso na caixa de log na parte inferior da janela.

6.  Ao final, um pop-up informará o sucesso da operação. Os arquivos de resultado (ex: `Rateio_Final_2025-09-22_10-23-00.xlsx`) serão salvos na mesma pasta onde a aplicação foi executada.

## Estrutura do Projeto
````
automacao_rateio/
├── 📁 core/                            # Lógica de negócio (Pandas)
│   ├── processamento_rateio.py
│   └── atualizacao_massiva.py
│
├── 📁 ui/                              # Interface gráfica (PySide6)
│   ├── main_window.py
│   └── worker.py
│
├── 📁 assets/
│   └── app_icon.png                    # Logo da aplicação
│
├── main.py                             # Ponto de entrada da aplicação
└── requirements.txt                    # Lista de dependências
````
## Melhorias Futuras

-   [ ] **Empacotamento em .exe:** Criar um executável standalone com PyInstaller para distribuição em máquinas sem Python instalado.
-   [ ] **Arquivo de Configuração:** Mover nomes de abas e colunas para um arquivo `config.ini` para facilitar a adaptação a mudanças na planilha.
-   [ ] **Testes Automatizados:** Implementar testes unitários para garantir a estabilidade do código a longo prazo.
