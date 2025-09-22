# <img width="40" height="30" alt="app_icon" src="https://github.com/user-attachments/assets/ebc7d54c-6c43-4c21-927c-9bd1aad5affa" /> NEXUS - Automação de Planilhas (Rateio & Atualização)

Aplicação desktop desenvolvida em Python com PySide6 para automatizar processos de planilhas que eram originalmente executados por scripts VBA. O objetivo é fornecer uma solução mais rápida, robusta e de fácil manutenção.

## Funcionalidades Principais

-   **Gerar Rateio:** Consolida dados das abas "Controle de Radios", "Controle de Componentes" e "Controle Solicitações" para gerar um relatório de rateio final, salvo em um novo arquivo Excel.
-   **Atualização Massiva:** Atualiza as abas "Controle de Radios" e "Controle de Componentes" com informações da aba "Base de Dados", como Centro de Custo e Gestor.
-   **Interface Gráfica:** Painel de controle amigável para selecionar o arquivo de origem e executar as automações, com feedback de progresso em tempo real.
-   **Segurança:** Gera novos arquivos de resultado com data e hora no nome, prevenindo que dados sejam sobrescritos acidentalmente.
-   **Validação:** Verifica a estrutura da planilha de entrada antes de iniciar o processamento para evitar erros inesperados.

## Screenshot

<img width="1538" height="696" alt="image" src="https://github.com/user-attachments/assets/66f1a9c5-7006-4e9b-91d5-5980816430b1" /> 

**● Rateio**

<img width="1555" height="711" alt="image" src="https://github.com/user-attachments/assets/9c5b5749-0919-4371-b297-ed7d871ce28b" />
 <img width="874" height="635" alt="image" src="https://github.com/user-attachments/assets/f68e4c1b-a7af-4fd0-a126-7a005607c131" /> 
 
 **● Atualização Massiva** 
 
 <img width="1452" height="717" alt="image" src="https://github.com/user-attachments/assets/f85d1a63-cb9d-4ff1-921a-60637ff3f684" />
<img width="1555" height="722" alt="image" src="https://github.com/user-attachments/assets/dbf56597-3b70-45d4-9031-5667daeb3a1c" />



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
