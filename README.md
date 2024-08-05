# YouTube Data Automation

## Objetivo dos Scripts

Estes dois scripts Python foram desenvolvidos para gerenciar e atualizar estatísticas de vídeos do YouTube em dois formatos distintos: um banco de dados SQL Server e um arquivo Excel. Ambos utilizam a API do YouTube para coletar dados e realizam operações de leitura e escrita em arquivos locais.

### Script 1: `ts_web_db_youtube.py`

Este script é responsável por extrair estatísticas de vídeos do YouTube a partir de uma lista contida em um arquivo Excel e inserir esses dados em um banco de dados SQL Server.

#### Funcionalidades:

1. **Carregar Dados do Excel**: Lê um arquivo Excel contendo informações sobre vídeos do YouTube.
2. **Conectar ao Banco de Dados**: Estabelece uma conexão com um banco de dados SQL Server.
3. **Obter Estatísticas do YouTube**: Realiza chamadas à API do YouTube para obter estatísticas de visualizações, likes e comentários.
4. **Inserir Dados no Banco de Dados**: Insere as estatísticas obtidas no banco de dados.
5. **Log e Notificações**: Registra logs de operações e exibe mensagens de notificação em caso de erro.

#### Requisitos:

- **Bibliotecas Python**: `pandas`, `requests`, `pyodbc`, `tkinter`
- **Configuração**: Substitua a chave da API do YouTube e ajuste as configurações de conexão com o banco de dados conforme necessário.

### Script 2: `ts_web_excel_youtube.py`

Este script lê informações de vídeos do YouTube a partir de um arquivo Excel, coleta estatísticas via API do YouTube, e atualiza um segundo arquivo Excel com essas estatísticas.

#### Funcionalidades:

1. **Carregar Dados do Excel**: Lê um arquivo Excel com dados dos vídeos e outro arquivo para armazenar as estatísticas.
2. **Obter Estatísticas do YouTube**: Faz chamadas à API do YouTube para obter visualizações, likes e comentários.
3. **Atualizar Excel com Estatísticas**: Adiciona as novas estatísticas ao arquivo Excel existente ou cria um novo se necessário.
4. **Log e Notificações**: Registra logs de operações e exibe mensagens de notificação em caso de erro.

#### Requisitos:

- **Bibliotecas Python**: `pandas`, `requests`, `tkinter`
- **Configuração**: Substitua a chave da API do YouTube conforme necessário.

## Uso

### Executar `yt_ts.py` ou `Excel-yt_ts.py`

1. Clone o Repositório:
   ```bash
   git clone https://github.com/ClaudioLimahub/python_YouTube_TS.git

2. Instale as Dependências:
   pip install pandas requests pyodbc

3. Configure o Ambiente:
Atualize a chave da API do YouTube no script.
Ajuste a configuração de conexão com o banco de dados SQL Server.

4. Execute o Script:
   python yt_ts.py ou python Excel-yt_ts.py

## Contribuições
Sinta-se à vontade para contribuir com melhorias ou correções para esses scripts. Abra uma issue ou um pull request para colaborar.
