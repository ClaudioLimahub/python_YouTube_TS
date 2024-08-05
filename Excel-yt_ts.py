import pandas as pd  # Biblioteca utilizada para manipular os dados de excel
import requests  # Biblioteca utilizada para manipular os dados de API
from datetime import datetime
import tkinter as tk  # Biblioteca utilizada para manipular a popup de notificação
from tkinter import messagebox

# Função para escrever log
def write_log(message):
    log_file_path = 'C:/Users/gblsj/OneDrive/Área de Trabalho/ts/python/youtube_videos/log.txt'
    with open(log_file_path, 'a', encoding='utf-8') as log_file:
        log_file.write(message + '\n')

# Função para exibir uma mensagem de pop-up
def show_popup(message):
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal
    messagebox.showwarning("Aviso", message)
    root.destroy()  # Fecha a janela após clicar em "OK"

# Obtém a data atual no formato dd/MM/yyyy
data_atual = datetime.now().strftime('%d/%m/%Y')

# Passo 1: Carregar o arquivo Excel
file_path = 'C:/Users/gblsj/OneDrive/Documentos/Python/RPA/ts_web_db_youtube/Videos TS.xlsx'
df = pd.read_excel(file_path)

# Passo 2: Carregar o arquivo Excel onde serão armazenadas as estatísticas
file_path_stats = 'C:/Users/gblsj/OneDrive/Documentos/Python/RPA/ts_web_db_youtube/dados_TS_YT.xlsx'
df_stats = pd.read_excel(file_path_stats)

# Adicionar colunas para estatísticas, se não existirem
if 'views' not in df_stats.columns:
    df_stats['views'] = pd.NA
if 'likes' not in df_stats.columns:
    df_stats['likes'] = pd.NA
if 'data' not in df_stats.columns:
    df_stats['data'] = pd.NA

# Escrever a data atual no início do log
current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
write_log(f'Início do script: {current_time}')

# Função para obter estatísticas do vídeo
def get_video_statistics(video_id):
    api_key = 'AIzaSyBQo8bdQ3TQI9iQQCJ_H48L4gtUBw5BGZg'  # Substitua pela sua chave de API
    url = f'https://www.googleapis.com/youtube/v3/videos?part=statistics&id={video_id}&key={api_key}'
    response = requests.get(url)

    # Imprimir o status code
    print(f'Status Code da API para o vídeo {nome} - {video_id}: {response.status_code}')
    write_log(f'Status Code da API para o vídeo {nome} - {video_id}: {response.status_code}')

    if response.status_code == 200:
        data = response.json()
        if 'items' in data and len(data['items']) > 0:
            stats = data['items'][0]['statistics']
            view_count = stats.get('viewCount', 0)
            like_count = stats.get('likeCount', 0)
            comment_count = stats.get('commentCount', 0)
            # Escrever a data atual no início do log
            print(f'Estatísticas do vídeo {video_id}:')
            write_log(f'Estatísticas do vídeo {video_id}:')
            print(f'  - View Count: {view_count}')
            write_log(f'  - View Count: {view_count}')
            print(f'  - Like Count: {like_count}')
            write_log(f'  - Like Count: {like_count}')
            print(f'  - Comment Count: {comment_count}')
            write_log(f'  - Comment Count: {comment_count}')
            return view_count, like_count, comment_count
        else:
            error_message = f'Erro ao obter dados do vídeo {video_id}: Status Code {response.status_code}'
            write_log(error_message)
            show_popup(error_message)
    return None, None, None

# Passo 3: Atualizar o DataFrame com as novas estatísticas
for index, row in df.iterrows():
    nome = row['nome']
    video_id = row['id']
    era = row['era']

    # Obter estatísticas do vídeo
    view_count, like_count, comment_count = get_video_statistics(video_id)

    # Adicionar nova linha ao DataFrame de estatísticas
    new_row = {
        'id': video_id,
        'video': nome,
        'era': era,
        'views': view_count,
        'likes': like_count,
        'data': data_atual
    }
    df_stats = df_stats._append(new_row, ignore_index=True)

# Passo 4: Salvar o DataFrame atualizado de volta no mesmo arquivo Excel
df_stats.to_excel(file_path_stats, index=False)

# Escrever "Fim" no final do log com espaçamento
write_log('Fim\n\n----------------------------------------\n\n')
print('Fim\n\n----------------------------------------\n\n')
