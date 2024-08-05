import pandas as pd # Biblioteca utilizada para manipular os dados de excel
import requests # Biblioteca utilizada para manipular os dados de API
import pyodbc # Biblioteca utilizada para manipular o banco de dados
from datetime import datetime
import tkinter as tk # Biblioteca utilizada para manipular a popup de notificação
from tkinter import messagebox

# Função para escrever log
def write_log(message):
    log_file_path = '(Seu diretório)'
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
file_path = '(Seu diretório)'
df = pd.read_excel(file_path)

# Escrever a data atual no início do log
current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
write_log(f'Início do script: {current_time}')

# Passo 2: Configurar a string de conexão
conn_str = (
    r'DRIVER={SQL Server};'
    r'SERVER=Claudio;'
    r'DATABASE=taylor_swift;'
    r'Trusted_Connection=yes;'
)

# Passo 3: Conectar ao banco de dados
try:
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()
    write_log('Conexão ao banco de dados estabelecida com sucesso!')
except pyodbc.Error as e:
    error_message = f'Erro ao tentar conectar ao banco de dados: {e}'
    write_log(error_message)
    show_popup(error_message)
    raise SystemExit(error_message)

# Passo 4: Percorrer as linhas e extrair valores
for index, row in df.iterrows():
    nome = row['nome']
    video_id = row['id']
    era = row['era']
    # Você pode adicionar mais processamento aqui se necessário

    # Passo 5: Realizar chamada de API para obter estatísticas do vídeo
    def get_video_statistics(video_id):
        api_key = '(Sua chave de API)'  # Substitua pela sua chave de API
        url = f'https://www.googleapis.com/youtube/v3/videos?part=statistics&id={video_id}&key={api_key}'
        response = requests.get(url)

        # Imprimir o status code
        print(f'Status Code da API para o vídeo {video_id}, {nome}: {response.status_code}')
        write_log(f'Status Code da API para o vídeo {video_id}, {nome}: {response.status_code}')

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
                # Inserir dados no banco de dados
                insert_query = """
                INSERT INTO youtube_videos (id, video, era, views, likes, data)
                VALUES (?, ?, ?, ?, ?, ?)
                """
                if view_count is not None and like_count is not None:
                    cursor.execute(insert_query, (video_id, nome, era, view_count, like_count, data_atual))
                    conn.commit()
                    write_log(f'Registro inserido no banco de dados com sucesso')
                return view_count, like_count, comment_count
            else:
                error_message = f'Erro ao obter dados do vídeo {video_id}, {nome}: Status Code {response.status_code}'
                write_log(error_message)
                show_popup(error_message)
        return None, None, None

    # Passo 6: Obter estatísticas do vídeo
    view_count, like_count, comment_count = get_video_statistics(video_id)

# Passo 7: Fechar a conexão
cursor.close()
conn.close()

# Escrever "Fim" no final do log com espaçamento
write_log('Fim\n\n----------------------------------------\n\n')
print('Fim\n\n----------------------------------------\n\n')
