import os
import sys
import shutil
import requests
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
import time
import pytest


# Caminhos
pasta_central = r'C:\Users\Aluno\Documents\Pasta_Central'
pasta_exel = r'C:\Users\Aluno\Documents\Exels'
pasta_destino_zip = r'C:\Users\Aluno\Documents\Finalizado'

# Argumentos passados: login_IHX, senha_IHX, login_CAF, senha_CAF, nome_turma, id_arquivo, caminho_arquivo_inprogress
login_IHX = sys.argv[1]
senha_IHX = sys.argv[2]
login_CAF = sys.argv[3]
senha_CAF = sys.argv[4]
nome_turma = sys.argv[5]
id_arquivo = sys.argv[6]  # ID do arquivo que está em progresso
caminho_arquivo_inprogress = sys.argv[7]

# API URL para exclusão do arquivo
url_delete_file = f'http://127.0.0.1:8000/api/documentos/{id_arquivo}/delete-in-progress/'

# Criação da pasta com o nome da turma, se não existir
nova_pasta = os.path.join(pasta_central, nome_turma)
if not os.path.exists(nova_pasta):
    os.makedirs(nova_pasta)

# Copiando o arquivo para a nova pasta
nome_arquivo = os.path.basename(caminho_arquivo_inprogress)
caminho_destino = os.path.join(pasta_exel, nome_arquivo)

try:
    shutil.copy(caminho_arquivo_inprogress, caminho_destino)
    print(f'Arquivo copiado com sucesso para {caminho_destino}!')
except Exception as e:
    print(f'Ocorreu um erro ao copiar o arquivo: {e}')

## selenium
@pytest.mark.selenium
def teste_selenium():

    CAF_page = "https://caf.sesisenaisp.org.br/"

    # Configurando o caminho para o GeckoDriver (Firefox)
    firefox_service = Service(r'C:\Users\Aluno\Downloads\geckodriver\geckodriver.exe')  # Substitua pelo caminho correto
    driver = webdriver.Firefox(service=firefox_service)
    
    try:
        driver.get(CAF_page)
        time.sleep(2)

        campo_login_CAF = driver.find_element(By.XPATH, "/html/body/form/div[3]/div[2]/div[1]/span/div/div[2]/input")
        campo_senha_CAF = driver.find_element(By.XPATH, "/html/body/form/div[3]/div[2]/div[1]/span/div/div[4]/div/input")

        campo_login_CAF.send_keys(login_CAF)
        campo_senha_CAF.send_keys(senha_CAF)
        campo_senha_CAF.send_keys(Keys.RETURN)

        time.sleep(2)

    except Exception as e:
        print(f"Erro ao fazer login: {e}")

    finally:
        driver.quit()









# Compactando a nova pasta e movendo para o diretório finalizado
caminho_zip = os.path.join(pasta_destino_zip, nome_turma)  # Caminho para o arquivo zip final

try:
    # Compacta a pasta em um arquivo .zip
    shutil.make_archive(caminho_zip, 'zip', nova_pasta)
    print(f'Pasta {nova_pasta} compactada com sucesso em {caminho_zip}.zip!')

    # Remove a pasta original que não está compactada
    shutil.rmtree(nova_pasta)
    print(f'Pasta {nova_pasta} removida com sucesso após compactação!')
except Exception as e:
    print(f'Ocorreu um erro ao compactar ou remover a pasta: {e}')

# Excluindo o arquivo da pasta Exels
try:
    if os.path.exists(caminho_destino):
        os.remove(caminho_destino)
        print(f'Arquivo {caminho_destino} removido com sucesso!')
    else:
        print(f'O arquivo {caminho_destino} não existe.')
except Exception as e:
    print(f'Ocorreu um erro ao remover o arquivo: {e}')

# Fazendo requisição para excluir o objeto com status "Em Progresso"
try:
    response = requests.delete(url_delete_file)
    if response.status_code == 204:
        print(f"Objeto com ID {id_arquivo} em progresso excluído com sucesso!")
    else:
        print(f"Erro ao excluir o objeto. Status: {response.status_code}")
        print(f"Resposta: {response.json()}")
except Exception as e:
    print(f"Ocorreu um erro ao tentar excluir o objeto: {e}")

# Criando um novo objeto Fineshed_file via API
url_finished_file = 'http://127.0.0.1:8000/api/concluido/'  # URL para criar o Fineshed_file
zip_path = f"{caminho_zip}.zip"  # Caminho do arquivo compactado

# Enviando a requisição POST para criar o novo Fineshed_file
try:
    with open(zip_path, 'rb') as zip_file:
        files = {'arquivo_fineshed': zip_file}
        data = {'turma': nome_turma}
        response = requests.post(url_finished_file, data=data, files=files)

    if response.status_code == 201:
        print('Novo objeto Fineshed_file criado com sucesso!')
    else:
        print(f'Ocorreu um erro ao criar o objeto Fineshed_file: {response.status_code} - {response.text}')
except Exception as e:
    print(f'Ocorreu um erro ao enviar a requisição para criar o Fineshed_file: {e}')
