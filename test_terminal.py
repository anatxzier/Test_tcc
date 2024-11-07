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
from selenium.webdriver.common.action_chains import ActionChains
import re
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import Select
import pyautogui
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import openpyxl

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
    IHX_page = "http://sn502actrcp1:8000/admin/"
    #mudar

    wait = WebDriverWait(driver, 10) 

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

            # Navegando para a aba de pessoas
        WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/form/div[4]/span/ul/li[4]/a"))
        ).click()

        workbook = openpyxl.load_workbook(caminho_destino)
        sheet = workbook.active

        for row in sheet.iter_rows(min_row=2, values_only=True):
            print(row)  # Imprime a linha completa para verificar os dados
            identificacao = row[4]  # Acessa a quinta coluna, ajustando conforme necessário
            print(f"Identificação: {identificacao}")  # Imprime o valor específico

            procurar_pessoa = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/form/div[7]/div/span/div/div[2]/div[1]/input"))
            )

            procurar_pessoa.clear()
            procurar_pessoa.send_keys(identificacao)
            procurar_pessoa.send_keys(Keys.RETURN)

            WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.ID, "btnCarteirinhas"))
            ).click()
            print("Carteirinha clicada.")

            WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.CSS_SELECTOR, ".ui-dialog")))

            # Espera e troca para o iframe
            iframe = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "iframe"))
            )
            driver.switch_to.frame(iframe)

            time.sleep(2)

            # Localizar o elemento pelo ID (no caso, "dssGridViewExpanderButton")
            toggle = driver.find_elements(By.ID, "dssGridViewExpanderButton")
            ultimo_toggle = toggle[-1]

            # Clicar no elemento para exibir mais informações
            ultimo_toggle.click()

            #copiar e salvar numero de matricula dos alunos

            

            time.sleep(2)
            div_elements = driver.find_elements(By.TAG_NAME, "div")

            div_element = div_elements[-8]
            print(f"ESSE: {div_element.text}")
            # Extrair o texto completo do elemento
            texto_completo = div_element.text

            # Usar regex para capturar o código numérico após "Código de Barras Traseiro:"
            match = re.search(r"Código de Barras Traseiro: \s*(\d+)", texto_completo)

            if match:
                codigo_barras = match.group(1)
                print("Código de Barras Traseiro:", codigo_barras)

                with open("codes.txt", "a") as file:
                    file.write(f"{codigo_barras}\n")
            else:
                print("Código de Barras Traseiro não encontrado.")

            buttons = WebDriverWait(driver, 15).until(
                EC.presence_of_all_elements_located((By.ID, "btnEmitir"))
            )

            button_dict = {}
            for button in buttons:
                button_name = button.get_attribute("name")
                try:
                    suffix = int(button_name.split('$')[4][3:])
                    button_dict[suffix] = button
                except (IndexError, ValueError):
                    continue

            # Ordena os botões pelo sufixo em ordem decrescente
            sorted_suffixes = sorted(button_dict.keys(), reverse=True)

            btn_to_click = None
            for suffix in sorted_suffixes:
                btn_to_click = button_dict[suffix]
                
                # Verifica se o botão está habilitado
                if btn_to_click.is_enabled():
                    btn_to_click.click()
                    print(f"Clicado o botão habilitado com o sufixo: {suffix}")
                    break
            else:
                print("Nenhum botão habilitado encontrado para clique.")

################### faça aqui
################### dar um jeito de selecionar essa modal e depois sim

        #clicar em sim

        #clicar no botão verde de ativar

            time.sleep(2)

            driver.refresh()

            time.sleep(2)


        try:
            driver.get(IHX_page)

            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/div[4]/form/div[1]/input"))
            )

            input1_IHX = driver.find_element(By.XPATH, "/html/body/div[4]/form/div[1]/input")
            

            input2_IHX = driver.find_element(By.XPATH, "/html/body/div[4]/form/div[2]/input[1]")
            
            time.sleep(3)
            input1_IHX.send_keys(login_IHX)
            input2_IHX.send_keys(senha_IHX)
            input2_IHX.send_keys(Keys.RETURN)
            time.sleep(3)
            print("login feito com sucesso")

            element_IHX = driver.find_element(By.XPATH, "/html/body/div/article/div[2]/center/div/div[2]/div")
            # Mover o cursor até o elemento para acionar o evento de hover
            actions = ActionChains(driver)
            actions.move_to_element(element_IHX).perform()

            time.sleep(2)
            image_element = driver.find_element(By.XPATH, "/html/body/div/article/div[2]/center/div/div[2]/div/table/tbody/tr[2]/td[3]")
            image_element.click()

            time.sleep(2)
            try:
                input_aluno_element = driver.find_element(By.NAME, "q")
                print('Campo de entrada encontrado.')
            except Exception as e:
                print(f'Campo de entrada não encontrado: {e}')
                driver.quit()  # Encerra o navegador se o campo de entrada não for encontrado
                exit()

            # Abre o arquivo de códigos e lê todos os códigos
            with open("codes.txt", "r") as file:
                codes = [code.strip() for code in file.readlines()]  # Remove espaços em branco ou quebras de linha

            # Tenta buscar cada código na lista
            for code in codes:
                attempts = 0
                while attempts < 3:  # Limita as tentativas para reencontrar o elemento se houver erro
                    try:
                        # Aguardando o campo de pesquisa
                        input_aluno_element = wait.until(EC.presence_of_element_located((By.NAME, "q")))
                        input_aluno_element.clear()
                        input_aluno_element.send_keys(code)
                        input_aluno_element.send_keys(Keys.RETURN)
                        print(f"Código {code} pesquisado.")
                        time.sleep(2)

                        # Espera o link do aluno ser encontrado
                        elemento_nome = wait.until(EC.presence_of_element_located((By.XPATH, "//a[contains(@href, '/admin/MODAluno/aluno/')]")))
                        elemento_nome.click()

                        # Navega até o botão QR Code e clica
                        qrcode_element = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/article/header/ul/li[1]/a")))
                        qrcode_element.click()

                        driver.switch_to.window(driver.window_handles[-1])

                        time.sleep(2)
                        
                        pyautogui.keyDown('ctrl')
                        pyautogui.press('p')
                        pyautogui.keyUp('ctrl')


                        time.sleep(2)
                        pyautogui.press('UP', presses=2)
                        pyautogui.press('TAB', presses=4)
                        pyautogui.press("ENTER")

                    
    
                        # Localiza e clica no botão de salvar após a seleção
                        button_save = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.XPATH, "/html/body/print-preview-app//print-preview-sidebar//print-preview-button-strip//div/cr-button[1]"))
                        )
                        button_save.click()

                        # Retorna ao contexto principal fora do iframe, se necessário
                        driver.switch_to.default_content()

                        driver.close()
                    
                        driver.switch_to.window(driver.window_handles[0])
                        
                        break 
                    except (StaleElementReferenceException, NoSuchElementException) as e:
                        print(f"Elemento não encontrado no DOM para o código {code}. Tentando novamente...")
                        attempts += 1
                        time.sleep(1)  # Tempo de espera para tentar novamente
                    except Exception as e:
                        print(f"Erro ao pesquisar o código {code}: {e}")
                        break  # Sai do loop se houver outro tipo de erro
        except Exception as e:
            print("Falhou algo:", e)
    finally:
        driver.quit()



#####################

            ### Clicar no botão de sair DESCOMENTAR DEPOIS (sair do iframe)

            # try:
            #     sair_element = WebDriverWait(driver, 15).until(
            #         EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/button/span[1]"))
            #     )
            #     sair_element.click()
            #     print("Logout realizado com sucesso.")
            # except Exception as e:
            #     print(f"Erro ao clicar em 'Sair': {e}")
##################################################




# finally:
#         driver.quit()








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
