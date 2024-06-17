import os
import time
import pandas as pd
import dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from colorama import Fore, Style, init
from tkinter import Tk, filedialog, messagebox

# Inicializar colorama
init(autoreset=True)

# Carregar variáveis de ambiente
dotenv.load_dotenv()

print('Bem-vindo Micael Barros')

# Função para selecionar diretório
def selecionar_diretorio():
    root = Tk()
    root.withdraw()  # Esconde a janela principal
    messagebox.showinfo("Seleção de Diretório", "Escolha o diretório de trabalho")
    selected_directory = filedialog.askdirectory()
    root.destroy()
    return selected_directory

# Perguntar ao usuário se deseja usar o diretório padrão ou selecionar um novo
def obter_diretorio():
    usar_diretorio_padrao = input("Deseja usar o diretório padrão? (s/n): ").strip().lower()
    if usar_diretorio_padrao == 's':
        return r'C:\Users\FRANCISCO.SANTOS\Desktop\MICAEL\PROJETO\GPM'
    else:
        return selecionar_diretorio()

# Corrigir o caminho do diretório
directory_path = obter_diretorio()

# Concatenar o nome da pasta "Arquivos" ao diretório principal
directory_path = os.path.join(directory_path, "Arquivos")

# Verificação do caminho do diretório
print(f'Verificando o diretório: {directory_path}')
if os.path.exists(directory_path):
    print(f'O diretório {directory_path} existe.')
else:
    print(f'O diretório {directory_path} não existe.')
    exit()

# Subpastas a serem consideradas
subpastas = ["KML", "OT", "PRE APR", "SGD"]

# Carregar os arquivos de todas as subpastas inicialmente
try:
    arquivos = {}
    for subpasta in subpastas:
        caminho_subpasta = os.path.join(directory_path, subpasta)
        arquivos[subpasta] = os.listdir(caminho_subpasta)
        print(f'Arquivos encontrados na subpasta {subpasta}:', arquivos[subpasta])
except Exception as err:
    print('Erro ao acessar o diretório:', err)
    exit()

def obter_dados_do_excel(nome_arquivo='BASE.xlsx', nome_aba='HOME', linha=2):
    df = pd.read_excel(nome_arquivo, sheet_name=nome_aba)

    # Imprimir os nomes das colunas para depuração
    print("Nomes das colunas no DataFrame:", df.columns)

    if linha > len(df):
        raise IndexError(f"A linha {linha} não existe no arquivo {nome_arquivo}")

    link = str(df.at[linha-1, 'Link']).strip()  # Remover espaços extras
    file_names = {
        "KML": str(df.at[linha-1, 'KML']).strip(),  # Remover espaços extras
        "OT": str(df.at[linha-1, 'OT']).strip(),    # Remover espaços extras
        "PRE APR": str(df.at[linha-1, 'PRE APR']).strip(),  # Remover espaços extras
        "SGD": str(df.at[linha-1, 'SGD']).strip()   # Remover espaços extras
    }
    cod = df.at[linha-1, 'cod']
    
    # Adicionar debug prints
    print(f"Lendo linha {linha}: link = {link}, file_names = {file_names}, cod = {cod}")
    
    # Remover o sufixo ".0" dos nomes dos arquivos e do código
    for key in file_names:
        if file_names[key].endswith('.0'):
            file_names[key] = file_names[key][:-2]

    if isinstance(cod, float) and cod.is_integer():
        cod = int(cod)
    cod = str(cod)

    if pd.isna(link):
        raise ValueError(f"A célula 'Link' na linha {linha} está vazia ou não foi encontrada.")
    
    if pd.isna(cod):
        raise ValueError(f"A célula 'cod' na linha {linha} está vazia ou não foi encontrada.")
    
    return link, file_names, cod

def procurar_arquivo(subpasta, file_name):
    # Procurar o arquivo diretamente na subpasta específica
    for file in arquivos[subpasta]:
        if str(file_name) in file:
            return os.path.join(directory_path, subpasta, file)
    return None

def acessar_link(driver, link, file_path, cod):
    driver.get(link)   
    try:
        # Esperar o iframe aparecer e mudar para ele
        frame = WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, f'iframe[src*="obras_anexos.php?cod={cod}"]'))
        )
        driver.switch_to.frame(frame)
        print('Iframe encontrado em acessar_link.')

        # Esperar pelo input dentro do iframe
        input_handle = WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.NAME, 'arq[]'))
        )
        print('Input encontrado em acessar_link.')

        if os.path.exists(file_path):
            input_handle.send_keys(file_path)
            print(f'Upload do arquivo {file_path} realizado.')

            # Clique no botão "Salvar Anexos"
            ENVIAR = WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.NAME, 'Salvar_btn'))
            )
            ENVIAR.click()
            print('Botão "Salvar Anexos" clicado.')
            
            # Esperar alguns segundos para os alertas aparecerem e serem tratados
            WebDriverWait(driver, 10).until(EC.alert_is_present()).accept()
            print('Primeiro alerta tratado.')

            WebDriverWait(driver, 10).until(EC.alert_is_present()).accept()
            print('Segundo alerta tratado.')
            return True  # Indicar que o arquivo foi anexado com sucesso
            
        else:
            print(f'Arquivo {file_path} não encontrado.')
            return False  # Indicar que o arquivo não foi encontrado
    except Exception as error:
        print(f'Erro ao esperar pelo seletor dentro do iframe em acessar_link: {error}')
        return False  # Indicar que ocorreu um erro

def cancelar_arquivos(driver, cod, keywords):
    try:
        # Encontrar o iframe específico e mudar o contexto para ele
        frame = WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, f'iframe[src*="obras_anexos.php?cod={cod}"]'))
        )
        driver.switch_to.frame(frame)
        print('Iframe encontrado em cancelar_arquivos.')

        arquivos_cancelados = []  # Lista para armazenar os arquivos cancelados

        while True:
            found_file = False  # Flag para verificar se um arquivo com a palavra-chave foi encontrado e cancelado
            try:
                # Encontrar todos os elementos <tr> dentro do iframe
                rows = WebDriverWait(driver, 10).until(
                    EC.presence_of_all_elements_located((By.XPATH, "//tr"))
                )

                for row in rows:
                    try:
                        # Procurar o link que contém a palavra-chave dentro do <tr>
                        for keyword in keywords:
                            lowercase_keyword = keyword.lower()  # Convert keyword to lowercase
                            link_elements = row.find_elements(By.XPATH, f".//a[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{lowercase_keyword}')]")
                            if link_elements:
                                for link in link_elements:
                                    # Se o link for encontrado, buscar o botão 'Cancelar registro' correspondente
                                    cancel_button = row.find_element(By.XPATH, ".//img[@title='Cancelar registro']")
                                    cancel_button.click()
                                    
                                    # Esperar pelo primeiro alerta e aceitá-lo
                                    WebDriverWait(driver, 10).until(EC.alert_is_present()).accept()
                                    print('Primeiro alerta tratado.')

                                    # Esperar pelo segundo alerta e aceitá-lo
                                    WebDriverWait(driver, 10).until(EC.alert_is_present()).accept()
                                    print('Segundo alerta tratado.')

                                    # Adicionar o arquivo cancelado à lista
                                    arquivos_cancelados.append(link.text)
                                    print(f"O arquivo {link.text} foi cancelado.")
                                    found_file = True
                                    
                                    # Esperar 3 segundos para a tela atualizar
                                    time.sleep(3)
                        
                    except Exception as e:
                        # Se não encontrar um link ou botão no <tr> específico, continuar para o próximo
                        continue

                if not found_file:
                    print(f"Nenhum arquivo com as palavras-chave {keywords} encontrado. Saindo do loop.")
                    break  # Sair do loop se nenhum arquivo com a palavra-chave foi encontrado nesta iteração

            except Exception as e:
                # Se não encontrar mais elementos <tr>, sair do loop
                print("Erro ou nenhum elemento <tr> encontrado. Saindo do loop.")
                break

    except Exception as e:
        print(f'Erro ao mudar para o iframe em cancelar_arquivos: {e}')
    
    print(f'Todos os arquivos com as palavras-chave {keywords} foram cancelados.')
    return arquivos_cancelados


def obter_credenciais():
    email = os.getenv('EMAIL')
    senha = os.getenv('SENHA')
    return email, senha

def robo():
    print("Iniciando o robô...")

    # Menu para o usuário escolher a operação
    print(Fore.RED + "(1) Apenas cancelar")
    print(Fore.BLUE + "(2) Apenas anexar")
    print(Fore.GREEN + "(3) Ambos")
    
    while True:
        operacao = input("Você deseja: (1/2/3): ").strip()
        if operacao in ['1', '2', '3']:
            break
        else:
            print(Fore.YELLOW + "Escolha inválida. Tente novamente.")

    anexados = []
    nao_anexados = []
    cancelados = []
    nao_cancelados = []

    service = Service(ChromeDriverManager().install())
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(service=service, options=options)

    driver.get('https://beqce.gpm.srv.br/index.php')
    email, senha = obter_credenciais()

    driver.find_element(By.ID, 'idLogin').send_keys(email)
    driver.find_element(By.ID, 'idSenha').send_keys(senha)
    driver.find_element(By.ID, 'idSenha').send_keys(Keys.ENTER)

    WebDriverWait(driver, 60).until(EC.url_changes('https://beqce.gpm.srv.br/index.php'))

    linha = 1

    while True:
        try:
            link, file_names, cod = obter_dados_do_excel(linha=linha)

            for subpasta in subpastas:
                file_name = file_names[subpasta]
                if file_name != 'nan':
                    arquivo_encontrado = procurar_arquivo(subpasta, file_name)

                    if arquivo_encontrado:
                        print(f'Acessando link da linha {linha}:', link)
                        print(f'Procurando arquivo da linha {linha}:', arquivo_encontrado)
                        print(f'Procurando cod da linha {linha}:', cod)
                        driver.get(link)

                        if operacao in ['1', '3']:
                            # Atualize a chamada para cancelar arquivos com a palavra-chave correta
                            arquivos_cancelados = cancelar_arquivos(driver, cod, [ subpasta])
                            cancelados.extend(arquivos_cancelados)
                            if not arquivos_cancelados:
                                nao_cancelados.append(f"Nenhum arquivo {subpasta} encontrado na linha {linha}")

                        if operacao in ['2', '3']:
                            sucesso_anexar = acessar_link(driver, link, arquivo_encontrado, cod)
                            if sucesso_anexar:
                                anexados.append(os.path.basename(arquivo_encontrado))
                            else:
                                nao_anexados.append(os.path.basename(arquivo_encontrado))

            linha += 1

        except Exception as e:
            print(f"Erro ao processar linha {linha}: {e} processo encerrado")
            break

    print("\nResultados:")
    print(Fore.GREEN + "Arquivos anexados:")
    for arquivo in anexados:
        print(Fore.GREEN + arquivo)
    print(Fore.RED + "Arquivos não anexados:")
    for arquivo in nao_anexados:
        print(Fore.RED + arquivo)
    print(Fore.GREEN + "Arquivos cancelados:")
    for arquivo in cancelados:
        print(Fore.GREEN + arquivo)
    print(Fore.RED + "Arquivos não cancelados:")
    for arquivo in nao_cancelados:
        print(Fore.RED + arquivo)
    print(Style.RESET_ALL)

if __name__ == "__main__":
    robo()

