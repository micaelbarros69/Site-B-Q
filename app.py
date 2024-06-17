import os
import time
import shutil
import pandas as pd
import dotenv
from flask import Flask, render_template, request, redirect, url_for
from flask_socketio import SocketIO, emit
from werkzeug.utils import secure_filename
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

dotenv.load_dotenv()

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['SECRET_KEY'] = 'secret!'
socketio = SocketIO(app)

# Função para processar arquivos
def processar_arquivos(data):
    tipo = data['tipo']
    folder = data['folder']
    folder_path = os.path.join(app.config['UPLOAD_FOLDER'], folder)
    
    if not os.path.exists(folder_path):
        emit('erro', {'erro': f'Pasta {folder} não encontrada'})
        return
    
    for filename in os.listdir(folder_path):
        filepath = os.path.join(folder_path, filename)
        print(f"Processando arquivo {filename} do tipo {tipo} na pasta {folder}")
        # Aqui você pode adicionar a lógica para processar cada arquivo com base no tipo
        socketio.emit('resultado', {'arquivo': filename, 'status': 'anexado', 'tipo': tipo, 'folder': folder})

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        folder = request.form.get('folder')
        folder_path = os.path.join(app.config['UPLOAD_FOLDER'], folder)
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
        
        files = request.files.getlist('files[]')
        for file in files:
            if file:
                filename = secure_filename(file.filename)
                file.save(os.path.join(folder_path, filename))
        socketio.emit('upload_complete', {'folder': folder, 'files': [file.filename for file in files]})
        return redirect(url_for('upload_file'))
    return render_template('index.html')

@socketio.on('processar_arquivos')
def handle_processar_arquivos(json):
    processar_arquivos(json)


# Função para obter dados do Excel
def obter_dados_do_excel(nome_arquivo='BASE.xlsx', nome_aba='HOME'):
    df = pd.read_excel(nome_arquivo, sheet_name=nome_aba)
    dados = []
    for index, row in df.iterrows():
        link = str(row['Link']).strip()
        file_names = {
            "KML": str(row['KML']).strip(),
            "OT": str(row['OT']).strip(),
            "PRE APR": str(row['PRE APR']).strip(),
            "SGD": str(row['SGD']).strip()
        }
        cod = row['cod']
        for key in file_names:
            if file_names[key].endswith('.0'):
                file_names[key] = file_names[key][:-2]
        if isinstance(cod, float) and cod.is_integer():
            cod = int(cod)
        cod = str(cod)
        if pd.isna(link) or pd.isna(cod):
            continue
        dados.append((link, file_names, cod))
    return dados

# Função para iniciar o driver do Selenium
def iniciar_driver():
    service = Service(ChromeDriverManager().install())
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(service=service, options=options)
    return driver

# Função para acessar link e fazer upload de arquivos
def acessar_link(driver, link, file_path, cod):
    driver.get(link)
    try:
        frame = WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, f'iframe[src*="obras_anexos.php?cod={cod}"]'))
        )
        driver.switch_to.frame(frame)

        input_handle = WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.NAME, 'arq[]'))
        )

        if os.path.exists(file_path):
            input_handle.send_keys(file_path)
            ENVIAR = WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.NAME, 'Salvar_btn'))
            )
            ENVIAR.click()
            WebDriverWait(driver, 10).until(EC.alert_is_present()).accept()
            WebDriverWait(driver, 10).until(EC.alert_is_present()).accept()
            return True  
        else:
            return False  
    except Exception as error:
        return False  

# Função para cancelar arquivos
def cancelar_arquivos(driver, cod, keywords):
    try:
        frame = WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, f'iframe[src*="obras_anexos.php?cod={cod}"]'))
        )
        driver.switch_to.frame(frame)

        arquivos_cancelados = []  
        while True:
            found_file = False  
            try:
                rows = WebDriverWait(driver, 10).until(
                    EC.presence_of_all_elements_located((By.XPATH, "//tr"))
                )

                for row in rows:
                    try:
                        for keyword in keywords:
                            lowercase_keyword = keyword.lower()  
                            link_elements = row.find_elements(By.XPATH, f".//a[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{lowercase_keyword}')]")
                            if link_elements:
                                for link in link_elements:
                                    cancel_button = row.find_element(By.XPATH, ".//img[@title='Cancelar registro']")
                                    cancel_button.click()
                                    WebDriverWait(driver, 10).until(EC.alert_is_present()).accept()
                                    WebDriverWait(driver, 10).until(EC.alert_is_present()).accept()
                                    arquivos_cancelados.append(link.text)
                                    found_file = True
                                    time.sleep(3)
                    except Exception:
                        continue

                if not found_file:
                    break  

            except Exception:
                break

    except Exception:
        pass
    
    return arquivos_cancelados

# Função para obter credenciais
def obter_credenciais():
    email = os.getenv('EMAIL')
    senha = os.getenv('SENHA')
    return email, senha

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    files = request.files.getlist('files[]')
    for file in files:
        file.save(os.path.join(app.config['UPLOAD_FOLDER'], file.filename))
    
    socketio.emit('upload_complete', {'files': [file.filename for file in files]})
    return '', 204

@socketio.on('processar_arquivos')
def processar_arquivos(data):
    try:
        dados_excel = obter_dados_do_excel()
        driver = iniciar_driver()
        driver.get('https://beqce.gpm.srv.br/index.php')
        email, senha = obter_credenciais()

        driver.find_element(By.ID, 'idLogin').send_keys(email)
        driver.find_element(By.ID, 'idSenha').send_keys(senha)
        driver.find_element(By.ID, 'idSenha').send_keys(Keys.ENTER)
        WebDriverWait(driver, 60).until(EC.url_changes('https://beqce.gpm.srv.br/index.php'))

        for link, file_names, cod in dados_excel:
            driver.get(link)
            cancelados = cancelar_arquivos(driver, cod, ["OT"])
            for subpasta, filename in file_names.items():
                if filename != 'nan':
                    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                    sucesso_anexar = acessar_link(driver, link, file_path, cod)
                    if sucesso_anexar:
                        emit('resultado', {'arquivo': filename, 'status': 'anexado'})
                    else:
                        emit('resultado', {'arquivo': filename, 'status': 'não anexado'})

        driver.quit()
    except Exception as e:
        emit('erro', {'erro': str(e)})

if __name__ == "__main__":
    if not os.path.exists(app.config['UPLOAD_FOLDER']):
        os.makedirs(app.config['UPLOAD_FOLDER'])
    socketio.run(app, debug=True)
