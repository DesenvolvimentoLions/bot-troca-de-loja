import os
import time
import glob
import sqlite3
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import Select
import logging
import sqlite3
import csv
import pandas as pd
from datetime import datetime

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('automacao_xml.log'),
        logging.StreamHandler()
    ]
)

def criar_banco():
    conn = sqlite3.connect('veiculos.db')
    cursor = conn.cursor()
    
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS veiculos (
        id TEXT PRIMARY KEY,
        placa TEXT,
        loja_origem TEXT,
        vendedor_selecionado TEXT,
        data_transferencia DATETIME
    )
    ''')
    
    conn.commit()
    conn.close()

def salvar_veiculo(id, placa, loja_origem, vendedor_selecionado):
    if not all([id, placa, loja_origem, vendedor_selecionado]):
        raise ValueError("Todos os campos são obrigatórios")
        
    conn = sqlite3.connect('veiculos.db')
    cursor = conn.cursor()
    
    try:
        cursor.execute('''
        INSERT INTO veiculos (id, placa, loja_origem, vendedor_selecionado, data_transferencia)
        VALUES (?, ?, ?, ?, ?)
        ''', (id, placa, loja_origem, vendedor_selecionado, datetime.now()))
        
        conn.commit()
    
    except sqlite3.IntegrityError:
        cursor.execute('''
        UPDATE veiculos 
        SET placa=?, loja_origem=?, vendedor_selecionado=?, data_transferencia=?
        WHERE id=?
        ''', (placa, loja_origem, vendedor_selecionado, datetime.now(), id))
        
        conn.commit()
    
    except Exception as e:
        logging.error(f"Erro ao salvar veículo: {str(e)}")
        raise
    
    finally:
        conn.close()

def buscar_todos_veiculos():
    conn = sqlite3.connect('veiculos.db')
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    try:
        cursor.execute('SELECT * FROM veiculos ORDER BY data_transferencia DESC')
        veiculos = cursor.fetchall()
        return veiculos
    
    except Exception as e:
        logging.error(f"Erro ao buscar veículos: {str(e)}")
        raise
    
    finally:
        conn.close()

# Criar WebDriver
def iniciar_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-infobars")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)

    return driver

# Função de login
def login(driver):
    try:
        logging.info("Iniciando processo de login...")
        driver.get("https://www.moneycarweb.com.br/default.aspx")

        wait = WebDriverWait(driver, 10)
        email_field = wait.until(EC.presence_of_element_located((By.ID, "txtEmail")))
        email_field.send_keys("thassia@lionsseminovos.com.br")

        password_field = driver.find_element(By.ID, "txtSenhaEmail")
        password_field.send_keys("Yj9s9hbb@")

        login_button = driver.find_element(By.ID, "Button3")
        login_button.click()

        logging.info("Login realizado com sucesso")
        return True
    except Exception as e:
        logging.error(f"Erro no login: {e}")
        return False

# Função para acessar a página de XMLs
"""
lista de lojas 
<select name="ctl00$VNDropDownList1" onchange="trap=true;setTimeout('__doPostBack(\'ctl00$VNDropDownList1\',\'\')', 0)" id="VNDropDownList1" class="Drop" style="font-family:Tahoma;font-size:8.5pt;width:210px;">
	<option value="922">Lions-ATC</option>
	<option value="1042">Lions-BM</option>
	<option value="1080">Lions-CG</option>
	<option value="846">Lions-DC</option>
	<option value="941">Lions-IT</option>
	<option selected="selected" value="628">Lions-MT</option>
	<option value="1081">Lions-NI</option>
	<option value="993">Lions-NT</option>
	<option value="1100">Lions-OS</option>
	<option value="1087">Lions-VP</option>

</select>
"""
def entrarNaPagina(driver):
    try:
        logging.info("Acessando a página de veículos")

        wait = WebDriverWait(driver, 20)
        
        CampoStatusLionsOS = wait.until(EC.element_to_be_clickable((By.ID, "VNDropDownList1")))

        driver.execute_script("""
            var selectElement = arguments[0];
            selectElement.value = "1042";  // Valor de "Lions-OS"
            selectElement.dispatchEvent(new Event('change', { bubbles: true }));
            selectElement.dispatchEvent(new Event('input', { bubbles: true }));
            setTimeout(function() { __doPostBack('ctl00$VNDropDownList1', '') }, 0);
        """, CampoStatusLionsOS)

        logging.info("Status 'Lions-OS' selecionado corretamente.")

        # Clicar no botão Faturamento
        faturamento_btn = wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ToolBar1_i5")))
        faturamento_btn.click()
        logging.info("Botão Faturamento clicado")

        # Aguardar o frame ser carregado e mudar para ele
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "master_frame1")))
        logging.info("Mudou para o frame 'master_frame1'")

        # Aguardar o dropdown estar presente e visível
        CampoStatus = wait.until(EC.element_to_be_clickable((By.ID, "VNDropDownList1")))

        # Alterar o status para "Vendido"
        driver.execute_script("""
            var selectElement = arguments[0];
            selectElement.value = '11';  // Código da opção "VENDIDO"
            selectElement.dispatchEvent(new Event('change', { bubbles: true }));
            selectElement.dispatchEvent(new Event('input', { bubbles: true }));
            setTimeout(function() { __doPostBack('VNDropDownList1', '') }, 0);
        """, CampoStatus)

        logging.info("Status 'Vendido' selecionado corretamente")

        # Alterar o número de resultados para 1000
        results_per_page = wait.until(EC.element_to_be_clickable((By.NAME, "VNGridView1$ctl01$ctl10")))
        select = Select(results_per_page)
        select.select_by_value("1000")
        logging.info("Número de resultados alterado para 1000")

        time.sleep(2)

        # Aguardar botão de pesquisa e clicar
        pesquisar_btn = wait.until(EC.element_to_be_clickable((By.ID, "Button1")))
        driver.execute_script("arguments[0].click();", pesquisar_btn)
        logging.info("Pesquisa iniciada")

        # Aguardar resultados
        tabela = wait.until(EC.presence_of_element_located((By.XPATH, "//table[@id='VNGridView1']//tr[3]")))
        logging.info("Resultados carregados com sucesso")

    except Exception as e:
        logging.error(f"Erro ao entrar na página de XMLs: {e}")

def verificar_e_desbloquear(driver, wait):
    try:
        lock_element = wait.until(EC.presence_of_element_located((By.ID, "ckbBloqueado")))
        src = lock_element.get_attribute("src")
        
        if "lockc.gif" in src:
            logging.info("Veículo bloqueado. Desbloqueando...")
            lock_element.click()
            time.sleep(1)
            return True
        else:
            logging.info("Veículo já está desbloqueado")
            return True
            
    except Exception as e:
        logging.error(f"Erro ao verificar/desbloquear veículo: {e}")
        return False

# [Imports permanecem os mesmos...]

def AbrirArquivo(driver):
    try:
        criar_banco()
        wait = WebDriverWait(driver, 20)
        
        # Loop de paginação
        pagina_atual = 1
        while True:

            rows = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//tr[td/a[contains(@href, 'VeiculoGeral.aspx?id=')]]")))
            logging.info(f"Processando página {pagina_atual} - Total de veículos na página: {len(rows)}")

            for i in range(1, len(rows)):
                try:
                    row = rows[i]
                    
                    # Capturar ID e placa
                    id_element = row.find_element(By.XPATH, ".//td[2]")
                    placa_element = row.find_element(By.XPATH, ".//td[3]")
                    
                    vehicle_id = id_element.text
                    placa = placa_element.text
                    
                    logging.info(f"Processando veículo ID: {vehicle_id}, Placa: {placa}")
                    
                    # Clica no link de edição
                    edit_link = row.find_element(By.XPATH, f".//td/a[contains(@href, 'VeiculoGeral.aspx?id={vehicle_id}')]")
                    edit_link.click()
                    
                    # Espera a nova aba abrir
                    wait.until(lambda d: len(d.window_handles) > 1)
                    
                    # Muda para a nova aba
                    driver.switch_to.window(driver.window_handles[-1])
                    
                    # Espera o iframe carregar
                    wait = WebDriverWait(driver, 30)
                    
                    # Contexto padrão
                    driver.switch_to.default_content()
                    
                    # Iframe tabs_iframe
                    iframe = wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "tabs_iframe")))
                    
                    # Verifica e desbloqueia
                    if not verificar_e_desbloquear(driver, wait):
                        logging.error(f"Não foi possível verificar/desbloquear o veículo {vehicle_id}. Pulando para o próximo...")
                        driver.close()
                        driver.switch_to.window(driver.window_handles[0])
                        driver.switch_to.default_content()
                        driver.switch_to.frame("master_frame1")
                        continue

                    time.sleep(1)

                    # Processamento do vendedor
                    try:
                        wait = WebDriverWait(driver, 10)
                        
                        driver.switch_to.default_content()
                        
                        logging.info("Aguardando e clicando na aba de vendas.")
                        vendas = wait.until(EC.element_to_be_clickable((By.ID, "tabHeader_3")))
                        driver.execute_script("arguments[0].click();", vendas)
                        
                        iframe = wait.until(EC.presence_of_element_located((By.ID, "tabs_iframe")))
                        driver.switch_to.frame(iframe)
                        
                        logging.info("Aguardando o dropdown de vendedores.")
                        Vendedores = wait.until(EC.presence_of_element_located((By.ID, "VNDropDownListVendedores")))
                        
                        select = Select(Vendedores)
                        vendedor_selecionado = select.first_selected_option.text
                        logging.info(f"Vendedor selecionado: {vendedor_selecionado}")
                        
                        if vendedor_selecionado and vendedor_selecionado != "Selecione...":
                            salvar_veiculo(vehicle_id, placa, "Lions-BM", vendedor_selecionado)
                            logging.info(f"Dados do veículo ID: {vehicle_id}, Placa: {placa}, Vendedor: {vendedor_selecionado} salvos no banco")
                            
                            # Remover vendedor
                            select.select_by_value("0")
                            
                            salvar = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="form1"]/div[6]/div[17]/table/tbody/tr/td[8]/a')))
                            driver.execute_script("arguments[0].click();", salvar)
                            logging.info(f"Vendedor removido do veículo {vehicle_id}")
                        
                    except Exception as e:
                        logging.error(f"Erro ao processar vendedor do veículo {vehicle_id}: {e}")
                        continue

                    # Processamento da loja
                    try:
                        driver.switch_to.default_content()
                        
                        logging.info("Aguardando e clicando na aba inicial.")
                        home = wait.until(EC.element_to_be_clickable((By.ID, "tabHeader_1")))
                        driver.execute_script("arguments[0].click();", home)
                        
                        iframe = wait.until(EC.presence_of_element_located((By.ID, "tabs_iframe")))
                        driver.switch_to.frame(iframe)

                        lojas = wait.until(EC.element_to_be_clickable((By.ID, "DropTroca")))
                        select = Select(lojas)
                        select.select_by_value("628")  # Lions-MT

                        time.sleep(1)

                        driver.execute_script("""
                            var event = new Event('change', { bubbles: true });
                            arguments[0].dispatchEvent(event);
                        """, lojas)
                                    
                        logging.info(f"Loja do veículo {vehicle_id} alterada.")
                        
                        botaotr = wait.until(EC.element_to_be_clickable((By.ID, "btTroca")))
                        driver.execute_script("arguments[0].click();", botaotr)

                        alert = WebDriverWait(driver, 10).until(EC.alert_is_present())
                        alert.accept()

                        time.sleep(2)
                        
                    except Exception as e:
                        logging.error(f"Erro ao processar troca de loja do veículo {vehicle_id}: {e}")

                    finally:
                        driver.close()
                        driver.switch_to.window(driver.window_handles[0])
                        driver.switch_to.default_content()
                        driver.switch_to.frame("master_frame1")

                except Exception as e:
                    logging.error(f"Erro ao processar o veículo: {e}{vehicle_id}")
                    if len(driver.window_handles) > 1:
                        driver.close()
                        driver.switch_to.window(driver.window_handles[0])
                        driver.switch_to.default_content()
                        driver.switch_to.frame("master_frame1")

            # Verifica se existe próxima página
            try:
                next_page = wait.until(EC.presence_of_element_located((By.XPATH, f"//a[contains(@href, 'Page${pagina_atual + 1}')]")))
                pagina_atual += 1
                driver.execute_script("arguments[0].click();", next_page)

                wait.until(EC.presence_of_element_located((By.XPATH, "//table[@id='VNGridView1']//tr[3]")))                
                # Volta para o frame correto após a mudança de página
                driver.switch_to.default_content()
                driver.switch_to.frame("master_frame1")
                
                logging.info(f"Avançando para página {pagina_atual}")
                
            except Exception as e:
                logging.info("Não há mais páginas para processar")
                break

    except Exception as e:
        logging.error(f"Erro geral no processamento: {e}")
        if len(driver.window_handles) > 1:
            driver.close()
            driver.switch_to.window(driver.window_handles[0])

def exibir_dados():
    try:
        veiculos = buscar_todos_veiculos()
        print("\nVeículos registrados:")
        print("ID | Placa | Loja Origem | Vendedor | Data Transferência")
        print("-" * 60)
        for veiculo in veiculos:
            print(f"{veiculo['id']} | {veiculo['placa']} | {veiculo['loja_origem']} | {veiculo['vendedor_selecionado']} | {veiculo['data_transferencia']}")
    except Exception as e:
        logging.error(f"Erro ao exibir dados: {str(e)}")
        print("Erro ao exibir os dados dos veículos.")
def exportar_dados(tipo_arquivo='csv', nome_arquivo='veiculos_export'):
    try:
        # Aproveita a função que já busca os dados
        veiculos = buscar_todos_veiculos()
        
        if tipo_arquivo.lower() == 'csv':
            nome_arquivo = f"{nome_arquivo}.csv"
            with open(nome_arquivo, 'w', newline='', encoding='utf-8') as arquivo:
                writer = csv.writer(arquivo)
                # Cabeçalho
                writer.writerow(['ID', 'Placa', 'Loja Origem', 'Vendedor', 'Data Transferência'])
                # Dados
                for veiculo in veiculos:
                    writer.writerow([
                        veiculo['id'],
                        veiculo['placa'],
                        veiculo['loja_origem'],
                        veiculo['vendedor_selecionado'],
                        veiculo['data_transferencia']
                    ])
            logging.info(f"Dados exportados com sucesso para {nome_arquivo}")
            
        elif tipo_arquivo.lower() == 'xlsx':
            nome_arquivo = f"{nome_arquivo}.xlsx"
            # Convertendo para formato que o pandas entende
            dados = []
            for veiculo in veiculos:
                dados.append({
                    'ID': veiculo['id'],
                    'Placa': veiculo['placa'],
                    'Loja Origem': veiculo['loja_origem'],
                    'Vendedor': veiculo['vendedor_selecionado'],
                    'Data Transferência': veiculo['data_transferencia']
                })
            
            # Criando DataFrame e salvando
            df = pd.DataFrame(dados)
            df.to_excel(nome_arquivo, index=False)
            logging.info(f"Dados exportados com sucesso para {nome_arquivo}")
            
    except Exception as e:
        logging.error(f"Erro ao exportar dados: {str(e)}")
        print(f"Erro ao exportar os dados dos veículos: {str(e)}")
def main():
    driver = iniciar_driver()

    if login(driver):
        entrarNaPagina(driver)
        AbrirArquivo(driver)
        exibir_dados()  # Mostra os dados após terminar o processamento
        exportar_dados(tipo_arquivo='xlsx', nome_arquivo='veiculos_export')  # Exporta os dados após exibir

    else:
        logging.error("Login falhou. Encerrando script.")

if __name__ == "__main__":
    # inicializar_banco()  # Adicione esta linha antes do main()

    main()