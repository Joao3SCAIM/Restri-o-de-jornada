from imports import *


edge_driver_path = r"C:\Users\joao.escaim\Downloads\edgedriver_win64\msedgedriver.exe"
from selenium.webdriver.edge.service import Service
driver_service = Service(edge_driver_path)
options = webdriver.EdgeOptions()
options.add_argument("--start-maximized")
driver = webdriver.Edge(service=driver_service, options=options)
wait = WebDriverWait(driver, 30)

load_dotenv()
USER_PORTAL = os.getenv("USER_PORTAL")
PASSWORD_PORTAL = os.getenv("PASSWORD_PORTAL")

def fazer_login():
    print("Realizando login...")
    driver.get("https://portal.gpssa.com.br/GPS/Login.aspx")
    time.sleep(3)
    driver.find_element(By.ID, "txtUsername-inputEl").send_keys(USER_PORTAL)
    driver.find_element(By.ID, "txtPassword-inputEl").send_keys(PASSWORD_PORTAL)
    time.sleep(1)
    driver.find_element(By.ID, "btnLogin-btnInnerEl").click()
    time.sleep(3)
    print("Login concluído.")

def processar_dispositivo(imei, crs):
    print(f"\nProcessando dispositivo - IMEI: {imei}")
    driver.get("https://portal.gpssa.com.br/SAD/Dispositivos?IDPAGINA=562&GRUPOABA=SAD%20-%20Cadastro%20de%20Dispositivos&_dc=1743508635081")
    
    # Esperar o campo de busca de IMEI carregar
    wait.until(EC.presence_of_element_located((By.ID, "textfield-1027-inputEl")))
    time.sleep(2)
    
    campo_imei = driver.find_element(By.ID, "textfield-1027-inputEl")
    campo_imei.clear()
    campo_imei.send_keys(imei)
    time.sleep(5)  # aguarda o resultado da busca carregar
    
    try:
        # Captura a primeira linha da tabela
        primeira_linha = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'tr.x-grid-row')))
        
        # Pega as colunas da linha
        colunas = primeira_linha.find_elements(By.TAG_NAME, 'td')
        
        imei_encontrado = None
        for td in colunas:
            texto = td.text.strip()
            if texto.isdigit() and len(texto) >= 15:
                imei_encontrado = texto
                break
        
        if imei_encontrado != imei:
            print(f"⚠️ IMEI na tabela ({imei_encontrado}) diferente do preenchido ({imei}). Pulando.")
            return
        
        print(f"✅ IMEI confere ({imei_encontrado}). Clicando em Editar...")
        botao_editar = primeira_linha.find_element(By.CSS_SELECTOR, 'a[role="button"] .icon-pencil')
        botao_editar.click()
        time.sleep(3)
        
        print("Adicionando CRs...")
        elemento_cr = driver.find_element(By.ID, "cbbQueryCr-inputEl")
        cr_list = [cr.strip()[:5] for cr in crs.split(',')]  # considera CRs separados por vírgula
        
        for cr in cr_list:
            if cr:
                elemento_cr.clear()
                elemento_cr.send_keys(cr)
                time.sleep(3)
                elemento_cr.send_keys(Keys.ENTER)
                time.sleep(3)
        
            # Selecionar checkbox (se necessário)
            try:
                checkbox = driver.find_element(By.XPATH, "//span[contains(@class, 'x-column-header-checkbox')]")
                checkbox.click()
            except:
                pass
            
            botoes = driver.find_elements(By.XPATH, "//span[contains(@class, 'icon-pencil')]")
            
            if len(botoes) >= 2:
                botoes[1].click()
            else:
                print("Menos de dois botões encontrados.")

            try:
                checkbox_label = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//label[@for='chkSelf-inputEl']"))
                )

                # Verifica se está marcado pelo input hidden
                try:
                    driver.find_element(By.XPATH, "//input[@type='hidden' and @name='chkSelf']")
                    # Se marcado, clica duas vezes
                    checkbox_label.click()
                    checkbox_label.click()
                except:
                    # Se desmarcado, clica uma vez
                    checkbox_label.click()

            except Exception as e:
                print("Erro ao clicar no checkbox via label:", e)

            # Salvar alterações
            print("Salvando alterações...")
            save_button = driver.find_element(By.XPATH, "//span[@data-ref='btnInnerEl' and contains(@class, 'x-btn-inner') and text()='Salvar']")
            save_button.click()
            time.sleep(10)
        print(f"Processamento concluído para IMEI {imei}")
        
    except TimeoutException:
        print("Nenhuma linha encontrada na tabela após busca do IMEI. Pulando.")
    except NoSuchElementException:
        print("Elemento esperado não encontrado. Pulando.")
    except Exception as e:
        print(f"Erro inesperado: {e}. Pulando.")

def main():
    try:
        fazer_login()
        caminho = Path("relatorios/restricoes_jornada.xlsx")
        df = pd.read_excel(caminho)
        df = df.dropna(subset=["CR", "IMEI"]).drop_duplicates(subset=["CR", "IMEI"])

        for _, row in df.iterrows():
            imei = str(row["IMEI"]).strip()
            crs = str(row["CR"]).strip()

            if len(imei) < 5 or imei.count("0") > 6:
                print(f"IMEI inválido (curto ou muitos zeros): {imei}. Pulando...")
                continue
            
            processar_dispositivo(imei, crs)
    finally:
        driver.quit()
        print("Processo concluído. Navegador fechado.")

if __name__ == "__main__":
    main()
