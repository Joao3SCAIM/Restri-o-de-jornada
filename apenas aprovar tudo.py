from imports import *

load_dotenv()

USER_PORTAL = os.getenv("USER_PORTAL")
PASSWORD_PORTAL = os.getenv("PASSWORD_PORTAL")

edge_driver_path = r"C:\Users\joao.escaim\Downloads\edgedriver_win64 (2)\msedgedriver.exe"
driver_service = Service(edge_driver_path)
options = webdriver.EdgeOptions()
options.add_argument("--start-maximized")

driver = webdriver.Edge(service=driver_service, options=options)
wait = WebDriverWait(driver, 30)

def realizar_login():
    print("Realizando login...")
    driver.get("https://portal.gpssa.com.br/GPS/Login.aspx")
    time.sleep(3)
    driver.find_element(By.ID, "txtUsername-inputEl").send_keys(USER_PORTAL)
    driver.find_element(By.ID, "txtPassword-inputEl").send_keys(PASSWORD_PORTAL)
    time.sleep(1)
    driver.find_element(By.ID, "btnLogin-btnInnerEl").click()
    time.sleep(3)
    print("Login concluído.")

def navegar_para_apontamento():
    print("Navegando para a página de apontamento...")
    driver.get("https://portal.gpssa.com.br/SAD/ponto?IDPAGINA=445&GRUPOABA=8.1%20-%20Sistema%20de%20Apontamento")
    wait.until(EC.presence_of_element_located((By.ID, "ComboBoxCampo-inputEl")))
    print("Página de apontamento carregada.")

def configurar_busca():
    print("Configurando busca...")
    campo = wait.until(EC.element_to_be_clickable((By.ID, "ComboBoxCampo-inputEl")))
    campo.clear()
    campo.send_keys("DIRETOR(A) EXECUTIVO")
    time.sleep(1.5)

    query = driver.find_element(By.ID, "ComboBoxQuery-inputEl")
    query.clear()
    query.send_keys("BRIAN SILVA")
    time.sleep(2)
    query.send_keys(Keys.RETURN)
    time.sleep(2)

    driver.find_element(By.ID, "ButtonBuscar-btnInnerEl").click()
    print("Busca iniciada...")

    # Espera carregar dados
    WebDriverWait(driver, 800).until(
        EC.text_to_be_present_in_element(
            (By.XPATH, '//div[contains(@class, "x-grid-cell-inner") and contains(@style, "text-align:center;color:green;font-weight: bold;")]'),
            "GPSa"
        )
    )

    driver.find_element(By.ID, "gridcolumn-1044-textEl").click()
    time.sleep(10)
    driver.find_element(By.ID, "gridcolumn-1044-textEl").click()
    time.sleep(10)

def aprovar_todas_as_restricoes():
    print("Iniciando aprovação de todas as linhas...")
    pagina_atual = 1

    while True:
        print(f"Processando página {pagina_atual}...")

        linhas = wait.until(
            EC.presence_of_all_elements_located((By.XPATH, "//table[contains(@class, 'x-grid-item')]"))
        )

        for i, linha in enumerate(linhas, start=1):
            try:
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", linha)
                time.sleep(0.5)

                btn_detalhes = linha.find_element(By.XPATH, ".//span[contains(@class, 'icon-applicationviewdetail')]")
                driver.execute_script("arguments[0].click();", btn_detalhes)
                print(f"Linha {i}: Detalhes abertos")

                try:
                    wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'x-window')]")))
                    time.sleep(3)

                    # Clicar em "Enviar não integrados"
                    btn_enviar = wait.until(
                        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Enviar não integrados')]"))
                    )
                    btn_enviar.click()
                    print("Clicado em 'Enviar não integrados'")

                    # Confirmar com "OK"
                    time.sleep(3)
                    btn_ok = wait.until(
                        EC.element_to_be_clickable((By.XPATH, "(//span[contains(text(), 'OK')])[1]"))
                    )
                    btn_ok.click()
                    print("Clicado em 'OK'")
                    time.sleep(1)

                except Exception as e:
                    print(f"Erro ao aprovar linha {i}: {str(e)}")

                finally:
                    try:
                        btn_fechar = wait.until(
                            EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'x-tool-close')]"))
                        )
                        btn_fechar.click()
                        time.sleep(1)
                    except:
                        pass

            except Exception as e:
                print(f"Erro ao processar linha {i}: {str(e)}")

        try:
            btn_proxima = driver.find_element(By.XPATH, "//span[contains(@class, 'x-tbar-page-next')]")
            btn_proxima.click()
            print("Indo para próxima página...")
            pagina_atual += 1
            wait.until(EC.staleness_of(linhas[0]))
            time.sleep(3)
        except:
            print("Não há mais páginas para processar.")
            break

def main():
    try:
        realizar_login()
        breakpoint()

        navegar_para_apontamento()
        breakpoint()

        configurar_busca()
        breakpoint()

        aprovar_todas_as_restricoes()
        breakpoint()

    except Exception as e:
        print(f"\nErro durante execução: {str(e)}")
    finally:
        driver.quit()
        print("Processo finalizado. Navegador fechado.")

if __name__ == "__main__":
    main()
