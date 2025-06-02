from imports import *

# Carrega as variáveis do arquivo .env
load_dotenv()
USER_PORTAL = os.getenv("USER_PORTAL")
PASSWORD_PORTAL = os.getenv("PASSWORD_PORTAL")
# Configurações iniciais
edge_driver_path = r"C:\Users\joao.escaim\Downloads\edgedriver_win64\msedgedriver.exe"
driver_service = Service(edge_driver_path)
options = webdriver.EdgeOptions()
options.add_argument("--start-maximized")

# Inicializar driver
driver = webdriver.Edge(service=driver_service, options=options)
wait = WebDriverWait(driver, 30)

# Criar planilha Excel
wb = Workbook()
ws = wb.active
ws.title = "Restrições de Jornada"
headers = ["CR", "IMEI", "IMEI 2", "Nome", "Matrícula", "Data", "Processado em"] 
ws.append(headers)

# Formatar cabeçalho
for cell in ws[1]:
    cell.font = Font(bold=True)

def realizar_login():
    # Função para realizar o login
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
    
    # Esperar resultados carregarem
    elemento_gpsa = WebDriverWait(driver, 800).until(
        EC.text_to_be_present_in_element(
            (By.XPATH, '//div[contains(@class, "x-grid-cell-inner") and contains(@style, "text-align:center;color:green;font-weight: bold;")]'), 
            "GPSa"
        )
    )
    # clicar no botão integração e observação 
    driver.find_element(By.ID, "gridcolumn-1044-textEl").click()
    time.sleep(10)
    driver.find_element(By.ID, "gridcolumn-1044-textEl").click() 
    time.sleep(10)

def processar_tabela():
    print("\nIniciando processamento da tabela...")
    pagina_atual = 1
    tabelas_sem_restricao = 0
    
    while True:
        print(f"\nProcessando página {pagina_atual}...")
        
        # Localizar todas as linhas da tabela
        linhas = wait.until(
            EC.presence_of_all_elements_located((By.XPATH, "//table[contains(@class, 'x-grid-item')]"))
        )
        print(f"Encontradas {len(linhas)} linhas nesta página")
        
        restricoes_encontradas = False
        
        for i, linha in enumerate(linhas, start=1):
            try:
                # Scroll para a linha
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", linha)
                time.sleep(0.5)
                
                # Verificar restrição de jornada
                celula_restricao = linha.find_element(By.XPATH, ".//td[contains(@data-columnid, 'gridcolumn-1044')]//div")
                
                if "RESTRIÇÃO DE JORNADA" in celula_restricao.text:
                    restricoes_encontradas = True
                    print(f"\nLinha {i}: Processando restrição...")
                    
                    # Extrair dados da linha
                    nome = linha.find_element(By.XPATH, ".//td[contains(@data-columnid, 'gridcolumn-1023')]//div").text
                    matricula = linha.find_element(By.XPATH, ".//td[contains(@data-columnid, 'gridcolumn-1029')]//div").text
                    data = linha.find_element(By.XPATH, ".//td[contains(@class, 'x-grid-cell-datecolumn')]//div").text
                    cr = linha.find_element(By.XPATH, ".//td[contains(@data-columnid, 'gridcolumn-1032')]//div").text.split(" - ")[0]
                    
                    print(f"Dados: CR={cr}, Nome={nome}, Matrícula={matricula}, Data={data}")
                    
                    # Abrir detalhes
                    btn_detalhes = linha.find_element(By.XPATH, ".//span[contains(@class, 'icon-applicationviewdetail')]")
                    driver.execute_script("arguments[0].click();", btn_detalhes)
                    print("Janela de detalhes aberta")
                    
                    # Processar janela de detalhes
                    try:
                        # Esperar janela abrir
                        wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'x-window')]")))
                        time.sleep(2)
                        
                        # Função para extrair todos os IMEIs (15 dígitos) ou retornar ["1"] 
                        def extrair_imeis(driver):
                            
                            celulas = driver.find_elements(
                                By.XPATH,
                                "//td[contains(@class, 'x-grid-cell')]"
                                "/div[contains(@class, 'x-grid-cell-inner') "
                                "and translate(text(), '0123456789', '') = '']"
                            )
                            
                            imeis = []
                            # Busca todos os números com 15 dígitos (IMEI)
                            for celula in celulas:
                                texto = celula.text.strip()
                                if len(texto) == 15 and texto.isdigit():
                                    imeis.append(texto)
                            
                            return imeis if imeis else ["1"]  # Padrão se não encontrar IMEI

                        imeis = extrair_imeis(driver)
                        imei1 = imeis[0] if len(imeis) > 0 else "1"
                        imei2 = imeis[1] if len(imeis) > 1 else ""  # Segundo IMEI ou vazio
                        
                        print(f"IMEIs encontrados: {imei1}, {imei2}")

                        # Adicionar à planilha
                        data_processamento = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                        ws.append([cr, imei1, imei2, nome, matricula, data, data_processamento])  # Adicionado imei2
                        print("Dados registrados na planilha")
                        
                        # Só clica em "Enviar não integrados" se pelo menos um IMEI tiver 15 dígitos
                        if any(len(imei) == 15 for imei in imeis):
                            btn_enviar = wait.until(
                                EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Enviar não integrados')]"))
                            )
                            btn_enviar.click()
                            print("Clicado em 'Enviar não integrados'")
                            
                            # Confirmar envio
                            btn_ok = wait.until(
                                EC.element_to_be_clickable((By.XPATH, "(//span[contains(text(), 'OK')])[1]"))
                            )
                            btn_ok.click()
                            time.sleep(2)
                        else:
                            print("Nenhum IMEI com 15 dígitos encontrado - pulando envio")
                        
                    except Exception as e:
                        print(f"Erro ao processar detalhes: {str(e)}")
                    finally:
                        # Fechar janela de detalhes
                        try:
                            btn_fechar = wait.until(
                                EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'x-tool-close')]"))
                            )
                            btn_fechar.click()
                            time.sleep(1)
                        except:
                            pass
                
                else:
                    print(f"Linha {i}: Sem restrição de jornada")
            
            except Exception as e:
                print(f"Erro ao processar linha {i}: {str(e)}")
                continue
        
        # Contar tabelas sem restrição ou finalizar após 3
        if not restricoes_encontradas:
            tabelas_sem_restricao += 1
            print(f"Nenhuma restrição encontrada nesta página ({tabelas_sem_restricao}/3)")
            
            if tabelas_sem_restricao >= 3: 
                print("3 páginas sem restrições encontradas - finalizando processo")
                break
        else:
            tabelas_sem_restricao = 0  # Resetar contador se encontrar restrições
        
        # Verificar próxima página
        try:
            btn_proxima = driver.find_element(By.XPATH, "//span[contains(@class, 'x-tbar-page-next')]")
            btn_proxima.click()
            print("\nIndo para próxima página...")
            pagina_atual += 1
            
            # Esperar nova página carregar
            wait.until(EC.staleness_of(linhas[0]))
            time.sleep(3)
        except:
            print("\nNão há mais páginas para processar")
            break

def salvar_planilha():
    # Criar diretório se não existir
    if not os.path.exists("relatorios"):
        os.makedirs("relatorios")
    
    # Nome do arquivo com timestamp
    nome_arquivo = "relatorios/restricoes_jornada.xlsx"
    wb.save(nome_arquivo)
    print(f"\nPlanilha salva em: {os.path.abspath(nome_arquivo)}")

def main():
    try:
        realizar_login()
        navegar_para_apontamento()
        configurar_busca()
        processar_tabela()
        salvar_planilha()
    except Exception as e:
        print(f"\nErro durante execução: {str(e)}")
    finally:
        driver.quit()
        print("Processo concluído. Navegador fechado.")
        print("\nIniciando execução de recadastrar_GPSA.py...")
        subprocess.run(["python", "recadastrar_GPSA.py"])
        print("recadastrar_GPSA.py finalizado.")

def executar_tudo():
    main()

schedule.every().monday.at("07:00").do(executar_tudo)
print("Agendador iniciado. Aguardando horário programado...")

while True:
    schedule.run_pending()
    time.sleep(1)
