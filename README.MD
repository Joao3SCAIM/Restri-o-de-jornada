##Automação de Restrição de Jornada
Este projeto automatiza a extração de dados de restrições de jornada no portal GPSA utilizando Selenium com Microsoft Edge, registra as informações em uma planilha Excel e executa um script adicional de recadastro automaticamente.

##Funcionalidades
Login automático no portal GPSA com credenciais carregadas via .env.

Navegação para a página de apontamento.

Configuração e execução de busca personalizada.

Processamento paginado da tabela de registros, extraindo informações detalhadas (IMEI, nome, matrícula, CR, data).

Registro dos dados extraídos em uma planilha Excel formatada.

Agendamento da execução do processo toda segunda-feira às 7h.

Execução automática do script recadastrar_GPSA.py após o processamento.

