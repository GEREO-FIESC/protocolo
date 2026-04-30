import os
import logging
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, Page

# 1. Configuração de Logging Estruturado
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger("CorreiosAutomation")

load_dotenv()
LOGIN = "correiosfiesc"
PASSWORD = os.getenv("SENHA")

def login_correios(page: Page) -> bool:
    """Realiza o login no portal CAS."""
    url_login = "https://cas.correios.com.br/login?service=https%3A%2F%2Fsfe.correios.com.br%2F"
    try:
        logger.info(f"Navegando para {url_login}")
        page.goto(url_login, wait_until="domcontentloaded")

        logger.info("Inserindo credenciais de acesso.")
        page.get_by_role("textbox", name="Usuário").fill(LOGIN)
        page.get_by_role("textbox", name="Senha").fill(PASSWORD)
        page.get_by_role("button", name="ENTRAR").click()

        page.wait_for_load_state("networkidle")
        
        if "sfe.correios.com.br" in page.url:
            logger.info("Login efetuado com sucesso!")
            return True
        else:
            logger.error("Falha no login: Redirecionamento não identificado.")
            return False
    except Exception as e:
        logger.error(f"Erro crítico durante o login: {str(e)}")
        return False

def preencher_filtros_fatura(page: Page, caminho_destino: str) -> bool:
    """Executa o preenchimento dos filtros e captura o download via expect_download."""
    try:
        # 1. Acessar página de fatura
        logger.info("Acessando menu Fatura.")
        page.get_by_role("link", name="Fatura", exact=True).click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1000)

        # 2. Seleção de filtros
        logger.info("Selecionando filtros de consulta.")
        page.get_by_label("Contrato Todos 9912336464 -").select_option("9912336464  ")
        page.wait_for_timeout(500)
        
        page.get_by_label("Centro de Custos Todos 393510").select_option("393510")
        page.wait_for_timeout(500)
        
        page.get_by_label("Status da Fatura Todos Acordo").select_option("A")
        page.wait_for_timeout(500)
        
        # 3. Execução da pesquisa inicial
        page.get_by_role("button", name="Pesquisar").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1000)

        # 4. Configuração para download
        page.get_by_role("link", name="Gerar Extrato").click()
        page.wait_for_timeout(500)
        
        page.get_by_role("radio", name="Opção de Extrato *: Extrato").check()
        page.get_by_label("Formato do Arquivo :").select_option("CSV")
        page.wait_for_timeout(500)

        # 5. Captura do download com timeout estendido
        logger.info("Aguardando download via PrimeFaces...")
        
        with page.expect_download(timeout=60000) as download_info:
            # Dispara o clique nativo via JS conforme sua preferência
            page.evaluate('document.getElementById("form-clientes:btnPesquisar").click()')
        
        download = download_info.value
        download.save_as(caminho_destino)
        
        logger.info(f"Download salvo com sucesso em: {caminho_destino}")
        return True

    except Exception as e:
        logger.error(f"Erro ao capturar download: {str(e)}")
        page.screenshot(path="erro_captura_download.png")
        return False
    
# --- Execução Principal ---
if __name__ == "__main__":
    # Caminho final definido
    caminho_final = r"C:\Users\felipe-ferreira\Music\extrato_sintetico\Extrato_Sintetico.csv"
    
    # Garante que a pasta de destino exista
    pasta_destino = os.path.dirname(caminho_final)
    if not os.path.exists(pasta_destino):
        os.makedirs(pasta_destino)
        
    with sync_playwright() as p:
        user_data_dir = os.path.join(os.getcwd(), "user_data_profile")
        
        context = p.chromium.launch_persistent_context(
            user_data_dir=user_data_dir,
            headless=False,
            accept_downloads=True,
            args=[
                "--start-maximized",
                "--disable-blink-features=AutomationControlled",
                "--disable-infobars"
            ],
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36"
        )
        
        context.add_init_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        
        page = context.pages[0]

        if login_correios(page):
            # Passando o caminho_final para a função de download
            if preencher_filtros_fatura(page, caminho_final):
                logger.info("Automação concluída com sucesso.")
            else:
                logger.error("Falha na etapa de filtros.")
        else:
            logger.error("Falha na etapa de login.")

        page.wait_for_timeout(5000) 
        context.close()