import os
import time
from time import sleep
from datetime import datetime, date, timedelta
import calendar
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from DCTF_WEB_TESTE import obter_vencimento_dctfweb_por_excel
from EFD_Contribuicao import obter_vencimento_efd_contribuicoes_por_excel
from EFD_Reinf import obter_vencimento_efd_reinf_por_excel
# from SolutionPacket.Solution_bank import server_bank


# =========================
# DOWNLOAD / SELENIUM
# =========================
def extracao_site(download_dir):
    options = Options()
    options.add_experimental_option("prefs", {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        #Mantido, mas agora o fluxo espera Excel. Se baixar PDF, vai avisar.
        "plugins.always_open_pdf_externally": True,
        "plugins.plugins_disabled": ["Chrome PDF Viewer"]
    })

    servico = Service(ChromeDriverManager().install())
    navegador = webdriver.Chrome(service=servico, options=options)
    navegador.maximize_window()
    navegador.implicitly_wait(15)

    try:
        sleep(1)
        navegador.get("https://www.gov.br/receitafederal/pt-br/assuntos/agenda-tributaria")
        navegador.refresh()
        sleep(1)

        try:
            aceitar_cookies_button = navegador.find_element(By.XPATH, "//*[text()='Aceitar cookies']")
            aceitar_cookies_button.click()
            sleep(1)
        except Exception:
            pass

        navegador.find_element(By.XPATH, "//a[contains(@href, 'anexo-ade-corat')]").click()
        sleep(7)  # depois você pode trocar por WebDriverWait
    finally:
        navegador.quit()


def aguardar_excel_baixado(pasta, timeout=60):
    """
    Espera terminar download e retorna o caminho do arquivo .xlsx mais recente.
    """
    inicio = time.time()

    while time.time() - inicio < timeout:
        arquivos = os.listdir(pasta)

        # download ainda em andamento
        if any(a.lower().endswith(".crdownload") for a in arquivos):
            time.sleep(1)
            continue

        xlsxs = [a for a in arquivos if a.lower().endswith(".xlsx")]
        if xlsxs:
            caminhos = [os.path.join(pasta, a) for a in xlsxs]
            return max(caminhos, key=os.path.getmtime)

        time.sleep(1)

    raise TimeoutError("Nenhum arquivo .xlsx foi baixado dentro do tempo esperado.")

# =========================
# BANCO - INSERT LEGADO (DIA) [mantido p/ chumbados]
# =========================
def inserir_dados_bd(cursor_, nome_, dia_venci, empresa):
    hoje = datetime.today()
    ano = hoje.year
    mes = hoje.month

    try:
        dia = int(dia_venci)
    except (TypeError, ValueError):
        print(f"Dia inválido '{dia_venci}' para {nome_}; pulando.")
        return

    ultimo_dia = calendar.monthrange(ano, mes)[1]
    if dia < 1 or dia > ultimo_dia:
        print(f"Dia inválido {dia} para {mes:02d}/{ano} em {nome_}; pulando.")
        return

    data_entrega = date(ano, mes, dia)
    data_consulta = hoje.date()
    quantidade_dias = 3
    data_alerta = data_entrega - timedelta(days=quantidade_dias)

    cursor_.execute(f"""
        INSERT INTO [dbo].[Agenda_Tributaria]
        (Nome, Data_entrega, Data_consulta, quantidade_dias, data_alerta, empresa)
        VALUES ('{nome_}', '{data_entrega}', '{data_consulta}', '{quantidade_dias}', '{data_alerta}', '{empresa}')
    """)
    cursor_.commit()

# =========================
# BANCO - INSERT NOVO (DATA COMPLETA)
# =========================
def inserir_dados_bd_data(cursor_, nome_, data_entrega, empresa, quantidade_dias=3):
    """
    Usa data completa (datetime.date) já calculada via Excel.
    """
    if not data_entrega:
        print(f"Sem data de entrega para {nome_}; pulando.")
        return

    data_consulta = datetime.today().date()
    data_alerta = data_entrega - timedelta(days=quantidade_dias)

    try:
        # evita duplicidade simples (nome + data + empresa)
        existente = cursor_.execute("""
            SELECT 1
            FROM [dbo].[Agenda_Tributaria]
            WHERE Nome = ? AND Data_entrega = ? AND empresa = ?
        """, (nome_, data_entrega, empresa)).fetchone()

        if existente:
            print(f"Já existe {nome_} em {data_entrega} para {empresa}; não inserido.")
            return

        cursor_.execute("""
            INSERT INTO [dbo].[Agenda_Tributaria]
            (Nome, Data_entrega, Data_consulta, quantidade_dias, data_alerta, empresa)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (nome_, data_entrega, data_consulta, quantidade_dias, data_alerta, empresa))
        cursor_.commit()

        print(f"Inserido {nome_} | entrega={data_entrega} | alerta={data_alerta}")

    except Exception as e:
        print(f"Falha ao inserir {nome_}: {e}")


# =========================
# NOVO FLUXO VIA EXCEL
# =========================
def processar_agenda_excel(cursor_, empresa, caminho_excel):
    print(f"Processando agenda via Excel: {caminho_excel}")

    # DCTFWeb
    try:
        venc_dctfweb = obter_vencimento_dctfweb_por_excel(caminho_excel)
        if venc_dctfweb:
            inserir_dados_bd_data(cursor_, "DCTFWeb", venc_dctfweb, empresa)
        else:
            print("DCTFWeb não encontrada no Excel.")
    except Exception as e:
        print(f"DCTFWeb via Excel: {e}")

    # EFD.Contribuições
    try:
        venc_efd_contrib = obter_vencimento_efd_contribuicoes_por_excel(caminho_excel)
        if venc_efd_contrib:
            inserir_dados_bd_data(cursor_, "EFD.Contribuições", venc_efd_contrib, empresa)
        else:
            print("EFD.Contribuições não encontrada no Excel.")
    except Exception as e:
        print(f"EFD.Contribuições via Excel: {e}")

    # EFD-Reinf
    try:
        venc_efd_reinf = obter_vencimento_efd_reinf_por_excel(caminho_excel)
        if venc_efd_reinf:
            inserir_dados_bd_data(cursor_, "EFD-Reinf", venc_efd_reinf, empresa)
        else:
            print("EFD-Reinf não encontrada no Excel.")
    except Exception as e:
        print(f"EFD-Reinf via Excel: {e}")


def limpar_arquivos_download(pasta_downloads, extensoes=(".xlsx", ".crdownload")):
    for arquivo in os.listdir(pasta_downloads):
        if arquivo.lower().endswith(tuple(ext.lower() for ext in extensoes)):
            try:
                os.remove(os.path.join(pasta_downloads, arquivo))
            except Exception as e:
                print(f"Não consegui remover {arquivo}: {e}")


def main(cursor_):
    download_dir = os.path.join(os.getcwd(), "agenda")
    os.makedirs(download_dir, exist_ok=True)

    # limpa restos de execuções anteriores pra não pegar arquivo velho
    limpar_arquivos_download(download_dir)

    # 1) Baixa agenda no site
    extracao_site(download_dir)

    # 2) Aguarda Excel baixado
    try:
        arquivo_excel = aguardar_excel_baixado(download_dir, timeout=90)
        print(f"Excel baixado: {arquivo_excel}")
    except Exception as e:
        print(f"Falha ao aguardar Excel baixado: {e}")
        return

    # 3) Processa via módulos importados
    processar_agenda_excel(cursor_, "Paranoa", arquivo_excel)

def teste_excel_sem_banco(caminho_excel):
    print(f"Lendo Excel: {caminho_excel}")

    try:
        venc_dctfweb = obter_vencimento_dctfweb_por_excel(caminho_excel)
        print("DCTFWeb:", venc_dctfweb)
    except Exception as e:
        print("DCTFWeb:", e)

    try:
        venc_efd_contrib = obter_vencimento_efd_contribuicoes_por_excel(caminho_excel)
        print("EFD.Contribuições:", venc_efd_contrib)
    except Exception as e:
        print("EFD.Contribuições:", e)

    try:
        venc_efd_reinf = obter_vencimento_efd_reinf_por_excel(caminho_excel)
        print("EFD-Reinf:", venc_efd_reinf)
    except Exception as e:
        print("[ERRO] EFD-Reinf:", e)


if __name__ == "__main__":
    download_dir = os.path.join(os.getcwd(), "agenda")
    os.makedirs(download_dir, exist_ok=True)

    limpar_arquivos_download(download_dir)
    extracao_site(download_dir)

    try:
        arquivo_excel = aguardar_excel_baixado(download_dir, timeout=90)
        print(f"[OK] Excel baixado: {arquivo_excel}")
        teste_excel_sem_banco(arquivo_excel)
    except Exception as e:
        print(f"[ERRO] Teste sem banco falhou: {e}")

    # banco = server_bank.Bank('Agenda tributaria')
    # cursor = banco.bank_connection('user_rpa', '%*4Us5z$', '186.193.228.29', 'rpa')
    # main(cursor)

    # chumbados (mantém por enquanto)
    empresas_dias = {
        'Dipam': '30',
        'FCI': '30',
        'GIA': '20',
        'GISS_CIVIL': '10',
        'GISS_PRESTADO': '15',
        'GISS_TOMADOS': '20',
        'Licenciamento': '30',
        'Nos_conformes': '28',
        'Sped': '20'
    }
    empre = 'Paranoa'

    # for nome, dia in empresas_dias.items():
    #     inserir_dados_bd(cursor, nome, dia, empre)
    # cursor.close()

    print("Arquivo da Agenda pronto para usar fluxo via Excel.")
  
