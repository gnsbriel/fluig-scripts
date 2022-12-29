
# ----- Imports ----- #

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from pathlib import Path
import sys
import csv
import time
import itertools
import configparser


def signin(browser, website, timeout):

    browser.get(website["host"])

    # Passo para efetuar login
    WebDriverWait(browser, timeout).until(EC.element_to_be_clickable(
        (By.ID, "username"))).send_keys(website["username"])
    WebDriverWait(browser, timeout).until(EC.element_to_be_clickable(
        (By.ID, "password"))).send_keys(website["password"])
    WebDriverWait(browser, timeout).until(EC.element_to_be_clickable(
        (By.ID, "submitLogin"))).click()

    # Passo necessário após efetuado login
    WebDriverWait(browser, timeout).until(EC.element_to_be_clickable(
        (By.CLASS_NAME, "close"))).click()


def write_to_csv(browser, website, timeout, start_time):
    with open("tickets.csv", "r", encoding="utf-8") as csv_file:

        csv_reader = csv.reader(csv_file, delimiter=";")
        next(csv_reader)  # pula primeria linha ao ler o arquivo tickets.csv

        for line in csv_reader:

            NUMSOLICITACAO = line[0]

            # Acessa a solicitação
            browser.get(f"{website['host']}/portal/p/1/pageworkflowview?app_ecm_workflowview_detailsProcessInstanceID={NUMSOLICITACAO}")

            # Entra no Iframe correto
            WebDriverWait(browser, timeout).until(EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR, "iframe[id='workflowView-cardViewer']")))

            # Captura os dados
            SOLICITANTE = WebDriverWait(browser, timeout).until(EC.presence_of_element_located((By.CSS_SELECTOR, "span[id='txt_dados_solicitante']"))).get_attribute('textContent')
            DATADAEMISSAO = WebDriverWait(browser, timeout).until(EC.presence_of_element_located((By.CSS_SELECTOR, "span[id='txt_dados_data']"))).get_attribute('textContent')
            RAZAOSOCIALCLIENTE = WebDriverWait(browser, timeout).until(EC.presence_of_element_located((By.CSS_SELECTOR, "span[id='txt_dados_razao_social']"))).get_attribute('textContent')
            CNPJCLIENTE = WebDriverWait(browser, timeout).until(EC.presence_of_element_located((By.CSS_SELECTOR, "span[id='sp_dados_cliente']"))).get_attribute('textContent')

            codigodosprodutos = WebDriverWait(browser, timeout).until(EC.presence_of_all_elements_located((By.XPATH, "//table[@id='tb_info_servicos']/tbody/tr/td/div/span[contains(@id, 'produto')][1]")))
            codigos = filter(lambda x: x.is_displayed(), codigodosprodutos)

            tabeladosprodutos = WebDriverWait(browser, timeout).until(EC.presence_of_all_elements_located((By.XPATH, "//table[@id='tb_info_servicos']/tbody/tr/td/div/span[contains(@id, 'produto')][2]")))
            produtos = filter(lambda x: x.is_displayed(), tabeladosprodutos)

            quantidades = WebDriverWait(browser, timeout).until(EC.presence_of_all_elements_located((By.XPATH, "//table[@id='tb_info_servicos']/tbody/tr/td/span[contains(@id, 'tb_num_quantidade___')]")))

            valoresunitarios = WebDriverWait(browser, timeout).until(EC.presence_of_all_elements_located((By.XPATH, "//table[@id='tb_info_servicos']/tbody/tr/td/div/span[contains(@id, 'tb_txt_val_unit___')]")))

            valoresbrutos = WebDriverWait(browser, timeout).until(EC.presence_of_all_elements_located((By.XPATH, "//table[@id='tb_info_servicos']/tbody/tr/td/div/span[contains(@id, 'tb_txt_val_bruto___')]")))

            centrosdecusto = WebDriverWait(browser, timeout).until(EC.presence_of_all_elements_located((By.XPATH, "//table[@id='tb_info_servicos']/tbody/tr/td/span[contains(@name, 'tb_ztxt_centro_custo___')]")))

            DESCRICAODANOTA = WebDriverWait(browser, timeout).until(EC.presence_of_element_located((By.CSS_SELECTOR, "span[id='txta_observacao_nota']"))).get_attribute('textContent')
            DESCRICAODOFATURAMENTO = WebDriverWait(browser, timeout).until(EC.presence_of_element_located((By.CSS_SELECTOR, "span[id='txt_observacao_fatura']"))).get_attribute('textContent')

            # Preenche o arquivo relatorio.csv
            with open('Documents/Relatorio.csv', 'a') as f:
                original_stdout = sys.stdout
                sys.stdout = f

                for (codigo, produto, qtde, valor, valorbruto, centrodecusto) in zip(codigos, produtos, quantidades, valoresunitarios, valoresbrutos, centrosdecusto):
                    a = valorbruto.text.replace(".", "")
                    b = a.replace(",", ".")
                    valortotal = b.replace(".", ",")
                    # NÚMERO FLUIGæSOLICITANTEæDATA DE EMISSÃOæRAZÃO SOCIAL DO CLIENTEæCNPJ DO CLIENTEæCÓDIGO DO PRODUTOæPRODUTOæQUANTIDADEæVALOR UNITÁRIOæVALOR TOTAL BRUTOæ CENTRO DE CUSTOæDESCRIÇÃO DA NOTAæDECRIÇÃO DO FATURAMENTO
                    print(f"{NUMSOLICITACAO}æ"
                          f"{SOLICITANTE}æ"
                          f"{DATADAEMISSAO}æ"
                          f"{RAZAOSOCIALCLIENTE}æ"
                          f"{CNPJCLIENTE}æ"
                          f"{codigo.text}æ"
                          f"{produto.text}æ"
                          f"{qtde.text}æ"
                          f"{valor.text}æ"
                          f"{valortotal}æ"
                          f"{centrodecusto.text}æ"
                          f"{DESCRICAODANOTA}æ"
                          f"{DESCRICAODOFATURAMENTO}"
                          )

                sys.stdout = original_stdout
    print("--- %s seconds ---" % (time.time() - start_time))


def main():
    start_time = time.time()

    cfg = configparser.ConfigParser()
    cfg.read("config.ini")

    browser = webdriver.Firefox(service_log_path='/dev/null')  # navegador utilizado
    timeout = 20  # em segundos

    destinationFolder = Path("Documents")
    destinationFolder.mkdir(parents=True, exist_ok=True)

    website = {
        "host": cfg.get("website", "host"),
        "username": cfg.get("website", "username"),
        "password": cfg.get("website", "password")
    }

    signin(browser, website, timeout)
    write_to_csv(browser, website, timeout, start_time)


if __name__ == "__main__":
    main()
