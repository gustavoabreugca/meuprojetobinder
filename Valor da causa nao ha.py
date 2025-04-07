from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from multiprocessing import Pool, Manager
import os
import time
import pandas as pd
import pyperclip
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

def processar_numero(numero):
    driver = None
    try:
        # Initialize Chrome for this process
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        driver.maximize_window()

        print(f"Acessando o TJSP para o processo: {numero}")
        driver.get("https://esaj.tjsp.jus.br/cpopg/open.do")
        time.sleep(0.5)

        pyperclip.copy(numero)
        campo_busca = driver.find_element(By.XPATH, "/html/body/div[2]/form/section/div[2]/div/div[1]/div[1]/span[1]/input[1]")
        campo_busca.clear()
        
        action_chains = ActionChains(driver)
        campo_busca.click()
        action_chains.key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
        
        time.sleep(0.5)
        campo_busca.send_keys(Keys.TAB)
        time.sleep(0.5)
             
        botao_consultar = driver.find_element(By.XPATH, "/html/body/div[2]/form/section/div[4]/div/input")
        botao_consultar.click()
        
        time.sleep(1.5)
        
        botao_mais = driver.find_element(By.XPATH, "/html/body/div[1]/div[3]/div/div[1]/a/span[1]")
        botao_mais.click()
        
        # Aguarda até que o campo valor seja atualizado (texto não vazio)
        elemento_valor = WebDriverWait(driver, 10).until(
            lambda d: d.find_element(By.ID, "valorAcaoProcesso") if d.find_element(By.ID, "valorAcaoProcesso").text.strip() != "" else False
        )
        valor_causa = elemento_valor.text.strip()
        if not valor_causa:
            valor_causa = "não encontrado"
        print(f"Processo {numero} - Valor da ação encontrado: {valor_causa}")
        
        if driver:
            driver.quit()
        return {'processo': numero, 'valor_causa': valor_causa}
        
    except Exception as e:
        print(f"Processo {numero} - Erro ao obter valor da ação: {e}")
        if driver:
            driver.quit()
        return {'processo': numero, 'valor_causa': "Não encontrado"}

def processar_lote(lote_processos):
    driver = None
    resultados = []
    encontrados = 0
    lote_id = os.getpid()
    
    try:
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        driver.maximize_window()
        wait = WebDriverWait(driver, 10)

        for numero in lote_processos:
            try:
                print(f"Acessando o TJSP para o processo: {numero}")
                driver.get("https://esaj.tjsp.jus.br/cpopg/open.do")
                time.sleep(0.5)

                pyperclip.copy(numero)
                campo_busca = driver.find_element(By.XPATH, "/html/body/div[2]/form/section/div[2]/div/div[1]/div[1]/span[1]/input[1]")
                campo_busca.clear()
                
                action_chains = ActionChains(driver)
                campo_busca.click()
                action_chains.key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
                
                time.sleep(0.5)
                campo_busca.send_keys(Keys.TAB)
                time.sleep(0.5)
                     
                botao_consultar = driver.find_element(By.XPATH, "/html/body/div[2]/form/section/div[4]/div/input")
                botao_consultar.click()
                
                time.sleep(1.5)
                
                botao_mais = driver.find_element(By.XPATH, "/html/body/div[1]/div[3]/div/div[1]/a/span[1]")
                botao_mais.click()
                
                # Aguarda até que o campo valor seja atualizado (texto não vazio)
                elemento_valor = WebDriverWait(driver, 10).until(
                    lambda d: d.find_element(By.ID, "valorAcaoProcesso") if d.find_element(By.ID, "valorAcaoProcesso").text.strip() != "" else False
                )
                valor_causa = elemento_valor.text.strip()
                if not valor_causa:
                    valor_causa = "não encontrado"
                print(f"Processo {numero} - Valor da ação encontrado: {valor_causa}")
                
                encontrados += 1
                resultados.append({'processo': numero, 'valor_causa': valor_causa})
                
                # Save every 100 successful finds
                if encontrados > 0 and encontrados % 100 == 0:
                    print(f"Salvando lote {lote_id} - {encontrados} resultados encontrados")
                    salvar_resultados(resultados, f'_parcial_lote_{lote_id}_encontrados_{encontrados}')
                
            except Exception as e:
                print(f"Processo {numero} - Erro ao obter valor da ação: {e}")
                resultados.append({'processo': numero, 'valor_causa': "Não encontrado"})
                continue

        # Save final batch results
        if resultados:
            print(f"Salvando resultados finais do lote {lote_id} - Total encontrados: {encontrados}")
            salvar_resultados(resultados, f'_final_lote_{lote_id}_encontrados_{encontrados}')

        return resultados
    
    finally:
        if driver:
            driver.quit()

def salvar_resultados(resultados, suffix=''):
    try:
        df = pd.DataFrame(resultados)
        filename = f'resultados{suffix}.xlsx'
        df.to_excel(filename, index=False)
        print(f"Arquivo Excel '{filename}' atualizado com sucesso!")
    except Exception as e:
        print(f"Erro ao salvar arquivo: {e}")

if __name__ == '__main__':
    with open('processos.txt', 'r', encoding='utf-8') as arquivo:
        processos = [linha.strip() for linha in arquivo if linha.strip()]
    
    print(f"Foram encontrados {len(processos)} processos.")
    
    # Processa todos os processos em um único navegador
    resultados_finais = processar_lote(processos)
    
    # Salva resultados finais
    salvar_resultados(resultados_finais, '_completo')
    print(f"Execução finalizada! Total de resultados: {len(resultados_finais)}")