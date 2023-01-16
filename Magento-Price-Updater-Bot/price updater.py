from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
import pandas as pd
from time import sleep
from datetime import datetime


link = "https://(SiteLinkHere)/index.php/admin/catalog_product"
login = "login"
password = "password"

class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'

def atualizar_valor():
    
    #seleciona o produto correto
    ActionChains(driver).move_to_element(teste_SKU).click(teste_SKU).perform()
    sleep(2)
    
    #clica na aba de preço
    botao_prices = driver.find_element(By.XPATH, '//*[@id="product_info_tabs_group_8"]/span')
    ActionChains(driver).move_to_element(botao_prices).click(botao_prices).perform()
    sleep(2)


    #verifica se o valor novo já esta aplicado
    valoratual = driver.find_element(By.XPATH, '//*[@id="price"]').get_attribute('Value')


    if float(valoratual) != valor_novo:

        #verifica se o novo valor é igual ao de uma promo ja ativa
        valorpromoatual = driver.find_element(By.XPATH, '//*[@id="special_price"]').get_attribute('Value')
    
        if valorpromoatual != "":

            if float(valorpromoatual) >= valor_novo :

                limite_promo = driver.find_element(By.XPATH, '//*[@id="special_to_date"]').get_attribute('Value')

                data = datetime.now()
                data_hoje = data.strftime("%d/%m/%Y")

                #verifica se a data limite é menor que a data de hoje
                if limite_promo < data_hoje:
                    
                    #atualiza valor principal
                    preço = driver.find_element(By.XPATH, '//*[@id="price"]')
                    preço.send_keys(Keys.CONTROL+"A")
                    preço.send_keys(valor_novo)

                    #limpa promo antiga
                    promo = driver.find_element(By.XPATH, '//*[@id="special_price"]')
                    promo.send_keys(Keys.CONTROL+"A")
                    promo.send_keys(Keys.BACKSPACE)

                    fromdate = driver.find_element(By.XPATH, '//*[@id="special_from_date"]')
                    fromdate.send_keys(Keys.CONTROL+"A")
                    fromdate.send_keys(Keys.BACKSPACE)
                    
                    prazo = driver.find_element(By.XPATH, '//*[@id="special_to_date"]')
                    prazo.send_keys(Keys.CONTROL+"A")
                    prazo.send_keys(Keys.BACKSPACE)
                
                if limite_promo >= data_hoje:

                    valornovo10 = float(valorpromoatual) * 1.1  

                    preço = driver.find_element(By.XPATH, '//*[@id="price"]')
                    preço.send_keys(Keys.CONTROL+"A")
                    preço.send_keys(valornovo10) 

            else:
                #atualiza valor principal
                preço = driver.find_element(By.XPATH, '//*[@id="price"]')
                preço.send_keys(Keys.CONTROL+"A")
                preço.send_keys(valor_novo)
                
        else: 

            #atualiza valor principal
            preço = driver.find_element(By.XPATH, '//*[@id="price"]')
            preço.send_keys(Keys.CONTROL+"A")
            preço.send_keys(valor_novo)

    else: 
        pass


    #valor promo
    if valor_promo != float('0'):

        if valor_promo >= valor_novo:

            valor_novo10 = valor_promo * 1.1  

            preço = driver.find_element(By.XPATH, '//*[@id="price"]')
            preço.send_keys(Keys.CONTROL+"A")
            preço.send_keys(valor_novo10) 

        promo = driver.find_element(By.XPATH, '//*[@id="special_price"]')
        promo.send_keys(Keys.CONTROL+"A")
        promo.send_keys(valor_promo)

        data = datetime.now()
        data_hoje = data.strftime("%d/%m/%Y")

        hoje = driver.find_element(By.XPATH, '//*[@id="special_from_date"]')
        hoje.send_keys(Keys.CONTROL+"A")
        hoje.send_keys(data_hoje)
        
        prazo = driver.find_element(By.XPATH, '//*[@id="special_to_date"]')
        prazo.send_keys(Keys.CONTROL+"A")
        prazo.send_keys(data_promo)
    

    #salva as alterações
    salvar = driver.find_element(By.XPATH, '//button[5]/span/span/span')
    salvar.click()

    sleep(2)
    

planilha = '1.xlsx'
dados = pd.read_excel(planilha)
dados[['Prç Venda']] = dados[['Prç Venda']].fillna(0)

produtos_sem_cadastro = []
produtos_desativados = []
produtos_desativados_final = []


print(bcolors.OKGREEN + "\nIniciando atualização de valores." + bcolors.ENDC)

options = webdriver.ChromeOptions()
options.headless = True
# driver = webdriver.Chrome(options=options)
driver = webdriver.Chrome()
driver.set_window_size(850, 550)
driver.get(link)

sleep(2)

driver.find_element(By.XPATH, '//*[@id="username"]').send_keys(login)
driver.find_element(By.XPATH, '//*[@id="login"]').send_keys(password)
driver.find_element(By.XPATH, '//*[@id="loginForm"]/div/div[5]/input').click()
sleep(2)


status = Select(driver.find_element(By.XPATH, '//*[@id="productGrid_product_filter_status"]'))
status.select_by_value('1')


for index, row in dados.iterrows():

    try:

        sku = str(row['Código'])
        valor_novo = float(row['Novo'])
        valor_promo = float(row['Prç Venda'])
        data_promo = str(row['Validade'])

        #pesquisa SKU
        SKU_search = driver.find_element(By.XPATH, '//*[@id="productGrid_product_filter_sku"]')
        SKU_search.send_keys(Keys.CONTROL+"A")
        SKU_search.send_keys(sku)
        SKU_search.send_keys(Keys.ENTER)
        
        linha = int(index)+2


        sleep(2)

        item = 1
        pag = 1

        numero_pag = int(driver.find_element(By.XPATH, f'//*[@id="productGrid"]/table/tbody/tr/td[1]/input').get_attribute('Value'))

        #verifica se tem mais de uma pag de produtos
        if numero_pag != 1:
            
            prox_pag = driver.find_element(By.XPATH, f'//*[@id="productGrid"]/table/tbody/tr/td[1]/input')
            prox_pag.send_keys(Keys.CONTROL + 'A')
            prox_pag.send_keys(pag)
            prox_pag.send_keys(Keys.ENTER)

            sleep(0.3)

        
        teste_SKU = driver.find_element(By.XPATH, f'//*[@id="productGrid_table"]/tbody/tr[{item}]/td[6]')
        teste_SKU_numero = str.strip(driver.find_element(By.XPATH, f'//*[@id="productGrid_table"]/tbody/tr[{item}]/td[6]').get_attribute('innerHTML'))

        quantidade_produtos = int(driver.find_element(By.XPATH, f'//*[@id="productGrid-total-count"]').get_attribute('innerHTML'))

        #verifica se tem mais de 20 produtos
        if quantidade_produtos > 20:
            
            #verifica os SKU's até bater com o certo
            while teste_SKU_numero.casefold() != str.strip(sku).casefold():

                teste_SKU = driver.find_element(By.XPATH, f'//*[@id="productGrid_table"]/tbody/tr[{item}]/td[6]')
                teste_SKU_numero = str.strip(driver.find_element(By.XPATH, f'//*[@id="productGrid_table"]/tbody/tr[{item}]/td[6]').get_attribute('innerHTML'))

                item += 1
                
                #muda de pagina quando verifica o ultimo item da pagina
                if item > 20:
                    
                    pag += 1
                    
                    prox_pag = driver.find_element(By.XPATH, f'//*[@id="productGrid"]/table/tbody/tr/td[1]/input')
                    
                    prox_pag.send_keys(Keys.CONTROL + 'A')
                    prox_pag.send_keys(pag)
                    prox_pag.send_keys(Keys.ENTER)
                    
                    sleep(1)
                    
                    item = 1
            
            #quando achar o SKU:
            if teste_SKU_numero.casefold() == str.strip(sku).casefold():
                
                sleep(0.3)
                atualizar_valor()
                print(f'SKU: {str.strip(sku)}, atualizado. Linha {linha}')
                
                sleep(0.3)
        
        #verifica se tem igual ou menos que 20 produtos
        if quantidade_produtos <= 20:
            
            #verifica os SKU's até bater com o certo
            while teste_SKU_numero.casefold() != str.strip(sku).casefold():
                
                teste_SKU = driver.find_element(By.XPATH, f'//*[@id="productGrid_table"]/tbody/tr[{item}]/td[6]')
                teste_SKU_numero = str.strip(driver.find_element(By.XPATH, f'//*[@id="productGrid_table"]/tbody/tr[{item}]/td[6]').get_attribute('innerHTML'))
                
                item += 1
            
            #quando achar o SKU:
            if teste_SKU_numero.casefold() == str.strip(sku).casefold():
                
                sleep(0.3)
                atualizar_valor()
                print(f'SKU: {str.strip(sku)}, atualizado. Linha {linha}')
                
                sleep(0.3)
    
    
    except:
        produtos_desativados.append(sku)
        pass

print(bcolors.OKGREEN + "\nAtualização de valores finalizada com sucesso." + bcolors.ENDC)


#testa produtos desativados
if produtos_desativados != []:
    
    print(bcolors.WARNING + "\nVerificando produtos desativados e descadastrados.\n" + bcolors.ENDC)
    
    status = Select(driver.find_element(By.XPATH, '//*[@id="productGrid_product_filter_status"]'))
    status.select_by_value('2')
    
    for skudesativado in produtos_desativados:
        

        try:

            #pesquisa SKU
            SKU_search = driver.find_element(By.XPATH, '//*[@id="productGrid_product_filter_sku"]')
            SKU_search.send_keys(Keys.CONTROL+"A")
            SKU_search.send_keys(skudesativado)
            SKU_search.send_keys(Keys.ENTER)
            
            
            sleep(2)
            
            item = 1
            pag = 1
            
            numero_pag = int(driver.find_element(By.XPATH, f'//*[@id="productGrid"]/table/tbody/tr/td[1]/input').get_attribute('Value'))
            
            if numero_pag != 1:
                
                prox_pag = driver.find_element(By.XPATH, f'//*[@id="productGrid"]/table/tbody/tr/td[1]/input')
                prox_pag.send_keys(Keys.CONTROL + 'A')
                prox_pag.send_keys(pag)
                prox_pag.send_keys(Keys.ENTER)
                
                sleep(0.3)
                
            
            teste_SKU = driver.find_element(By.XPATH, f'//*[@id="productGrid_table"]/tbody/tr[{item}]/td[6]')
            teste_SKU_numero = str.strip(driver.find_element(By.XPATH, f'//*[@id="productGrid_table"]/tbody/tr[{item}]/td[6]').get_attribute('innerHTML'))
            
            quantidade_produtos = int(driver.find_element(By.XPATH, f'//*[@id="productGrid-total-count"]').get_attribute('innerHTML'))
            
            #verifica se tem mais de 20 produtos
            if quantidade_produtos > 20:
                
                #verifica os SKU's até bater com o certo
                while teste_SKU_numero.casefold() != str.strip(skudesativado).casefold():
                    
                    teste_SKU = driver.find_element(By.XPATH, f'//*[@id="productGrid_table"]/tbody/tr[{item}]/td[6]')
                    teste_SKU_numero = str.strip(driver.find_element(By.XPATH, f'//*[@id="productGrid_table"]/tbody/tr[{item}]/td[6]').get_attribute('innerHTML'))
                    
                    item += 1
                    
                    #muda de pagina quando verifica o ultimo item da pagina
                    if item > 20:
                        
                        pag += 1
                        
                        prox_pag = driver.find_element(By.XPATH, f'//*[@id="productGrid"]/table/tbody/tr/td[1]/input')
                        
                        prox_pag.send_keys(Keys.CONTROL + 'A')
                        prox_pag.send_keys(pag)
                        prox_pag.send_keys(Keys.ENTER)
                        
                        sleep(0.3)
                        
                        item = 1
                    
                #quando achar o SKU:
                if teste_SKU_numero.casefold() == str.strip(skudesativado).casefold():
                    
                    produtos_desativados_final.append(skudesativado)
                    
                    sleep(0.3)
            
            #verifica se tem igual ou menos que 20 produtos
            if quantidade_produtos <= 20:
                
                #verifica os SKU's até bater com o certo
                while teste_SKU_numero.casefold() != str.strip(skudesativado).casefold():
                    
                    teste_SKU = driver.find_element(By.XPATH, f'//*[@id="productGrid_table"]/tbody/tr[{item}]/td[6]')
                    teste_SKU_numero = str.strip(driver.find_element(By.XPATH, f'//*[@id="productGrid_table"]/tbody/tr[{item}]/td[6]').get_attribute('innerHTML'))
                    
                    item += 1
                
                #quando achar o SKU:
                if teste_SKU_numero.casefold() == str.strip(skudesativado).casefold():
                    
                    produtos_desativados_final.append(skudesativado)
                    
                    sleep(0.3)
                    
        except:
            
            produtos_sem_cadastro.append(skudesativado)
            
            pass
    
    
if produtos_desativados_final != []:
    
    print(bcolors.FAIL + f"Produtos desativados: {', '.join(produtos_desativados_final)}.\n" + bcolors.ENDC)

if produtos_sem_cadastro != []:
    
    print(bcolors.FAIL + f"Produtos não cadastrados no site: {', '.join(produtos_sem_cadastro)}.\n" + bcolors.ENDC)


driver.quit()
