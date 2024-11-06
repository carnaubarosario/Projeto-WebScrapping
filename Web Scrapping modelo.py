from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import pandas as pd

# Caminho do arquivo Excel com os EANs dos produtos
excel_path = r'c:\Users\seu_usuario\Documents\Produtos_EANs.xlsx'

# Carregar o arquivo Excel, garantindo que os EANs sejam lidos como strings
df = pd.read_excel(excel_path, dtype=str)

# Inicializar o navegador (ChromeDriver)
driver = webdriver.Chrome()

# Função para buscar o EAN no site Economiza Alagoas
def buscar_ean_no_site(ean):
    driver.get("https://economizaalagoas.sefaz.al.gov.br/")
    time.sleep(2)  # Aguardar o carregamento da página

    # Encontrar o campo de busca de EAN e digitar o código
    search_box = driver.find_element(By.ID, 'textoConsulta')
    search_box.clear()
    search_box.send_keys(ean)
    search_box.send_keys(Keys.RETURN)

    time.sleep(3)  # Aguardar os resultados carregarem
    
    try:
        # Encontrar o item da lista e clicar para abrir os detalhes
        item_lista = driver.find_element(By.XPATH, "//li[@class='mdl-list__item mdl-list__item--two-line']")
        item_lista.click()
        time.sleep(3)  # Aguardar os detalhes carregarem

        # Capturar as informações de preços
        cartoes = driver.find_elements(By.CLASS_NAME, 'cartao')
        resultados = {}

        for cartao in cartoes:
            try:
                # Capturar o nome do supermercado
                nome_supermercado = cartao.find_element(By.CLASS_NAME, 'cartao_contribuinte_bloco_esquerdo').text.strip().split('\n')[0]
                
                # Capturar o preço do produto
                preco = cartao.find_element(By.XPATH, "//span[@class='valor_ultima_venda']").text.strip()
                
                # Capturar o nome do produto
                nome_produto = cartao.find_element(By.CLASS_NAME, 'cartao_titulo_texto').text.strip()

                resultados[nome_supermercado] = {
                    'preco': preco,
                    'produto': nome_produto
                }
            except Exception as e:
                print(f"Erro ao capturar dados do cartão: {str(e)}")

        return resultados

    except Exception as e:
        print(f"Erro ao buscar EAN {ean}: {str(e)}")
        return None

# Iterar sobre os produtos na planilha e buscar preços no site
for index, row in df.iterrows():
    ean = row['EAN']  # Supondo que a coluna EAN esteja presente na planilha
    
    # Verificar se o EAN está vazio ou NaN
    if pd.isna(ean) or ean.strip() == '':
        ean = '-'  # Substitui por hífen e pula para o próximo
        print(f"EAN vazio na linha {index}, preenchido com hífen.")
        continue  # Pular para o próximo EAN

    ean = ean.split('.')[0]  # Garantir que não tenha caracteres extras
    resultados = buscar_ean_no_site(ean)
    
    if resultados:
        for supermercado, dados in resultados.items():
            # Armazenar os resultados na planilha
            df.at[index, supermercado] = f"{dados['produto']}: {dados['preco']}"

        print(f"EAN {ean} processado. Resultados: {resultados if resultados else 'Nenhum resultado'}")
    
    if index % 10 == 0:
        # Salvar periodicamente os resultados para evitar perda de dados
        temp_path = r'c:\Users\seu_usuario\Desktop\Pesquisa_Parcial.xlsx'
        df.to_excel(temp_path, index=False)
        print(f"Planilha salva parcialmente após {index + 1} linhas processadas.")

# Salvar o arquivo final com os resultados completos
df.to_excel(r'c:\Users\seu_usuario\Desktop\Pesquisa_Final.xlsx', index=False)

driver.quit()

print("Script finalizado. Planilha salva com todos os resultados.")
