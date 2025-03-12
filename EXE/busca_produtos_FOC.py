from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import json
import time

# Configuração do WebDriver
options = webdriver.ChromeOptions()
options.add_argument("--ignore-certificate-errors")
options.add_argument("--allow-running-insecure-content")
options.add_argument("--disable-web-security")
options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.0.0 Safari/537.36")

# NÃO USAR headless para depuração
# options.add_argument("--headless")  # REMOVA ou COMENTE essa linha

# Iniciar WebDriver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# URL da página de produtos
url = "view-source:https://admin.fomnichannel.com.br/Gerenciador/MercadoLivre/Categoria/251"  # Substitua pela URL real
driver.get(url)

# Espera para o carregamento da página
wait = WebDriverWait(driver, 30)

# Forçar scroll para carregar produtos dinâmicos
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
time.sleep(5)  # Espera 5 segundos para garantir o carregamento

# Buscar produtos
try:
    produtos = driver.find_elements(By.CSS_SELECTOR, "label.nome-produto")
    if produtos:
        dados = [produto.text for produto in produtos]

        # Exibe os produtos encontrados
        print("\nProdutos encontrados:")
        for nome in dados:
            print(f"- {nome}")

        # Salva os produtos em JSON
        with open("produtos.json", "w", encoding="utf-8") as f:
            json.dump(dados, f, ensure_ascii=False, indent=4)

        print("\nDados salvos em 'produtos.json'.")
    else:
        print("\n❌ Nenhum produto encontrado! Verifique o seletor ou a página.")

except Exception as e:
    print(f"\n❌ Erro ao buscar produtos: {e}")

finally:
    driver.quit()
