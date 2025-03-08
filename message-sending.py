import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import time
import urllib.parse
from selenium.common.exceptions import NoSuchElementException, TimeoutException

# Leitura da base de contatos
contatos_df = pd.read_excel("./ListaContatos.xlsx")

# Inicialização do navegador
navegador = webdriver.Chrome()
navegador.get("https://web.whatsapp.com/")

# Aguardar o carregamento da página
WebDriverWait(navegador, 90).until(EC.presence_of_element_located((By.ID, "side")))

# Lista para armazenar resultados do envio
resultado_envio = []

# Obter a data atual para o nome do arquivo de log
hoje = datetime.now()

# Início do script de envio
for index, row in contatos_df.iterrows():
    pessoa = row["Nome"]
    fone = str(row["Telefone"])  # Certifique-se de que o telefone seja uma string
    mensagem = f"Olá, {pessoa}! 🌟\n\nSeja bem-vindo(a)."

    link = f"https://web.whatsapp.com/send?phone=55{fone}&text={urllib.parse.quote(mensagem)}"

    navegador.get(link)

    try:
        # Aguardar o campo de mensagem
        WebDriverWait(navegador, 60).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="main"]/footer/div[1]/div/span/div/div[2]/div[1]/div[2]/div/p/span'))
        )

        # Enviar a mensagem
        navegador.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span/div/div[2]/div[1]/div[2]/div/p/span').send_keys(Keys.ENTER)
        time.sleep(7)

        resultado_envio.append({"Nome": pessoa, "Telefone": fone, "Status": "Enviado"})
    except (TimeoutException, NoSuchElementException) as e:
        print(f"Erro ao enviar mensagem para {pessoa} ({fone}): {e}")
        resultado_envio.append({"Nome": pessoa, "Telefone": fone, "Status": "Não Enviado"})

# Salvar resultados em arquivo Excel
resultado_df = pd.DataFrame(resultado_envio)
nome_arquivo = f"./log_{hoje.strftime('%Y-%m-%d')}.xlsx"
resultado_df.to_excel(nome_arquivo, sheet_name="Resultado", index=False)
print(f"Envio finalizado e resultados salvos no arquivo '{nome_arquivo}'.")

# Fechar o navegador
navegador.quit()