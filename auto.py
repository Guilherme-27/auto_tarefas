import pandas as pd
import unicodedata
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def remover_acentos(texto):
    return unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('ASCII')

# ABRIR CHROME COM:
# "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\Users\Guilherme\AppData\Local\Google\Chrome\User Data"
options = Options()
options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")

driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 20)

# === Leitura e preparação dos dados ===
df = pd.read_csv("Dados.csv", sep=";", skiprows=2, names=["Nota (%)", "Nome"])
df["Nota (%)"] = (
    df["Nota (%)"]
    .str.replace("%", "")
    .str.replace(",", ".")
    .astype(float)
    / 10  # transforma para escala de 0 a 10
)

# === Loop principal: busca + preenchimento ===
for index, row in df.iterrows():
    nome_original = row["Nome"]
    nome = remover_acentos(nome_original)
    nota = str(round(row["Nota (%)"], 1)).replace(".", ",")

    # Busca o aluno pelo nome
    campo_busca = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, "input.dt-input"))
    )
    campo_busca.clear()
    campo_busca.send_keys(nome)
    time.sleep(0.2)

    print(f"🔍 Buscando: {nome} → Nota: {nota}")

    try:
        # Espera o campo de nota aparecer e preenche
        campo_nota = WebDriverWait(driver, 4).until(
            EC.presence_of_element_located((By.NAME, "n.NotaAtribuida"))
        )
        campo_nota.clear()
        campo_nota.send_keys(nota)
        print("✅ Nota preenchida.")
    except Exception as e:
        print(f"⚠️ Erro ao tentar preencher a nota para {nome}: {e}")

    time.sleep(0.5)  # pausa mínima antes do próximo aluno