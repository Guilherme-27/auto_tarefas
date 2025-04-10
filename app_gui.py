import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import pandas as pd
from leitor_csv import (
    ler_csv_sed,
    converter_nota_para_escala_10,
    remover_duplicatas_maior_nota,
    salvar_csv_limpo
)
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import unicodedata
import time

def remover_acentos(texto):
    return unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('ASCII')

def iniciar_chrome_com_perfil():
    options = Options()
    options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    return webdriver.Chrome(options=options)

def iniciar_lancamento(driver, df, log_callback):
    # Seleciona aba correta
    for handle in driver.window_handles:
        driver.switch_to.window(handle)
        if "Lançamento das Avaliações" in driver.title:
            log_callback(f"✅ Aba correta selecionada: {driver.title}")
            break
    else:
        log_callback("❌ Aba 'Lançamento das Avaliações' não encontrada.")
        messagebox.showinfo("Aviso", "Acesse a aba manualmente e clique em OK para continuar.")

    # Loop de preenchimento
    for index, row in df.iterrows():
        try:
            nome_original = row["Nome do Aluno"]
            nome = remover_acentos(nome_original)
            nota = str(round(row["Nota (%)"], 1)).replace(".", ",")

            campo_busca = WebDriverWait(driver, 5).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, "input.dt-input"))
            )
            campo_busca.clear()
            campo_busca.send_keys(nome)
            # Aguarda até que apenas uma linha de aluno esteja visível na tabela (ajuste se necessário)
            WebDriverWait(driver, 3).until(
                lambda d: len(d.find_elements(By.NAME, "n.NotaAtribuida")) >= 1
            )
            log_callback(f"🔍 Buscando: {nome} → Nota: {nota}")

            campo_nota = WebDriverWait(driver, 2).until(
                EC.presence_of_element_located((By.NAME, "n.NotaAtribuida"))
            )
            campo_nota.clear()
            campo_nota.send_keys(nota)
            log_callback("✅ Nota preenchida.")
        except Exception:
            log_callback(f"⚠️ Aluno '{nome_original}' não encontrado ou nota não lançada.")
        time.sleep(0.5)

    # Final: limpa campo de busca, ativa foco e seleciona 100 resultados por página
    try:
        campo_busca = driver.find_element(By.CSS_SELECTOR, "input.dt-input")
        campo_busca.clear()
        campo_busca.click()  # força foco para garantir que a tabela recarregue
        log_callback("🧹 Campo de busca limpo e ativado.")

        seletor = WebDriverWait(driver, 3).until(
            EC.element_to_be_clickable((By.NAME, "tableDiarioClasse_length"))
        )
        for option in seletor.find_elements(By.TAG_NAME, 'option'):
            if option.text.strip() == '100':
                option.click()
                log_callback("📄 Visualização ajustada para 100 alunos por página.")
                break
    except Exception:
        log_callback("⚠️ Não foi possível ajustar visualização final.")

# === INTERFACE TKINTER ===
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Lançador de Notas SED")

        self.caminho_csv = ""

        self.btn_csv = tk.Button(root, text="📂 Selecionar CSV", command=self.selecionar_csv)
        self.btn_csv.pack(pady=5)

        self.btn_iniciar = tk.Button(root, text="▶️ Iniciar Lançamento", command=self.iniciar, state=tk.DISABLED)
        self.btn_iniciar.pack(pady=5)

        self.txt_log = tk.Text(root, height=20, width=60)
        self.txt_log.pack(padx=10, pady=10)

    def log(self, mensagem):
        self.txt_log.insert(tk.END, mensagem + "\n")
        self.txt_log.see(tk.END)
        self.root.update()

    def selecionar_csv(self):
        caminho = filedialog.askopenfilename(filetypes=[("Arquivos CSV", "*.csv")])
        if caminho:
            self.caminho_csv = caminho
            self.log(f"📄 CSV selecionado: {caminho}")
            self.btn_iniciar.config(state=tk.NORMAL)

    def iniciar(self):
        threading.Thread(target=self.executar_processo).start()

    def executar_processo(self):
        try:
            self.log("🚀 Iniciando limpeza do CSV...")
            df = ler_csv_sed(self.caminho_csv)
            df = converter_nota_para_escala_10(df)
            df = remover_duplicatas_maior_nota(df)
            salvar_csv_limpo(df)
            self.log("✅ CSV limpo gerado.")

            df_limpo = pd.read_csv("dados_limpos.csv", sep=";")
            self.log("🌐 Conectando ao navegador Chrome...")
            driver = iniciar_chrome_com_perfil()
            self.log("🧠 Iniciando automação...")
            iniciar_lancamento(driver, df_limpo, self.log)
            self.log("✅ Fim do processo.")

        except Exception as e:
            self.log(f"❌ Erro geral: {e}")
            messagebox.showerror("Erro", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()