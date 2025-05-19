import tkinter as tk
from tkinter import messagebox
import pandas as pd
import os
import requests
import csv
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By

# === CONFIGURAÇÕES ===
ARQUIVO_PLANILHA = "CVM_Links.xlsx"
ARQUIVO_LOG = "resultado_downloads.csv"
PASTA_FORMULARIOS = "formularios"

# === PREPARO ===
os.makedirs(PASTA_FORMULARIOS, exist_ok=True)
df = pd.read_excel(ARQUIVO_PLANILHA, header=None)
names_links = df[0].tolist()
companies = [(names_links[i], names_links[i+1]) for i in range(0, len(names_links)-1, 2)]

total = len(companies)
atual = 0

# === LOG ===
if not os.path.exists(ARQUIVO_LOG):
    with open(ARQUIVO_LOG, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Nome", "Status", "Arquivo"])

with open(ARQUIVO_LOG, newline='', encoding="utf-8") as f:
    ja_processados = {row['Nome'] for row in csv.DictReader(f)}

# === BROWSER ===
options = uc.ChromeOptions()
driver = uc.Chrome(options=options)

# === FUNÇÕES ===
def abrir_proximo():
    global atual
    while atual < total and companies[atual][0] in ja_processados:
        atual += 1
    if atual >= total:
        messagebox.showinfo("Concluído", "Todos os registros foram processados.")
        driver.quit()
        root.quit()
        return

    nome, link = companies[atual]
    label_status.config(text=f"{atual+1} de {total}: {nome}")
    driver.get(link)

def resolver_captcha():
    global atual
    nome, _ = companies[atual]
    try:
        links = driver.find_elements(By.PARTIAL_LINK_TEXT, "Formulário de Referência")
        if links:
            href = links[0].get_attribute("href")
            filename = f"{nome.strip().replace(' ', '_').replace('/', '_')}_FORMULARIO.pdf"
            path = os.path.join(PASTA_FORMULARIOS, filename)
            r = requests.get(href)
            with open(path, "wb") as f:
                f.write(r.content)

            if os.path.getsize(path) > 1000:
                status = "Baixado"
                print(f" {nome} - baixado com sucesso")
            else:
                os.remove(path)
                filename = ""
                status = "Erro: arquivo vazio"
        else:
            status = "Sem Formulário"
            filename = ""
    except Exception as e:
        status = f"Erro: {str(e)}"
        filename = ""

    with open(ARQUIVO_LOG, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow([nome, status, filename])

    atualiza_log_temporario(status)

def atualiza_log_temporario(status):
    label_resultado.config(text=f"Resultado: {status}")

def proximo():
    global atual
    atual += 1
    abrir_proximo()

# === GUI ===
root = tk.Tk()
root.title("Extrator de Formulários da CVM")
root.geometry("600x200")

label_status = tk.Label(root, text="Clique em 'Abrir próximo' para iniciar", font=("Arial", 12))
label_status.pack(pady=10)

btn_abrir = tk.Button(root, text="Abrir próximo", font=("Arial", 11), command=abrir_proximo)
btn_abrir.pack(pady=5)

btn_resolver = tk.Button(root, text="CAPTCHA resolvido (baixar)", font=("Arial", 11), command=resolver_captcha)
btn_resolver.pack(pady=5)

btn_proximo = tk.Button(root, text="Próximo", font=("Arial", 11), command=proximo)
btn_proximo.pack(pady=5)

label_resultado = tk.Label(root, text="Resultado: aguardando...", font=("Arial", 10))
label_resultado.pack(pady=10)

root.mainloop()
