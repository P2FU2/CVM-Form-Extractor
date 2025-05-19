import tkinter as tk
from tkinter import messagebox, ttk
import pandas as pd
import os
import requests
import urllib.parse
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from PIL import Image, ImageTk, ImageEnhance, ImageOps, ImageFilter
from datetime import datetime
import pytesseract
from io import BytesIO
from PIL import Image as PILImage
import traceback
import cv2
import numpy as np

# === CONFIGURAÇÕES ===
ARQUIVO_PLANILHA = "CVM_Links.xlsx"
ARQUIVO_LOG = "resultado_extracao.log"
PASTA_FORMULARIOS = "formularios"
PASTA_CAPTCHAS = "captchas"
ARQUIVO_HTML_DIAGNOSTICO = "diagnostico_captchas.html"
LOGO_PATH = "cvm_logo.png"

# === PREPARO ===
os.makedirs(PASTA_FORMULARIOS, exist_ok=True)
os.makedirs(PASTA_CAPTCHAS, exist_ok=True)
df = pd.read_excel(ARQUIVO_PLANILHA, header=None)
names_links = df[0].tolist()
companies = [(names_links[i], names_links[i+1]) for i in range(0, len(names_links)-1, 2)]

# === HTML DIAGNÓSTICO ===
with open(ARQUIVO_HTML_DIAGNOSTICO, "w", encoding="utf-8") as f:
    f.write("""
    <html><head><title>Diagnóstico OCR</title>
    <style>
    body { font-family: Arial; background: #1e1e1e; color: #f0f0f0; }
    li { margin-bottom: 20px; }
    .success { color: #00ff00; }
    .fail { color: #ff5555; }
    img { border: 1px solid #ccc; margin-top: 5px; }
    </style>
    </head><body><h1>Resultados OCR</h1><ul>
    """)

total = len(companies)
atual = 0
sucesso = 0
falha = 0
ocr_ativo = False

# === BROWSER ===
options = uc.ChromeOptions()
driver = uc.Chrome(options=options)

# === LOG ===
def registrar_log(nome, status, arquivo):
    global sucesso, falha
    horario = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    linha = f"[{horario}] {nome} | {status} | {arquivo}\n"
    with open(ARQUIVO_LOG, "a", encoding="utf-8") as f:
        f.write(linha)

    style = "success" if "Baixado" in status else "fail"
    if "Baixado" in status:
        sucesso += 1
    else:
        falha += 1

    with open(ARQUIVO_HTML_DIAGNOSTICO, "a", encoding="utf-8") as html:
        html.write(f"<li class='{style}'><strong>{nome}</strong><br>Status: {status}<br>Arquivo: {arquivo if arquivo else 'N/A'}<br>")
        if arquivo and os.path.exists(arquivo):
            rel_path = os.path.relpath(arquivo).replace('\\', '/')
            html.write(f"<img src='{rel_path}' height='60'><br>")
        html.write("</li>")

# === FUNÇÕES ===
def abrir_proximo():
    global atual
    if atual >= total:
        with open(ARQUIVO_HTML_DIAGNOSTICO, "a", encoding="utf-8") as html:
            html.write("</ul></body></html>")
        messagebox.showinfo("Concluído", f"Todos os registros foram processados.\nSucesso: {sucesso} | Falha: {falha}")
        driver.quit()
        root.quit()
        return

    nome, link = companies[atual]
    label_status.config(text=f"{atual+1} de {total}: {nome}")
    label_resultado.config(text="Aguardando resolução do CAPTCHA...")
    driver.get(link)
    progress_var.set((atual + 1) / total * 100)
    root.update_idletasks()

    if ocr_ativo:
        root.after(3000, executar_ocr_captcha)

def aplicar_preprocessamento_opencv(image_pil):
    image_np = np.array(image_pil.convert("L"))
    _, bin_img = cv2.threshold(image_np, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    bin_img = cv2.bitwise_not(bin_img)
    bin_img = cv2.medianBlur(bin_img, 3)
    kernel = np.ones((1, 1), np.uint8)
    bin_img = cv2.dilate(bin_img, kernel, iterations=1)
    bin_img = cv2.erode(bin_img, kernel, iterations=1)
    bin_img = cv2.resize(bin_img, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
    return Image.fromarray(bin_img)

def executar_ocr_captcha(tentativa=1):
    global atual
    nome, _ = companies[atual]
    try:
        img_elem = driver.find_element(By.XPATH, "//img[contains(@src, 'captcha/aspcaptcha.asp')]")
        png_data = img_elem.screenshot_as_png
        image = PILImage.open(BytesIO(png_data))

        nome_base = nome.strip().replace(' ', '_').replace('/', '_')
        original_path = os.path.join(PASTA_CAPTCHAS, f"{nome_base}_original.png")
        image.save(original_path)

        img_resized = image.resize((150, 50))
        img_tk = ImageTk.PhotoImage(img_resized)
        captcha_image_label.config(image=img_tk)
        captcha_image_label.image = img_tk

        # === PRÉ-PROCESSAMENTO COM OPENCV ===
        image_proc = aplicar_preprocessamento_opencv(image)

        processado_path = os.path.join(PASTA_CAPTCHAS, f"{nome_base}_processado.png")
        image_proc.save(processado_path)

        captcha_text = pytesseract.image_to_string(
            image_proc,
            config='--psm 8 -c tessedit_char_whitelist=0123456789'
        )
        captcha_text = ''.join(filter(str.isdigit, captcha_text))[:4]

        if captcha_text and len(captcha_text) == 4:
            input_box = driver.find_element(By.NAME, "strCAPTCHA")
            input_box.clear()
            input_box.send_keys(captcha_text)
            driver.find_element(By.XPATH, "//input[@value='Prosseguir']").click()
            root.after(5000, resolver_captcha)
        elif tentativa < 2:
            root.after(2000, lambda: executar_ocr_captcha(tentativa+1))
        else:
            registrar_log(nome, "Erro OCR: não leu 4 dígitos", processado_path)
            raise ValueError("OCR não conseguiu identificar 4 dígitos após 2 tentativas")
    except Exception as e:
        erro_detalhe = traceback.format_exc(limit=1)
        registrar_log(nome, f"Erro OCR: {str(e)}", "")
        label_resultado.config(text=f"Erro OCR: {str(e)}")
        root.after(2000, proximo)


def resolver_captcha():
    global atual, sucesso, falha
    nome, _ = companies[atual]
    try:
        label_resultado.config(text="Buscando link do formulário...")
        links = driver.find_elements(By.PARTIAL_LINK_TEXT, "Formulário de Referência")
        if links:
            href = links[0].get_attribute("href")
            full_url = urllib.parse.urljoin(driver.current_url, href)
            filename = f"{nome.strip().replace(' ', '_').replace('/', '_')}_FORMULARIO.pdf"
            path = os.path.join(PASTA_FORMULARIOS, filename)
            r = requests.get(full_url)
            with open(path, "wb") as f:
                f.write(r.content)
            if os.path.getsize(path) > 1000:
                status = "Baixado"
                sucesso += 1
            else:
                os.remove(path)
                filename = ""
                status = "Erro: arquivo vazio"
                falha += 1
        else:
            status = "Sem Formulário"
            filename = ""
            falha += 1
    except Exception as e:
        status = f"Erro: {str(e)}"
        filename = ""
        falha += 1
    registrar_log(nome, status, filename)
    atualiza_log_temporario(status)
    if ocr_ativo:
        proximo()

def atualiza_log_temporario(status):
    label_resultado.config(text=f"Resultado: {status} | Sucesso: {sucesso} | Falha: {falha}")
    progress_var.set((atual + 1) / total * 100)
    root.update_idletasks()

def proximo():
    global atual
    atual += 1
    abrir_proximo()

def pular():
    global falha
    nome, _ = companies[atual]
    status = "Pulado pelo usuário"
    registrar_log(nome, status, "")
    falha += 1
    atualiza_log_temporario(status)
    proximo()

def iniciar_ocr_auto():
    global ocr_ativo
    ocr_ativo = True
    desabilitar_botoes()
    abrir_proximo()

def abortar_ocr():
    global ocr_ativo
    ocr_ativo = False
    habilitar_botoes()
    label_resultado.config(text="Leitura OCR interrompida. Continue manualmente.")

def desabilitar_botoes():
    btn_abrir.config(state="disabled")
    btn_resolver.config(state="disabled")
    btn_proximo.config(state="disabled")
    btn_pular.config(state="disabled")
    btn_ocr.config(state="disabled")
    btn_abort_ocr.config(state="normal")

def habilitar_botoes():
    btn_abrir.config(state="normal")
    btn_resolver.config(state="normal")
    btn_proximo.config(state="normal")
    btn_pular.config(state="normal")
    btn_ocr.config(state="normal")
    btn_abort_ocr.config(state="disabled")

# === GUI ===
root = tk.Tk()
root.title("Extrator de Formulários da CVM")
root.geometry("800x600")
root.configure(bg="#1e1e1e")

ascii_art = "FORM EXTRACTOR"
label_ascii = tk.Label(root, text=ascii_art, font=("Courier", 16, "bold"), fg="#FFA500", bg="#1e1e1e")
label_ascii.pack(pady=(5, 0))

label_by = tk.Label(root, text="by P2FU2", font=("Arial", 10, "italic"), fg="#FFA500", bg="#1e1e1e")
label_by.pack()

label_status = tk.Label(root, text="Clique em 'Abrir próximo' para iniciar", font=("Arial", 12), fg="#FFA500", bg="#1e1e1e")
label_status.pack(pady=5)

btn_abrir = tk.Button(root, text="Abrir próximo", font=("Arial", 11), command=abrir_proximo, bg="#2e2e2e", fg="#FFA500", width=25)
btn_abrir.pack(pady=5)

btn_resolver = tk.Button(root, text="CAPTCHA resolvido (baixar)", font=("Arial", 11), command=resolver_captcha, bg="#2e2e2e", fg="#FFA500", width=25)
btn_resolver.pack(pady=5)

btn_proximo = tk.Button(root, text="Próximo", font=("Arial", 11), command=proximo, bg="#2e2e2e", fg="#FFA500", width=25)
btn_proximo.pack(pady=5)

btn_pular = tk.Button(root, text="Pular Empresa", font=("Arial", 11), command=pular, bg="#882222", fg="white", width=25)
btn_pular.pack(pady=5)

btn_ocr = tk.Button(root, text="Iniciar Leitura OCR Automática", font=("Arial", 11), command=iniciar_ocr_auto, bg="#444444", fg="#FFA500", width=30)
btn_ocr.pack(pady=5)

btn_abort_ocr = tk.Button(root, text="Parar OCR e voltar para modo manual", font=("Arial", 11), command=abortar_ocr, bg="#AA0000", fg="white", width=30, state="disabled")
btn_abort_ocr.pack(pady=5)

captcha_image_label = tk.Label(root, bg="#1e1e1e")
captcha_image_label.pack(pady=5)

label_resultado = tk.Label(root, text="Resultado: aguardando...", font=("Arial", 10), fg="#FFA500", bg="#1e1e1e")
label_resultado.pack(pady=10)

style = ttk.Style()
style.theme_use('clam')
style.configure("orange.Horizontal.TProgressbar", troughcolor='#333333', background='#FFA500', thickness=20)
progress_var = tk.DoubleVar()
progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100, length=600, style="orange.Horizontal.TProgressbar")
progress_bar.pack(pady=10)

if os.path.exists(LOGO_PATH):
    img = Image.open(LOGO_PATH)
    img = img.resize((100, 40), Image.LANCZOS)
    logo_img = ImageTk.PhotoImage(img)
    logo_label = tk.Label(root, image=logo_img, bg="#1e1e1e")
    logo_label.place(relx=1.0, rely=1.0, x=-10, y=-10, anchor="se")

root.mainloop()