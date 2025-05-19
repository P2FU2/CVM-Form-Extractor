import tkinter as tk
from tkinter import messagebox, ttk
import pandas as pd
import os
import requests
import urllib.parse
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from PIL import Image, ImageTk, ImageEnhance, ImageOps, ImageFilter
from datetime import datetime, timedelta
import pytesseract
from io import BytesIO
from PIL import Image as PILImage
import traceback
import cv2
import numpy as np
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import unicodedata

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
tempo_inicio = None
tempo_medio_por_item = None

# === BROWSER ===
options = uc.ChromeOptions()
driver = uc.Chrome(options=options)

# === LOG ===
# Set para empresas que já tiveram falha única
empresas_falha = set()
# Set para empresas sem formulário
empresas_sem_formulario = set()

def normalizar_nome(nome):
    nome = nome.strip().lower()
    nome = unicodedata.normalize('NFKD', nome)
    nome = ''.join([c for c in nome if not unicodedata.combining(c)])
    return nome

def registrar_log(nome, status, arquivo, texto=""):
    global sucesso, falha
    horario = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    linha = f"[{horario}] {nome} | {status} | {arquivo}\n"
    with open(ARQUIVO_LOG, "a", encoding="utf-8") as f:
        f.write(linha)

    style = "success" if "Baixado" in status else "fail"
    nome_empresa = normalizar_nome(nome)
    if "Baixado" in status:
        sucesso += 1
    elif status == "Sem Formulário":
        if nome_empresa not in empresas_sem_formulario:
            empresas_sem_formulario.add(nome_empresa)
    else:
        if nome_empresa not in empresas_falha and nome_empresa not in empresas_sem_formulario:
            falha += 1
            empresas_falha.add(nome_empresa)

    with open(ARQUIVO_HTML_DIAGNOSTICO, "a", encoding="utf-8") as html:
        html.write(f"<li class='{style}'><strong>{nome}</strong><br>Status: {status}<br>Arquivo: {arquivo if arquivo else 'N/A'}<br>Texto OCR: {texto}<br>")
        nome_base = nome.strip().replace(' ', '_').replace('/', '_')
        processado_path = os.path.join(PASTA_CAPTCHAS, f"{nome_base}_processado.png")
        original_path = os.path.join(PASTA_CAPTCHAS, f"{nome_base}_original.png")
        rel_proc = os.path.relpath(processado_path).replace('\\', '/')
        rel_orig = os.path.relpath(original_path).replace('\\', '/')
        if os.path.exists(processado_path):
            html.write(f"<b>Processado:</b><br><img src='{rel_proc}' height='60'><br>")
        else:
            html.write("<b>Processado:</b> Imagem não disponível<br>")
        if os.path.exists(original_path):
            html.write(f"<b>Original:</b><br><img src='{rel_orig}' height='60'><br>")
        else:
            html.write("<b>Original:</b> Imagem não disponível<br>")
        html.write("</li>")
    atualiza_log_temporario(status)

# === FUNÇÕES ===
def formatar_tempo(segundos):
    if segundos < 60:
        return f"{int(segundos)} segundos"
    elif segundos < 3600:
        minutos = int(segundos / 60)
        segundos = int(segundos % 60)
        return f"{minutos} min {segundos} seg"
    else:
        horas = int(segundos / 3600)
        minutos = int((segundos % 3600) / 60)
        return f"{horas} hora(s) {minutos} min"

def atualizar_tempo_estimado():
    global tempo_medio_por_item, tempo_inicio
    if atual > 0:
        tempo_decorrido = (datetime.now() - tempo_inicio).total_seconds()
        tempo_medio_por_item = tempo_decorrido / atual
        tempo_restante = tempo_medio_por_item * (total - atual)
        return formatar_tempo(tempo_restante)
    return "Calculando..."

def abrir_proximo():
    global atual, tempo_inicio
    if atual == 0:
        tempo_inicio = datetime.now()
    
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
    
    # Atualizar progresso e tempo estimado
    progresso = (atual + 1) / total * 100
    progress_var.set(progresso)
    tempo_restante = atualizar_tempo_estimado()
    label_progresso.config(text=f"Progresso: {progresso:.1f}% | Tempo estimado: {tempo_restante}")
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
        # Espera explícita pelo CAPTCHA
        try:
            img_elem = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//img[contains(@src, 'captcha/aspcaptcha.asp')]"))
            )
        except Exception as e:
            registrar_log(nome, f"Erro OCR: CAPTCHA não encontrado", "")
            label_resultado.config(text="CAPTCHA não encontrado na página.")
            root.after(2000, proximo)
            return
        png_data = img_elem.screenshot_as_png
        image = PILImage.open(BytesIO(png_data))

        nome_base = nome.strip().replace(' ', '_').replace('/', '_')
        original_path = os.path.join(PASTA_CAPTCHAS, f"{nome_base}_original.png")
        image.save(original_path)

        # Redimensionar imagem original para exibição
        img_original_resized = image.resize((150, 50))
        img_original_tk = ImageTk.PhotoImage(img_original_resized)
        captcha_original_label.config(image=img_original_tk)
        captcha_original_label.image = img_original_tk

        # Processar imagem para OCR
        image_proc = aplicar_preprocessamento_opencv(image)
        processado_path = os.path.join(PASTA_CAPTCHAS, f"{nome_base}_processado.png")
        image_proc.save(processado_path)

        # Redimensionar imagem processada para exibição
        img_proc_resized = image_proc.resize((150, 50))
        img_proc_tk = ImageTk.PhotoImage(img_proc_resized)
        captcha_processado_label.config(image=img_proc_tk)
        captcha_processado_label.image = img_proc_tk

        ocr_tentativas = []
        for psm in [6, 7, 8, 13]:
            config = f"--psm {psm} -c tessedit_char_whitelist=0123456789"
            text = pytesseract.image_to_string(image_proc, config=config)
            digits = ''.join(filter(str.isdigit, text))[:4]
            ocr_tentativas.append((psm, text.strip(), digits))
            if len(digits) == 4:
                captcha_text = digits
                break
        else:
            captcha_text = ""

        if captcha_text and len(captcha_text) == 4:
            input_box = driver.find_element(By.NAME, "strCAPTCHA")
            input_box.clear()
            input_box.send_keys(captcha_text + "\n")
            root.after(5000, resolver_captcha)
        elif tentativa < 2:
            root.after(2000, lambda: executar_ocr_captcha(tentativa+1))
        else:
            detalhes = " | ".join([f"psm {p}: {t}" for p, t, _ in ocr_tentativas])
            registrar_log(nome, f"Erro OCR: não leu 4 dígitos", processado_path, captcha_text if captcha_text else detalhes)
            raise ValueError("OCR não conseguiu identificar 4 dígitos após múltiplos PSM")

    except Exception as e:
        erro_detalhe = traceback.format_exc(limit=1)
        registrar_log(nome, f"Erro OCR: {str(e)}", "")
        label_resultado.config(text=f"Erro OCR: {str(e)}")
        root.after(2000, proximo)


def resolver_captcha():
    global atual
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
    registrar_log(nome, status, filename)
    if ocr_ativo:
        proximo()

def atualiza_log_temporario(status):
    label_resultado.config(text=f"Resultado: {status}")
    label_sucesso_falha.config(text=f"Sucesso: {sucesso} | Falha: {falha} | Sem Formulário: {len(empresas_sem_formulario)}")
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

def obter_empresas_com_formulario():
    arquivos = os.listdir(PASTA_FORMULARIOS)
    empresas_com_formulario = set()
    for arq in arquivos:
        if arq.lower().endswith(".pdf"):
            nome_empresa = arq.replace("_FORMULARIO.pdf", "").replace("_", " ")
            empresas_com_formulario.add(nome_empresa.lower())
    return empresas_com_formulario

def obter_empresas_sem_formulario():
    return empresas_sem_formulario

def filtrar_empresas_faltantes():
    baixados = obter_empresas_com_formulario()
    sem_formulario = obter_empresas_sem_formulario()
    faltantes = []
    for nome, link in companies:
        nome_limpo = nome.strip().lower()
        if nome_limpo not in baixados and nome_limpo not in sem_formulario:
            faltantes.append((nome, link))
    return faltantes

def reprocessar_pendentes():
    global companies, total, atual, sucesso, falha
    atual = 0
    sucesso = 0
    falha = 0
    companies = filtrar_empresas_faltantes()
    total = len(companies)

    if total == 0:
        messagebox.showinfo("Reprocessamento", "Todas as empresas já foram processadas com sucesso.")
        return

    habilitar_botoes()
    abrir_proximo()

class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip = None
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.leave)

    def enter(self, event=None):
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 25
        
        self.tooltip = tk.Toplevel(self.widget)
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry(f"+{x}+{y}")
        
        label = tk.Label(self.tooltip, text=self.text, justify=tk.LEFT,
                        background="#ffffff", relief=tk.SOLID, borderwidth=1,
                        font=("Arial", "8", "normal"))
        label.pack()

    def leave(self, event=None):
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None

# === GUI ===
root = tk.Tk()
root.title("Extrator de Formulários da CVM")
root.geometry("800x600")
root.configure(bg="#1e1e1e")

ascii_art = """
                                         ██████╗██╗   ██╗███╗   ███╗                                               
                                        ██╔════╝██║   ██║████╗ ████║                                               
                                        ██║     ██║   ██║██╔████╔██║                                               
                                        ██║     ╚██╗ ██╔╝██║╚██╔╝██║                                               
                                        ╚██████╗ ╚████╔╝ ██║ ╚═╝ ██║                                               
                                         ╚═════╝  ╚═══╝  ╚═╝     ╚═╝                                               
                                                                                                                   
███████╗ ██████╗ ██████╗ ███╗   ███╗    ███████╗██╗  ██╗████████╗██████╗  █████╗  ██████╗████████╗ ██████╗ ██████╗ 
██╔════╝██╔═══██╗██╔══██╗████╗ ████║    ██╔════╝╚██╗██╔╝╚══██╔══╝██╔══██╗██╔══██╗██╔════╝╚══██╔══╝██╔═══██╗██╔══██╗
█████╗  ██║   ██║██████╔╝██╔████╔██║    █████╗   ╚███╔╝    ██║   ██████╔╝███████║██║        ██║   ██║   ██║██████╔╝
██╔══╝  ██║   ██║██╔══██╗██║╚██╔╝██║    ██╔══╝   ██╔██╗    ██║   ██╔══██╗██╔══██║██║        ██║   ██║   ██║██╔══██╗
██║     ╚██████╔╝██║  ██║██║ ╚═╝ ██║    ███████╗██╔╝ ██╗   ██║   ██║  ██║██║  ██║╚██████╗   ██║   ╚██████╔╝██║  ██║
╚═╝      ╚═════╝ ╚═╝  ╚═╝╚═╝     ╚═╝    ╚══════╝╚═╝  ╚═╝   ╚═╝   ╚═╝  ╚═╝╚═╝  ╚═╝ ╚═════╝   ╚═╝    ╚═════╝ ╚═╝  ╚═╝
                                                                                                                   
"""


label_ascii = tk.Label(root, text=ascii_art, font=("Courier", 16, "bold"), fg="#FFA500", bg="#1e1e1e")
label_ascii.pack(pady=(5, 0))

label_by = tk.Label(root, text="by P2FU2 - Version alpha 1.8", font=("Arial", 10, "italic"), fg="#FFA500", bg="#1e1e1e")
label_by.pack()

label_status = tk.Label(root, text="Clique em 'Iniciar' para comecar a leitura dos formularios", font=("Arial", 12), fg="#FFA500", bg="#1e1e1e")
label_status.pack(pady=5)

btn_abrir = tk.Button(root, text="Iniciar", font=("Arial", 11), command=abrir_proximo, bg="#2e2e2e", fg="#FFA500", width=25)
btn_abrir.pack(pady=5)

btn_resolver = tk.Button(root, text="CAPTCHA resolvido (baixar)", font=("Arial", 11), command=resolver_captcha, bg="#2e2e2e", fg="#FFA500", width=25)
btn_resolver.pack(pady=5)

btn_proximo = tk.Button(root, text="Próximo", font=("Arial", 11), command=proximo, bg="#2e2e2e", fg="#FFA500", width=25)
btn_proximo.pack(pady=5)

btn_pular = tk.Button(root, text="Pular Empresa", font=("Arial", 11), command=pular, bg="#882222", fg="white", width=25)
btn_pular.pack(pady=5)

btn_ocr = tk.Button(root, text="Iniciar Leitura OCR Automática", font=("Arial", 11), command=iniciar_ocr_auto, bg="#2e2e2e", fg="#FFA500", width=30)
btn_ocr.pack(pady=5)

btn_abort_ocr = tk.Button(root, text="Parar OCR e voltar para modo manual", font=("Arial", 11), command=abortar_ocr, bg="#882222", fg="white", width=30, state="disabled")
btn_abort_ocr.pack(pady=5)

btn_reprocessar = tk.Button(root, text="Reprocessar Pendentes", font=("Arial", 11), command=reprocessar_pendentes, bg="#2e2e2e", fg="#00ccff", width=30)
btn_reprocessar.pack(pady=5)

# Frame para as imagens do CAPTCHA
captcha_frame = tk.Frame(root, bg="#1e1e1e")
captcha_frame.pack(pady=5)

# Label para a imagem original
label_original = tk.Label(captcha_frame, text="Original:", font=("Arial", 10), fg="#FFA500", bg="#1e1e1e")
label_original.pack(side=tk.LEFT, padx=10)

captcha_original_label = tk.Label(captcha_frame, bg="#1e1e1e")
captcha_original_label.pack(side=tk.LEFT, padx=10)

# Label para a imagem processada
label_processado = tk.Label(captcha_frame, text="Processado:", font=("Arial", 10), fg="#FFA500", bg="#1e1e1e")
label_processado.pack(side=tk.LEFT, padx=10)

captcha_processado_label = tk.Label(captcha_frame, bg="#1e1e1e")
captcha_processado_label.pack(side=tk.LEFT, padx=10)

label_resultado = tk.Label(root, text="Resultado: aguardando...", font=("Arial", 10), fg="#FFA500", bg="#1e1e1e")
label_resultado.pack(pady=10)

label_progresso = tk.Label(root, text="Progresso: 0% | Tempo estimado: --", font=("Arial", 10), fg="#FFA500", bg="#1e1e1e")
label_progresso.pack(pady=5)

# Nova label para mostrar sucesso, falha e sem formulário
label_sucesso_falha = tk.Label(root, text="Sucesso: 0 | Falha: 0 | Sem Formulário: 0", font=("Arial", 11, "bold"), fg="#FFA500", bg="#1e1e1e")
label_sucesso_falha.pack(pady=2)

style = ttk.Style()
style.theme_use('clam')
style.configure("orange.Horizontal.TProgressbar", troughcolor='#333333', background='#FFA500', thickness=20)
progress_var = tk.DoubleVar()
progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100, length=600, style="orange.Horizontal.TProgressbar")
progress_bar.pack(pady=5)

if os.path.exists(LOGO_PATH):
    img = Image.open(LOGO_PATH)
    img = img.resize((100, 40), Image.LANCZOS)
    logo_img = ImageTk.PhotoImage(img)
    logo_label = tk.Label(root, image=logo_img, bg="#1e1e1e")
    logo_label.place(relx=1.0, rely=1.0, x=-10, y=-10, anchor="se")

# Adicionar tooltips aos botões
ToolTip(btn_abrir, "Inicia o processo de extração do primeiro formulário da lista")
ToolTip(btn_resolver, "Confirma que o CAPTCHA foi resolvido e baixa o formulário")
ToolTip(btn_proximo, "Avança para o próximo formulário da lista")
ToolTip(btn_pular, "Pula o formulário atual e avança para o próximo")
ToolTip(btn_ocr, "Inicia o processo automático de leitura de CAPTCHA usando OCR")
ToolTip(btn_abort_ocr, "Interrompe o processo automático de OCR e retorna ao modo manual")
ToolTip(btn_reprocessar, "Reprocessa apenas os formulários que falharam anteriormente")

# Função para ler empresas sem formulário do log
def carregar_empresas_sem_formulario():
    if not os.path.exists(ARQUIVO_LOG):
        return
    with open(ARQUIVO_LOG, encoding="utf-8") as f:
        for linha in f:
            if "| Sem Formulário |" in linha:
                partes = linha.split("|")
                if len(partes) > 1:
                    nome = normalizar_nome(partes[1])
                    empresas_sem_formulario.add(nome)
carregar_empresas_sem_formulario()

root.mainloop()