import tkinter as tk
from tkinter import messagebox, ttk, Menu
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
import logging
import json
from functools import lru_cache
from typing import Dict, Any, Optional
import threading
from queue import Queue
import time

# === CONFIGURAÇÕES ===
ARQUIVO_PLANILHA = "CVM_Links.xlsx"
ARQUIVO_LOG = "resultado_extracao.log"
PASTA_FORMULARIOS = "formularios"
PASTA_CAPTCHAS = "captchas"
ARQUIVO_HTML_DIAGNOSTICO = "diagnostico_captchas.html"
ARQUIVO_ESTADO = "estado_sessao.json"
DEBUG = False
MAX_RETRIES = 3
THREAD_COUNT = 2

# Configuração do logging
logging.basicConfig(
    level=logging.DEBUG if DEBUG else logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(ARQUIVO_LOG, encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# === FUNÇÕES AUXILIARES ===
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

class ImageCache:
    def __init__(self):
        self._cache: Dict[str, ImageTk.PhotoImage] = {}
        self._lock = threading.Lock()
    
    def get_photo_image(self, image: Image.Image, size: tuple) -> ImageTk.PhotoImage:
        """Retorna uma imagem do cache ou cria uma nova se não existir"""
        key = f"{size[0]}_{size[1]}"
        with self._lock:
            if key not in self._cache:
                resized = image.resize(size, Image.LANCZOS)
                self._cache[key] = ImageTk.PhotoImage(resized)
            return self._cache[key]

class EstadoSessao:
    def __init__(self):
        self.atual = 0
        self.sucesso = 0
        self.falha = 0
        self.empresas_processadas = set()
        self.ultima_empresa = None
        self._lock = threading.Lock()
    
    def salvar(self):
        """Salva o estado atual em um arquivo JSON"""
        with self._lock:
            estado = {
                'atual': self.atual,
                'sucesso': self.sucesso,
                'falha': self.falha,
                'empresas_processadas': list(self.empresas_processadas),
                'ultima_empresa': self.ultima_empresa
            }
            with open(ARQUIVO_ESTADO, 'w', encoding='utf-8') as f:
                json.dump(estado, f, ensure_ascii=False, indent=2)
    
    @classmethod
    def carregar(cls) -> 'EstadoSessao':
        """Carrega o estado de uma sessão anterior"""
        estado = cls()
        if os.path.exists(ARQUIVO_ESTADO):
            try:
                with open(ARQUIVO_ESTADO, 'r', encoding='utf-8') as f:
                    dados = json.load(f)
                    estado.atual = dados.get('atual', 0)
                    estado.sucesso = dados.get('sucesso', 0)
                    estado.falha = dados.get('falha', 0)
                    estado.empresas_processadas = set(dados.get('empresas_processadas', []))
                    estado.ultima_empresa = dados.get('ultima_empresa')
            except Exception as e:
                logger.error(f"Erro ao carregar estado: {e}")
        return estado

class ProcessadorThread(threading.Thread):
    def __init__(self, queue: Queue, estado: EstadoSessao):
        super().__init__()
        self.queue = queue
        self.estado = estado
        self.driver = None
        self.running = True

    def run(self):
        try:
            options = uc.ChromeOptions()
            self.driver = uc.Chrome(options=options)
            while self.running:
                try:
                    nome, link = self.queue.get(timeout=1)
                    self.processar_empresa(nome, link)
                    self.queue.task_done()
                except Queue.Empty:
                    continue
        except Exception as e:
            logger.error(f"Erro na thread de processamento: {e}")
        finally:
            if self.driver:
                self.driver.quit()

    def processar_empresa(self, nome: str, link: str):
        try:
            self.driver.get(link)
            for tentativa in range(MAX_RETRIES):
                try:
                    img_elem = self.driver.find_element(By.XPATH, "//img[contains(@src, 'captcha/aspcaptcha.asp')]")
                    png_data = img_elem.screenshot_as_png
                    image = PILImage.open(BytesIO(png_data))
                    
                    nome_base = nome.strip().replace(' ', '_').replace('/', '_')
                    original_path = os.path.join(PASTA_CAPTCHAS, f"{nome_base}_original.png")
                    image.save(original_path)
                    
                    image_proc = aplicar_preprocessamento_opencv(image)
                    processado_path = os.path.join(PASTA_CAPTCHAS, f"{nome_base}_processado.png")
                    image_proc.save(processado_path)
                    
                    captcha_text = self.tentar_ocr(image_proc)
                    if captcha_text:
                        input_box = self.driver.find_element(By.NAME, "strCAPTCHA")
                        input_box.clear()
                        input_box.send_keys(captcha_text + "\n")
                        time.sleep(2)
                        
                        if self.verificar_sucesso():
                            self.baixar_formulario(nome)
                            break
                except Exception as e:
                    logger.error(f"Tentativa {tentativa + 1} falhou para {nome}: {e}")
                    if tentativa == MAX_RETRIES - 1:
                        raise
        except Exception as e:
            logger.error(f"Erro ao processar {nome}: {e}")
            with self.estado._lock:
                self.estado.falha += 1
                self.estado.empresas_processadas.add(nome)
                self.estado.salvar()

    def tentar_ocr(self, image: Image.Image) -> Optional[str]:
        for psm in [6, 7, 8, 13]:
            config = f"--psm {psm} -c tessedit_char_whitelist=0123456789"
            text = pytesseract.image_to_string(image, config=config)
            digits = ''.join(filter(str.isdigit, text))[:4]
            if len(digits) == 4:
                return digits
        return None

    def verificar_sucesso(self) -> bool:
        try:
            return bool(self.driver.find_elements(By.PARTIAL_LINK_TEXT, "Formulário de Referência"))
        except:
            return False

    def baixar_formulario(self, nome: str):
        try:
            links = self.driver.find_elements(By.PARTIAL_LINK_TEXT, "Formulário de Referência")
            if links:
                href = links[0].get_attribute("href")
                full_url = urllib.parse.urljoin(self.driver.current_url, href)
                filename = f"{nome.strip().replace(' ', '_').replace('/', '_')}_FORMULARIO.pdf"
                path = os.path.join(PASTA_FORMULARIOS, filename)
                
                r = requests.get(full_url)
                with open(path, "wb") as f:
                    f.write(r.content)
                
                if os.path.getsize(path) > 1000:
                    with self.estado._lock:
                        self.estado.sucesso += 1
                        self.estado.empresas_processadas.add(nome)
                        self.estado.salvar()
                    return True
        except Exception as e:
            logger.error(f"Erro ao baixar formulário para {nome}: {e}")
        return False

# === GUI ===
class MainApplication:
    def __init__(self, root):
        self.root = root
        self.root.title("Extrator de Formulários da CVM")
        self.root.geometry("800x600")
        self.root.configure(bg="#1e1e1e")
        
        # Inicialização dos dados
        os.makedirs(PASTA_FORMULARIOS, exist_ok=True)
        os.makedirs(PASTA_CAPTCHAS, exist_ok=True)
        df = pd.read_excel(ARQUIVO_PLANILHA, header=None)
        names_links = df[0].tolist()
        self.companies = [(names_links[i], names_links[i+1]) for i in range(0, len(names_links)-1, 2)]
        self.total = len(self.companies)
        
        self.estado = EstadoSessao.carregar()
        self.image_cache = ImageCache()
        self.processadores = []
        self.queue = Queue()
        self.ocr_ativo = False
        
        self.setup_menu()
        self.setup_gui()
        self.setup_status_bar()
        self.setup_shortcuts()
        
        # Verificação do Tesseract
        try:
            tesseract_version = pytesseract.get_tesseract_version()
            logger.info(f"Tesseract versão: {tesseract_version}")
            self.status_bar.config(text=f"Tesseract {tesseract_version} | Pronto")
        except Exception as e:
            logger.error(f"Erro ao verificar Tesseract: {e}")
            messagebox.showerror("Erro", "Tesseract não está configurado corretamente. Por favor, verifique a instalação.")
            exit(1)

    def setup_menu(self):
        menubar = Menu(self.root)
        
        # Menu Arquivo
        file_menu = Menu(menubar, tearoff=0)
        file_menu.add_command(label="Abrir Planilha", command=self.abrir_planilha)
        file_menu.add_command(label="Salvar Estado", command=self.estado.salvar)
        file_menu.add_separator()
        file_menu.add_command(label="Sair", command=self.on_closing)
        menubar.add_cascade(label="Arquivo", menu=file_menu)
        
        # Menu Processamento
        process_menu = Menu(menubar, tearoff=0)
        process_menu.add_command(label="Iniciar Processamento", command=self.iniciar_processamento)
        process_menu.add_command(label="Parar Processamento", command=self.parar_processamento)
        process_menu.add_command(label="Reprocessar Falhas", command=self.reprocessar_falhas)
        menubar.add_cascade(label="Processamento", menu=process_menu)
        
        # Menu Configurações
        config_menu = Menu(menubar, tearoff=0)
        config_menu.add_checkbutton(label="Modo Debug", variable=tk.BooleanVar(value=DEBUG))
        config_menu.add_command(label="Configurar Threads", command=self.configurar_threads)
        menubar.add_cascade(label="Configurações", menu=config_menu)
        
        self.root.config(menu=menubar)

    def setup_gui(self):
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
        self.label_ascii = tk.Label(self.root, text=ascii_art, font=("Courier", 8), fg="#FFA500", bg="#1e1e1e")
        self.label_ascii.pack(pady=(5, 0))

        self.label_by = tk.Label(self.root, text="by P2FU2", font=("Arial", 10, "italic"), fg="#FFA500", bg="#1e1e1e")
        self.label_by.pack()

        self.label_status = tk.Label(self.root, text="Clique em 'Iniciar' para comecar a analise dos formularios", font=("Arial", 12), fg="#FFA500", bg="#1e1e1e")
        self.label_status.pack(pady=5)

        self.btn_abrir = tk.Button(self.root, text="Iniciar", font=("Arial", 11), command=self.abrir_proximo, bg="#2e2e2e", fg="#FFA500", width=25)
        self.btn_abrir.pack(pady=5)

        self.btn_resolver = tk.Button(self.root, text="CAPTCHA resolvido (baixar)", font=("Arial", 11), command=self.resolver_captcha, bg="#2e2e2e", fg="#FFA500", width=25)
        self.btn_resolver.pack(pady=5)

        self.btn_proximo = tk.Button(self.root, text="Próximo", font=("Arial", 11), command=self.proximo, bg="#2e2e2e", fg="#FFA500", width=25)
        self.btn_proximo.pack(pady=5)

        self.btn_pular = tk.Button(self.root, text="Pular Empresa", font=("Arial", 11), command=self.pular, bg="#882222", fg="white", width=25)
        self.btn_pular.pack(pady=5)

        self.btn_ocr = tk.Button(self.root, text="Iniciar Leitura OCR Automática", font=("Arial", 11), command=self.iniciar_ocr_auto, bg="#444444", fg="#FFA500", width=30)
        self.btn_ocr.pack(pady=5)

        self.btn_abort_ocr = tk.Button(self.root, text="Parar OCR e voltar para modo manual", font=("Arial", 11), command=self.abortar_ocr, bg="#882222", fg="white", width=30, state="disabled")
        self.btn_abort_ocr.pack(pady=5)

        self.btn_reprocessar = tk.Button(self.root, text="Reprocessar Pendentes", font=("Arial", 11), command=self.reprocessar_pendentes, bg="#2e2e2e", fg="#00ccff", width=30)
        self.btn_reprocessar.pack(pady=5)

        self.captcha_image_label = tk.Label(self.root, bg="#1e1e1e")
        self.captcha_image_label.pack(pady=5)

        self.label_resultado = tk.Label(self.root, text="Resultado: aguardando...", font=("Arial", 10), fg="#FFA500", bg="#1e1e1e")
        self.label_resultado.pack(pady=10)

        style = ttk.Style()
        style.theme_use('clam')
        style.configure("orange.Horizontal.TProgressbar", troughcolor='#333333', background='#FFA500', thickness=20)
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.root, variable=self.progress_var, maximum=100, length=600, style="orange.Horizontal.TProgressbar")
        self.progress_bar.pack(pady=10)


    def setup_status_bar(self):
        self.status_bar = ttk.Label(self.root, text="Pronto", relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def setup_shortcuts(self):
        self.root.bind('<Control-o>', lambda e: self.abrir_planilha())
        self.root.bind('<Control-s>', lambda e: self.estado.salvar())
        self.root.bind('<Control-q>', lambda e: self.on_closing())
        self.root.bind('<F5>', lambda e: self.iniciar_processamento())
        self.root.bind('<F6>', lambda e: self.parar_processamento())

    def abrir_planilha(self):
        # Implementar seleção de arquivo
        pass

    def iniciar_processamento(self):
        self.status_bar.config(text="Iniciando processamento...")
        time.sleep(2)  # Aguarda o ChromeDriver anterior ser liberado
        for _ in range(THREAD_COUNT):
            processador = ProcessadorThread(self.queue, self.estado)
            processador.start()
            self.processadores.append(processador)
        
        # Adicionar empresas à fila
        for nome, link in self.companies:
            if nome not in self.estado.empresas_processadas:
                self.queue.put((nome, link))

    def parar_processamento(self):
        self.status_bar.config(text="Parando processamento...")
        for processador in self.processadores:
            processador.running = False
        self.processadores.clear()

    def reprocessar_falhas(self):
        # Implementar reprocessamento de falhas
        pass

    def configurar_threads(self):
        # Implementar configuração de threads
        pass

    def on_closing(self):
        if messagebox.askokcancel("Sair", "Deseja realmente sair? O progresso será salvo."):
            self.parar_processamento()
            self.estado.salvar()
            self.root.destroy()

    def abrir_proximo(self):
        if self.estado.atual >= self.total:
            messagebox.showinfo("Concluído", f"Todos os registros foram processados.\nSucesso: {self.estado.sucesso} | Falha: {self.estado.falha}")
            self.estado.salvar()
            return

        nome, link = self.companies[self.estado.atual]
        self.label_status.config(text=f"{self.estado.atual+1} de {self.total}: {nome}")
        self.label_resultado.config(text="Aguardando resolução do CAPTCHA...")
        self.progress_var.set((self.estado.atual + 1) / self.total * 100)
        self.root.update_idletasks()

    def resolver_captcha(self):
        nome, _ = self.companies[self.estado.atual]
        self.status_bar.config(text="Buscando link do formulário...")
        self.estado.empresas_processadas.add(nome)
        self.estado.salvar()
        self.proximo()

    def proximo(self):
        self.estado.atual += 1
        self.estado.salvar()
        self.abrir_proximo()

    def pular(self):
        nome, _ = self.companies[self.estado.atual]
        self.estado.falha += 1
        self.estado.empresas_processadas.add(nome)
        self.estado.salvar()
        self.proximo()

    def iniciar_ocr_auto(self):
        self.ocr_ativo = True
        self.desabilitar_botoes()
        self.iniciar_processamento()

    def abortar_ocr(self):
        self.ocr_ativo = False
        self.habilitar_botoes()
        self.label_resultado.config(text="Leitura OCR interrompida. Continue manualmente.")

    def reprocessar_pendentes(self):
        self.estado.atual = 0
        self.estado.sucesso = 0
        self.estado.falha = 0
        self.companies = [c for c in self.companies if c[0] not in self.estado.empresas_processadas]
        self.total = len(self.companies)

        if self.total == 0:
            messagebox.showinfo("Reprocessamento", "Todas as empresas já foram processadas com sucesso.")
            return

        self.abrir_proximo()

    def desabilitar_botoes(self):
        self.btn_abrir.config(state="disabled")
        self.btn_resolver.config(state="disabled")
        self.btn_proximo.config(state="disabled")
        self.btn_pular.config(state="disabled")
        self.btn_ocr.config(state="disabled")
        self.btn_abort_ocr.config(state="normal")

    def habilitar_botoes(self):
        self.btn_abrir.config(state="normal")
        self.btn_resolver.config(state="normal")
        self.btn_proximo.config(state="normal")
        self.btn_pular.config(state="normal")
        self.btn_ocr.config(state="normal")
        self.btn_abort_ocr.config(state="disabled")

if __name__ == "__main__":
    root = tk.Tk()
    app = MainApplication(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()