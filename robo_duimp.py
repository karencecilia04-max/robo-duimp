#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Robo DUIMP: extrai dados das declaracoes de importacao (DUIMP) do Portal Unico Siscomex.

A pessoa entra no portal com o proprio certificado digital, num Chrome aberto pelo robo.
Depois do login, o Selenium se conecta a essa janela, busca as DUIMPs do CNPJ no periodo,
le as abas Carga e Item de cada uma e grava uma planilha. Opcionalmente, gera o extrato
em PDF de cada DUIMP na mesma visita.

A versao 5 le o que aparece na tela em vez de dados internos do Angular, que mudavam
a cada atualizacao do portal.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import time
import os
import re
import base64
import json
import socket
import shutil
import tempfile
import subprocess
from datetime import date, datetime

try:
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.chrome.options import Options
    from selenium.common.exceptions import TimeoutException
    SELENIUM_OK = True
except ImportError:
    SELENIUM_OK = False

try:
    from planilha import gerar_planilha
    OPENPYXL_OK = True
except ImportError:
    OPENPYXL_OK = False

from leitura import (DUIMP_PATTERN, dados_complementares, data_registro_no_texto, ncms_no_texto,
                     numero_br, numeros_duimp, peso_no_texto, primeiro_cnpj)


PORTAL_URL    = "https://portalunico.siscomex.gov.br"
SEARCH_URL    = PORTAL_URL + "/duimp/#/consultar-duimp"

# Possiveis localizacoes do Chrome no Windows
CHROME_PATHS = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
    os.path.expandvars(r"%PROGRAMFILES%\Google\Chrome\Application\chrome.exe"),
]

# Pastas padrao de saida, a partir da pasta do usuario, para funcionar em
# qualquer computador sem editar caminhos.
DEFAULT_OUT_DIR = os.path.join(os.path.expanduser("~"), "Desktop", "Robo DUIMP")
DEFAULT_PDF_DIR = os.path.join(DEFAULT_OUT_DIR, "PDFs")

# ---------------------------------------------------------------------------
class RoboDuimp:

    def __init__(self, root):
        self.root           = root
        self.driver         = None
        self._chrome_proc   = None
        self._temp_profile  = None
        self._stage_dir     = None   # pasta LOCAL temporaria de download
        self.running        = False
        self._login_event   = threading.Event()
        self._debug_port    = 9222

        root.title("Robo DUIMP - Portal Unico Siscomex v5.0")
        root.configure(bg="#1A3A5C")
        root.resizable(False, False)
        self._center(730, 710)

        self._build_ui()

        missing = []
        if not SELENIUM_OK:  missing.append("selenium")
        if not OPENPYXL_OK:  missing.append("openpyxl")
        if missing:
            messagebox.showerror("Dependencias faltando",
                "Instale com: pip install -r requirements.txt\n  " + ", ".join(missing))

    def _center(self, w, h):
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth()  - w) // 2
        y = (self.root.winfo_screenheight() - h) // 2
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    # Layout

    def _build_ui(self):
        hdr = tk.Frame(self.root, bg="#1A3A5C")
        hdr.pack(fill="x", padx=20, pady=(16, 4))
        tk.Label(hdr, text="ROBO DUIMP",
                 font=("Arial", 18, "bold"), fg="white", bg="#1A3A5C").pack(side="left")
        tk.Label(hdr, text="Portal Unico Siscomex v5.0",
                 font=("Arial", 10), fg="#90C8F0", bg="#1A3A5C").pack(side="left", padx=10, pady=(6,0))

        # Parametros
        pf = tk.LabelFrame(self.root, text="  Parametros de Busca  ",
                           font=("Arial", 10, "bold"), fg="white",
                           bg="#1A3A5C", bd=1, relief="groove")
        pf.pack(fill="x", padx=20, pady=4)
        hoje = date.today()
        self._param_row(pf, "CNPJ Importador:",  "cnpj",     "",                           0)
        self._param_row(pf, "Periodo - Inicio:", "data_ini", f"01/01/{hoje.year}",         1)
        self._param_row(pf, "Periodo - Fim:",    "data_fim", hoje.strftime("%d/%m/%Y"),    2)
        r3 = tk.Frame(pf, bg="#1A3A5C")
        r3.grid(row=3, column=0, sticky="w", padx=15, pady=6)
        tk.Label(r3, text="Pasta Planilha:", width=18, anchor="w",
                 fg="white", bg="#1A3A5C", font=("Arial", 10)).pack(side="left")
        self.out_var = tk.StringVar(value=DEFAULT_OUT_DIR)
        tk.Entry(r3, textvariable=self.out_var, width=36, font=("Arial", 10)).pack(side="left")
        tk.Button(r3, text="...", command=self._browse_out, width=3).pack(side="left", padx=4)

        # PDF
        pdf_lf = tk.LabelFrame(self.root, text="  Download de PDFs (Extrato DUIMP)  ",
                                font=("Arial", 10, "bold"), fg="white",
                                bg="#1A3A5C", bd=1, relief="groove")
        pdf_lf.pack(fill="x", padx=20, pady=4)
        self.pdf_check = tk.BooleanVar(value=False)
        ck = tk.Frame(pdf_lf, bg="#1A3A5C")
        ck.pack(fill="x", padx=15, pady=(8, 4))
        tk.Checkbutton(ck, text="Baixar PDF do extrato de cada DUIMP (via Gerar Extrato)",
                       variable=self.pdf_check, font=("Arial", 10), fg="white", bg="#1A3A5C",
                       selectcolor="#0D2137", activebackground="#1A3A5C",
                       activeforeground="white", command=self._toggle_pdf).pack(side="left")
        self.pdf_dir_frame = tk.Frame(pdf_lf, bg="#1A3A5C")
        tk.Label(self.pdf_dir_frame, text="Pasta dos PDFs:", width=18, anchor="w",
                 fg="white", bg="#1A3A5C", font=("Arial", 10)).pack(side="left")
        self.pdf_dir_var = tk.StringVar(value=DEFAULT_PDF_DIR)
        tk.Entry(self.pdf_dir_frame, textvariable=self.pdf_dir_var, width=36,
                 font=("Arial", 10)).pack(side="left")
        tk.Button(self.pdf_dir_frame, text="...",
                  command=self._browse_pdf, width=3).pack(side="left", padx=4)
        tk.Label(pdf_lf, text="", bg="#1A3A5C", height=1).pack()

        # Botao LOGIN (oculto ate necessario)
        self.btn_confirm_login = tk.Button(
            self.root,
            text="  LOGIN CONCLUIDO: CLIQUE AQUI PARA CONTINUAR  ",
            command=self._confirm_login,
            font=("Arial", 11, "bold"), bg="#E67E22", fg="white",
            activebackground="#CA6F1E", relief="flat", cursor="hand2", height=2
        )

        # Botoes acao
        bf = tk.Frame(self.root, bg="#1A3A5C")
        bf.pack(pady=8)
        self.btn_start = tk.Button(bf, text="  INICIAR EXTRACAO",
                                   command=self._start,
                                   font=("Arial", 12, "bold"), bg="#27AE60", fg="white",
                                   activebackground="#1E8449", width=22, height=2,
                                   relief="flat", cursor="hand2")
        self.btn_start.pack(side="left", padx=6)
        self.btn_stop = tk.Button(bf, text="  PARAR",
                                  command=self._stop,
                                  font=("Arial", 12, "bold"), bg="#C0392B", fg="white",
                                  width=12, height=2, relief="flat", cursor="hand2",
                                  state="disabled")
        self.btn_stop.pack(side="left", padx=6)

        self.lbl_prog = tk.Label(self.root, text="Aguardando inicio...",
                                 fg="#90C8F0", bg="#1A3A5C", font=("Arial", 10))
        self.lbl_prog.pack(anchor="w", padx=22)
        self.prog_var = tk.DoubleVar()
        ttk.Progressbar(self.root, variable=self.prog_var,
                        maximum=100, length=690).pack(padx=20, pady=(2, 4))

        lf = tk.LabelFrame(self.root, text="  Log de Execucao  ",
                            font=("Arial", 10, "bold"), fg="white",
                            bg="#1A3A5C", bd=1, relief="groove")
        lf.pack(fill="both", expand=True, padx=20, pady=(0, 14))
        self.log = tk.Text(lf, height=9, bg="#0D2137", fg="#E0E0E0",
                           font=("Consolas", 9), relief="flat", state="disabled")
        sb = ttk.Scrollbar(lf, command=self.log.yview)
        self.log.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.log.pack(fill="both", expand=True, padx=5, pady=5)
        self.log.tag_configure("ok",   foreground="#2ECC71")
        self.log.tag_configure("err",  foreground="#E74C3C")
        self.log.tag_configure("warn", foreground="#F39C12")
        self.log.tag_configure("info", foreground="#90C8F0")
        self.log.tag_configure("pdf",  foreground="#BB8FCE")

        self._log("Sistema pronto. Configure e clique em INICIAR.", "info")

    def _param_row(self, parent, label, attr, default, row):
        r = tk.Frame(parent, bg="#1A3A5C")
        r.grid(row=row, column=0, sticky="w", padx=15, pady=6)
        tk.Label(r, text=label, width=18, anchor="w",
                 fg="white", bg="#1A3A5C", font=("Arial", 10)).pack(side="left")
        var = tk.StringVar(value=default)
        setattr(self, attr + "_var", var)
        tk.Entry(r, textvariable=var, width=26, font=("Arial", 11)).pack(side="left")

    def _toggle_pdf(self):
        if self.pdf_check.get():
            self.pdf_dir_frame.pack(fill="x", padx=15, pady=(2, 8))
        else:
            self.pdf_dir_frame.pack_forget()

    # Helpers

    def _browse_out(self):
        d = filedialog.askdirectory(); 
        if d: self.out_var.set(d)

    def _browse_pdf(self):
        d = filedialog.askdirectory(); 
        if d: self.pdf_dir_var.set(d)

    def _log(self, msg, level="info"):
        def _do():
            self.log.config(state="normal")
            ts  = datetime.now().strftime("%H:%M:%S")
            sym = {"ok":"OK","err":"ERRO","warn":"AVISO","info":"-->","pdf":"PDF"}.get(level,"-->")
            self.log.insert("end", f"[{ts}] {sym:5s} {msg}\n", level)
            self.log.see("end")
            self.log.config(state="disabled")
        self.root.after(0, _do)

    def _set_progress(self, current, total, label=None):
        def _do():
            pct = (current / total * 100) if total else 0
            self.prog_var.set(pct)
            self.lbl_prog.config(
                text=label or f"Processando: {current}/{total}  ({pct:.0f}%)")
        self.root.after(0, _do)

    def _ui_done(self):
        self.root.after(0, lambda: self.btn_start.config(state="normal"))
        self.root.after(0, lambda: self.btn_stop.config(state="disabled"))
        self.root.after(0, lambda: self.btn_confirm_login.pack_forget())

    # Login manual

    def _confirm_login(self):
        self._login_event.set()
        self.root.after(0, lambda: self.btn_confirm_login.pack_forget())
        self._log("Login confirmado! Conectando ao Chrome...", "ok")

    def _show_login_button(self):
        self.btn_confirm_login.pack(fill="x", padx=20, pady=6, before=self.lbl_prog)

    def _wait_login(self, timeout=600):
        self._login_event.clear()
        self._log("-" * 52, "warn")
        self._log("ACAO NECESSARIA NO CHROME:", "warn")
        self._log("  1. Marque o checkbox 'Sou humano' (CAPTCHA)", "warn")
        self._log("  2. Clique em 'Entrar com certificado digital'", "warn")
        self._log("  3. Selecione o certificado e confirme", "warn")
        self._log("  4. Clique no botao LARANJA acima quando estiver", "warn")
        self._log("     dentro do portal (menu visivel)", "warn")
        self._log("-" * 52, "warn")
        self.root.after(0, self._show_login_button)

        deadline = time.time() + timeout
        while self.running and time.time() < deadline:
            if self._login_event.wait(timeout=1):
                return True
        if not self.running:
            return False
        self._log("Timeout aguardando login.", "err")
        return False

    # Control

    def _start(self):
        cnpj     = self.cnpj_var.get().strip()
        data_ini = self.data_ini_var.get().strip()
        data_fim = self.data_fim_var.get().strip()
        out_dir  = self.out_var.get().strip()

        if not re.match(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", cnpj):
            messagebox.showerror("Erro", "CNPJ invalido.\nFormato: XX.XXX.XXX/XXXX-XX"); return
        if not re.match(r"\d{2}/\d{2}/\d{4}", data_ini) or \
           not re.match(r"\d{2}/\d{2}/\d{4}", data_fim):
            messagebox.showerror("Erro", "Data invalida.\nFormato: DD/MM/AAAA"); return
        if not out_dir:
            messagebox.showerror("Erro", "Selecione a pasta de saida."); return

        pdf_dir = self.pdf_dir_var.get().strip() if self.pdf_check.get() else None
        os.makedirs(out_dir, exist_ok=True)
        if pdf_dir:
            os.makedirs(pdf_dir, exist_ok=True)

        self.running = True
        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.prog_var.set(0)
        self.lbl_prog.config(text="Iniciando...")

        threading.Thread(
            target=self._robot,
            args=(cnpj, data_ini, data_fim, out_dir, pdf_dir),
            daemon=True
        ).start()

    def _stop(self):
        self.running = False
        self._login_event.set()
        self._log("Parando...", "warn")
        self._cleanup_browser()
        self._ui_done()
        self.root.after(0, lambda: self.lbl_prog.config(text="Interrompido."))

    # Navegador

    def _find_chrome(self):
        """Localiza o executavel do Chrome."""
        for path in CHROME_PATHS:
            if os.path.exists(path):
                return path
        # Tentar via PATH
        for name in ("chrome", "google-chrome", "chromium"):
            found = shutil.which(name)
            if found:
                return found
        return None

    def _free_port(self):
        """Encontra uma porta livre para debug."""
        for port in range(9222, 9240):
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                try:
                    s.bind(("localhost", port))
                    return port
                except OSError:
                    continue
        return 9240

    def _open_chrome_normal(self, pdf_dir=None):
        """
        Abre o Chrome como um processo comum, num perfil temporario, para a pessoa
        fazer o login com o proprio certificado digital. O Selenium so se conecta
        depois, pela porta de depuracao.
        """
        chrome_exe = self._find_chrome()
        if not chrome_exe:
            raise Exception(
                "Google Chrome nao encontrado.\n"
                "Instale o Chrome em: https://www.google.com/chrome/")

        self._debug_port  = self._free_port()
        self._temp_profile = tempfile.mkdtemp(prefix="robo_duimp_")

        # O Chrome baixa para uma pasta LOCAL temporaria e o robo move o PDF
        # para a pasta final depois, assim funciona mesmo se a pasta final
        # estiver numa unidade de rede (ex: Z:) ou sem permissao de escrita
        # para o Chrome. A pasta tambem e gravada nas preferencias do perfil
        # (camada extra, independente do CDP).
        if pdf_dir:
            self._stage_dir = tempfile.mkdtemp(prefix="robo_duimp_dl_")
            try:
                pasta_default = os.path.join(self._temp_profile, "Default")
                os.makedirs(pasta_default, exist_ok=True)
                prefs = {
                    "download": {
                        "default_directory": self._stage_dir,
                        "prompt_for_download": False,
                        "directory_upgrade": True,
                    },
                    "savefile": {"default_directory": self._stage_dir},
                    "plugins": {"always_open_pdf_externally": True},
                    "profile": {"default_content_setting_values": {"automatic_downloads": 1}},
                }
                with open(os.path.join(pasta_default, "Preferences"), "w", encoding="utf-8") as f:
                    json.dump(prefs, f)
            except Exception as e:
                self._log(f"Aviso: nao foi possivel gravar preferencias de download: {e}", "warn")

        cmd = [
            chrome_exe,
            f"--remote-debugging-port={self._debug_port}",
            f"--user-data-dir={self._temp_profile}",
            "--no-first-run",
            "--no-default-browser-check",
            "--start-maximized",
            PORTAL_URL,
        ]
        self._chrome_proc = subprocess.Popen(cmd)
        self._log(f"Chrome aberto na porta de debug {self._debug_port}.", "ok")
        time.sleep(3)

    def _connect_selenium(self, pdf_dir=None):
        """
        Conecta o Selenium ao Chrome ja aberto, depois do login,
        pela porta de depuracao (debuggerAddress).
        """
        opts = Options()
        opts.add_experimental_option("debuggerAddress", f"localhost:{self._debug_port}")
        self.driver = webdriver.Chrome(options=opts)
        self.driver.set_script_timeout(600)
        self.driver.set_page_load_timeout(60)

        self._permitir_downloads(pdf_dir)
        self._log("Selenium conectado ao Chrome.", "ok")

    def _permitir_downloads(self, pdf_dir, verbose=False):
        """Autoriza downloads na aba/janela ATUAL e define a pasta de
        destino via CDP. O Chrome bloqueia downloads por padrao em paginas
        com o DevTools/Selenium anexado (confirmado: o mesmo fluxo manual
        funciona numa aba nova, sem o Selenium grudado, e falha na aba
        controlada), e essa autorizacao pode se perder quando a aba
        navega/troca de processo, entao e reaplicada logo antes de gerar o
        extrato. Com verbose=True o resultado vai para o log (antes as
        falhas eram engolidas em silencio). Retorna os comandos que
        funcionaram."""
        if not pdf_dir:
            return []
        destino = self._stage_dir or pdf_dir
        ok, erros = [], []
        for comando in ("Page.setDownloadBehavior", "Browser.setDownloadBehavior"):
            try:
                self.driver.execute_cdp_cmd(comando, {
                    "behavior":     "allow",
                    "downloadPath": destino,
                })
                ok.append(comando)
            except Exception as e:
                msg = str(e).strip().splitlines()[0] if str(e).strip() else type(e).__name__
                erros.append(f"{comando}: {msg}")
        if verbose:
            if ok:
                self._log(f"  Download liberado ({', '.join(ok)}).", "pdf")
            else:
                self._log("  Nao foi possivel liberar o download via CDP: "
                          + " | ".join(erros), "warn")
        return ok

    def _cleanup_browser(self):
        if self.driver:
            try: self.driver.quit()
            except: pass
            self.driver = None
        if self._chrome_proc:
            try: self._chrome_proc.terminate()
            except: pass
            self._chrome_proc = None
        if self._temp_profile:
            try: shutil.rmtree(self._temp_profile, ignore_errors=True)
            except: pass
            self._temp_profile = None
        if self._stage_dir:
            try: shutil.rmtree(self._stage_dir, ignore_errors=True)
            except: pass
            self._stage_dir = None

    # Robot

    def _robot(self, cnpj, data_ini, data_fim, out_dir, pdf_dir):
        duimps = []
        linhas = []
        try:
            # 1. Abrir o Chrome para o login
            self._log("Abrindo o Chrome para o login no portal...")
            self._open_chrome_normal(pdf_dir)

            # 2. Aguardar login MANUAL do usuario
            if not self._wait_login():
                return

            # 3. Conectar Selenium ao Chrome ja autenticado
            self._connect_selenium(pdf_dir)
            time.sleep(2)

            # 4. Busca
            self._log(f"Buscando DUIMPs, CNPJ {cnpj} | {data_ini} a {data_fim}")
            duimps = self._search(cnpj, data_ini, data_fim)
            if not self.running: return
            if not duimps:
                self._log("Nenhuma DUIMP encontrada.", "warn")
                messagebox.showwarning("Sem resultados",
                    "Nenhuma DUIMP encontrada.\nVerifique CNPJ e periodo.")
                return
            self._log(f"{len(duimps)} DUIMPs encontradas.", "ok")

            # 5. A coleta termina na ultima pagina de resultados, voltar
            # para a pagina 1 antes de processar (os DUIMPs sao visitados
            # na mesma ordem em que foram mapeados). Usa a paginacao da
            # propria grade, NAO reabre o formulario de busca, pois clicar
            # de novo em "Outros parametros" pode alternar (toggle) a
            # selecao e quebrar a consulta.
            self._ensure_single_window()
            self._ir_para_primeira_pagina()
            if not self.running: return

            # 6. Processar cada DUIMP: abrir, ler abas Carga/Item, gerar PDF
            total = len(duimps)
            self._log(f"Iniciando extracao ({total} DUIMPs)...")
            falhas = 0
            for i, (numero, data_registro) in enumerate(duimps):
                if not self.running: break
                self._set_progress(i, total, f"Processando: {i+1}/{total}, {numero}")
                self._log(f"[{i+1}/{total}] {numero}")
                try:
                    dados = self._processar_duimp(numero, data_registro, pdf_dir)
                    if dados:
                        linhas.append(dados)
                        self._log(
                            f"  Peso:{dados['peso']}  Qtd:{dados['qtd']}  "
                            f"CNPJ:{dados['cnpj']}  Ref:{dados['referencia']}", "ok")
                    else:
                        falhas += 1
                        self._log(f"  Nao foi possivel processar {numero}.", "warn")
                except Exception as e:
                    falhas += 1
                    self._log(f"  Erro em {numero}: {e}", "err")
                    self._ensure_single_window()

            if not self.running: return
            self._set_progress(total, total)
            if falhas:
                self._log(f"{falhas} DUIMP(s) com falha na extracao.", "warn")
            self._log("Extracao concluida!", "ok")

            # 7. Excel
            filename = ""
            if linhas:
                self._log("Gerando planilha Excel...")
                cnpj_d   = re.sub(r"\D", "", cnpj)
                filename = (f"DUIMP_{cnpj_d}_"
                            f"{data_ini.replace('/','')}_"
                            f"{data_fim.replace('/','')}.xlsx")
                out_path = os.path.join(out_dir, filename)
                gerar_planilha(linhas, out_path, cnpj, data_ini, data_fim)
                self._log(f"Planilha salva: {filename}", "ok")
                try: os.startfile(out_path)
                except: pass
            else:
                self._log("Nenhum dado coletado.", "err")

            pdf_info = f"\nPDFs: {pdf_dir}" if pdf_dir else ""
            self.root.after(0, lambda: messagebox.showinfo(
                "Concluido!",
                f"Extracao finalizada!\n\n"
                f"DUIMPs: {len(linhas)}\n"
                f"Planilha: {filename}\n"
                f"Pasta: {out_dir}" + pdf_info
            ))

        except Exception as e:
            import traceback
            self._log(f"Erro inesperado: {e}", "err")
            self._log(traceback.format_exc(), "err")
        finally:
            self.running = False
            self._ui_done()
            self.root.after(0, lambda: self.lbl_prog.config(text="Finalizado."))

    # Browser helpers

    def _type_masked_field(self, element_id, digits_only):
        """Preenche campo com ui-mask do AngularJS.
        Digita apenas os digitos, a mascara insere separadores automaticamente.
        Depois dispara change + $apply para sincronizar o ngModel."""
        from selenium.webdriver.common.keys import Keys
        try:
            el = self.driver.find_element(By.ID, element_id)
            self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
            el.click()
            time.sleep(0.2)
            # Posiciona no inicio e limpa
            el.send_keys(Keys.HOME)
            el.send_keys(Keys.SHIFT + Keys.END)
            el.send_keys(Keys.DELETE)
            time.sleep(0.15)
            # Digita caractere a caractere (mais confiavel com mascara)
            for ch in digits_only:
                el.send_keys(ch)
                time.sleep(0.05)
            time.sleep(0.2)
            # Dispara change e sincroniza AngularJS
            self.driver.execute_script("""
                try {
                    var el = document.getElementById(arguments[0]);
                    el.dispatchEvent(new Event('blur',   {bubbles:true}));
                    el.dispatchEvent(new Event('change', {bubbles:true}));
                    var $el  = angular.element(el);
                    var ctrl = $el.controller('ngModel');
                    if (ctrl) {
                        var scope = $el.scope();
                        if (scope && !scope.$$phase) scope.$apply();
                    }
                } catch(e) {}
            """, element_id)
            time.sleep(0.2)
            return True
        except Exception as e:
            self._log(f"  Erro ao preencher {element_id}: {e}", "warn")
            return False

    def _fill_search_form(self, cnpj, data_ini, data_fim):
        """Navega ate a tela de busca, preenche o formulario e clica Consultar.
        Nao coleta resultados, apenas dispara a pesquisa."""
        # Guardar params para reutilizar no fallback do contexto Angular
        self._s_cnpj     = cnpj
        self._s_data_ini = data_ini
        self._s_data_fim = data_fim

        self.driver.get(SEARCH_URL)

        # Aguardar a pagina carregar (esperar algum elemento do formulario aparecer)
        try:
            WebDriverWait(self.driver, 20).until(
                EC.presence_of_element_located((By.ID, "pucx-consultar")))
        except TimeoutException:
            self._log("  Aviso: botao Consultar nao apareceu em 20s", "warn")
        time.sleep(2)

        # Selecionar "Outros parametros"
        self._log("  Selecionando 'Outros parametros'...")
        outros_ok = self.driver.execute_script("""
            var all=document.querySelectorAll('input[type=radio],label,span,li');
            for(var i=0;i<all.length;i++){
                if((all[i].textContent||all[i].value||'').indexOf('Outros')>=0){
                    all[i].click();return true;
                }
            }
            return false;
        """)
        self._log(f"  'Outros parametros' selecionado: {outros_ok}")

        # Aguardar o campo CNPJ aparecer apos selecionar "Outros parametros"
        try:
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "inputCpfCnpj")))
            WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.ID, "inputCpfCnpj")))
        except TimeoutException:
            self._log("  Aviso: campo CNPJ nao ficou clicavel em 10s", "warn")
        time.sleep(0.5)

        # CNPJ, id="inputCpfCnpj", mascara ui-br-cpfcnpj-mask
        self._log("  Preenchendo CNPJ...")
        ok_cnpj = self._type_masked_field("inputCpfCnpj", re.sub(r'\D', '', cnpj))
        self._log(f"  CNPJ preenchido: {ok_cnpj}")

        # Data inicial, id="inputDataInicial", ui-mask="99/99/9999"
        self._log("  Preenchendo data inicial...")
        self._type_masked_field("inputDataInicial", re.sub(r'\D', '', data_ini))

        # Data final, id="inputDataFinal"
        self._log("  Preenchendo data final...")
        self._type_masked_field("inputDataFinal", re.sub(r'\D', '', data_fim))

        # Verificar valores lidos do DOM
        v_cnpj = self.driver.execute_script(
            "var e=document.getElementById('inputCpfCnpj'); return e?e.value:null;")
        v_ini  = self.driver.execute_script(
            "var e=document.getElementById('inputDataInicial'); return e?e.value:null;")
        v_fim  = self.driver.execute_script(
            "var e=document.getElementById('inputDataFinal'); return e?e.value:null;")
        self._log(f"  Valores, CNPJ:'{v_cnpj}'  ini:'{v_ini}'  fim:'{v_fim}'")

        # Se campos vazios, tentar novamente com espera maior
        if not v_cnpj or not v_ini or not v_fim:
            self._log("  Campos vazios, aguardando 3s e tentando novamente...", "warn")
            time.sleep(3)
            if not v_cnpj:
                self._type_masked_field("inputCpfCnpj", re.sub(r'\D', '', cnpj))
            if not v_ini:
                self._type_masked_field("inputDataInicial", re.sub(r'\D', '', data_ini))
            if not v_fim:
                self._type_masked_field("inputDataFinal", re.sub(r'\D', '', data_fim))

        time.sleep(0.5)

        # Consultar, id="pucx-consultar"
        self._log("  Clicando em Consultar...")
        try:
            btn = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.ID, "pucx-consultar")))
            self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
            btn.click()
        except Exception:
            self.driver.execute_script(
                "var b=document.getElementById('pucx-consultar');if(b)b.click();")
        time.sleep(6)

    def _search(self, cnpj, data_ini, data_fim):
        self._fill_search_form(cnpj, data_ini, data_fim)
        return self._collect_duimps()

    def _collect_duimps(self):
        """Mapeia numero + data de registro de cada DUIMP na listagem,
        percorrendo todas as paginas de resultado. Tenta ler as linhas da
        grade (numero + data juntos); se a estrutura da grade nao bater,
        cai no metodo comprovado de varrer o texto da pagina so pelo
        padrao do numero da DUIMP (data fica em branco e e recuperada na
        propria tela de detalhe)."""
        mapeados, seen, page = [], set(), 1
        while self.running:
            time.sleep(1.5)

            pares = []
            try:
                linhas_grid = self.driver.find_elements(By.XPATH, "//div[contains(@class,'ui-grid-row')]")
                for linha in linhas_grid:
                    try:
                        link = linha.find_element(By.XPATH, ".//a[contains(text(),'BR')]")
                        numero = link.text.strip()
                        if not DUIMP_PATTERN.match(numero):
                            continue
                        data_reg = ""
                        try:
                            celula = linha.find_element(
                                By.XPATH,
                                ".//div[contains(@class,'ui-grid-cell-contents') and "
                                "(contains(text(),'de 202') or contains(text(),'/'))]")
                            data_reg = celula.text.strip()
                        except Exception:
                            pass
                        pares.append((numero, data_reg))
                    except Exception:
                        continue
            except Exception:
                pares = []

            if pares:
                found = pares
            else:
                text = self.driver.execute_script("return document.body.innerText")
                found = [(n, "") for n in numeros_duimp(text)]

            novos = 0
            for numero, data_reg in found:
                if numero in seen:
                    continue
                seen.add(numero)
                mapeados.append((numero, data_reg))
                novos += 1

            self._log(f"  Pag.{page}: {novos} DUIMPs  (total:{len(mapeados)})")
            if novos == 0 and page == 1:
                time.sleep(4)
                text = self.driver.execute_script("return document.body.innerText")
                for numero in DUIMP_PATTERN.findall(text):
                    if numero not in seen:
                        seen.add(numero)
                        mapeados.append((numero, ""))
                break

            next_ok = self.driver.execute_script("""
                var c=document.querySelectorAll('button,a,li');
                for(var i=0;i<c.length;i++){
                    var t=(c[i].textContent||'').trim(),cls=c[i].className||'';
                    if(!c[i].disabled&&cls.indexOf('disabled')<0&&c[i].offsetParent!==null&&
                       (t==='>'||t.indexOf('roxima')>=0||cls.indexOf('next')>=0)){c[i].click();return true;}
                }
                return false;
            """)
            if not next_ok: break
            page += 1; time.sleep(2)
        return mapeados

    def _ir_para_primeira_pagina(self):
        """Volta para a pagina 1 da listagem usando a paginacao da propria
        grade (mesmo mecanismo que ja funciona em `_collect_duimps`). NAO
        reabre o formulario de busca, clicar de novo em 'Outros parametros'
        pode alternar (toggle) a selecao e deixar a consulta inutilizavel."""
        try:
            clicou = self.driver.execute_script("""
                var el = document.querySelector("button[ng-click='pageFirstPageClick()']");
                if (el && !el.disabled) { el.click(); return true; }
                return false;
            """)
            if clicou:
                time.sleep(3)
                return
        except Exception:
            pass

        for _ in range(20):
            if not self.running:
                return
            voltou = self.driver.execute_script("""
                var c=document.querySelectorAll('button,a,li');
                for(var i=0;i<c.length;i++){
                    var t=(c[i].textContent||'').trim(),cls=c[i].className||'';
                    if(!c[i].disabled&&cls.indexOf('disabled')<0&&c[i].offsetParent!==null&&
                       (t==='<'||t.indexOf('nterior')>=0||cls.indexOf('previous')>=0||cls.indexOf('prev')>=0)){
                        c[i].click();return true;
                    }
                }
                return false;
            """)
            if not voltou:
                break
            time.sleep(1.5)

    def _localizar_link_duimp(self, numero, tentativas=8):
        """Localiza o link da DUIMP na pagina atual da listagem. Se nao
        estiver visivel, avanca paginas ate achar (os processos sao
        visitados na mesma ordem em que foram mapeados, entao so e
        preciso avancar, nunca retroceder)."""
        for _ in range(tentativas):
            if not self.running:
                return None
            try:
                link = self.driver.find_element(By.XPATH, f"//a[contains(text(),'{numero}')]")
                if link.is_displayed():
                    return link
            except Exception:
                pass
            avancou = self.driver.execute_script("""
                var c=document.querySelectorAll('button,a,li');
                for(var i=0;i<c.length;i++){
                    var t=(c[i].textContent||'').trim(),cls=c[i].className||'';
                    if(!c[i].disabled&&cls.indexOf('disabled')<0&&c[i].offsetParent!==null&&
                       (t==='>'||t.indexOf('roxima')>=0||cls.indexOf('next')>=0)){c[i].click();return true;}
                }
                return false;
            """)
            if not avancou:
                break
            time.sleep(2)
        return None

    def _fechar_aba_detalhe(self, aba_principal):
        try:
            if len(self.driver.window_handles) > 1:
                self.driver.close()
                self.driver.switch_to.window(aba_principal)
            else:
                self.driver.switch_to.window(aba_principal)
        except Exception:
            self._ensure_single_window()
        time.sleep(1.5)

    def _extrair_data_registro(self):
        """Fallback: le a data de registro direto da tela de detalhe
        quando a listagem nao forneceu (grade com estrutura diferente)."""
        try:
            return data_registro_no_texto(self.driver.execute_script("return document.body.innerText"))
        except Exception:
            return ""

    def _clicar_aba(self, nome_aba, tentativas=2, timeout=25):
        """Clica na aba (Carga/Item) pelo texto do titulo. Tenta mais de
        uma vez porque algumas DUIMPs demoram a renderizar. Retorna True/
        False e registra um motivo legivel (o TimeoutException do Selenium
        vem com mensagem vazia e so a pilha do chromedriver)."""
        xpath = ("//span[contains(@class,'p-tabview-title') and "
                 f"normalize-space(.)='{nome_aba}']/ancestor::a")
        for tentativa in range(1, tentativas + 1):
            try:
                btn = WebDriverWait(self.driver, timeout).until(
                    EC.element_to_be_clickable((By.XPATH, xpath)))
                self.driver.execute_script("arguments[0].click();", btn)
                return True
            except TimeoutException:
                if tentativa < tentativas:
                    time.sleep(2)
            except Exception as e:
                self._log(f"  Erro ao clicar na aba {nome_aba}: {type(e).__name__}", "warn")
                if tentativa < tentativas:
                    time.sleep(2)
        self._log(f"  Aba '{nome_aba}' nao ficou clicavel em {timeout}s "
                  f"({tentativas} tentativas).", "warn")
        return False

    def _extrair_peso_bruto(self, wait):
        """Aba 'Carga', peso bruto total (kg). Tenta pelo XPath do campo e,
        se nao achar, le o texto da tela procurando 'Peso bruto'. Se falhar,
        reabre a aba e tenta mais uma vez antes de desistir."""
        xpath_peso = ("//div[contains(., 'Peso bruto (kg)')]/following-sibling::div//p | "
                      "//*[contains(text(), 'Peso bruto')]/following::p[1]")
        for tentativa in (1, 2):
            if not self._clicar_aba("Carga"):
                return ""
            time.sleep(2.5)
            try:
                el = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, xpath_peso)))
                txt = el.text.strip()
                if txt:
                    return txt
            except TimeoutException:
                pass
            except Exception as e:
                self._log(f"  Erro ao ler peso bruto: {type(e).__name__}", "warn")
            # Plano B: ler o texto visivel da tela
            try:
                peso = peso_no_texto(self.driver.execute_script("return document.body.innerText"))
                if peso:
                    return peso
            except Exception:
                pass
            if tentativa == 1:
                time.sleep(2)
        self._log("  Peso bruto nao encontrado na aba Carga (campo nao apareceu).", "warn")
        return ""

    def _extrair_ncms(self):
        """NCMs visiveis na aba Item."""
        try:
            return ncms_no_texto(self.driver.execute_script("return document.body.innerText"))
        except Exception:
            return []

    def _extrair_qtd_n_itens(self):
        """Le a tabela de itens e retorna (soma da quantidade comercializada,
        ou vazio se a coluna nao existir na grade resumida; numero de
        linhas/itens, usado para a faixa do PDF)."""
        result = self.driver.execute_script(r"""
            function colIndex(table, matchers) {
                var ths = table.querySelectorAll('thead th');
                for (var i = 0; i < ths.length; i++) {
                    var id = (ths[i].id || '').toLowerCase();
                    var txt = (ths[i].textContent || '').trim().toLowerCase();
                    for (var m = 0; m < matchers.length; m++) {
                        if (id.indexOf(matchers[m]) >= 0 || txt.indexOf(matchers[m]) >= 0) return i;
                    }
                }
                return -1;
            }
            var tables = document.querySelectorAll('.p-datatable, table');
            var qtd = [], rows = 0;
            for (var t = 0; t < tables.length; t++) {
                var table = tables[t];
                var itemIdx = colIndex(table, ['numeroitem', 'nº do item', 'n° do item', 'do item']);
                var qtdIdx = colIndex(table, ['quantidadecomercial', 'quantidade comercial', 'quantidade']);
                if (itemIdx < 0 && qtdIdx < 0) continue;
                var trs = table.querySelectorAll('tbody tr');
                for (var r = 0; r < trs.length; r++) {
                    var cells = trs[r].querySelectorAll('td');
                    if (qtdIdx >= 0 && cells[qtdIdx]) qtd.push(cells[qtdIdx].textContent.trim());
                }
                rows = trs.length;
                if (rows) break;
            }
            return {qtd: qtd, rows: rows};
        """) or {}

        qtd_raw = result.get("qtd", [])
        total = sum(numero_br(q) for q in qtd_raw) if qtd_raw else ""
        n_itens = int(result.get("rows") or 1)
        return total, n_itens

    def _extrair_item_dados(self, wait):
        """Aba 'Item': CNPJ e nome do adquirente, NCM, quantidade e referencia
        interna. Retorna (cnpj, nome, ncm_list, qtd, referencia, n_itens)."""
        cnpj, nome, referencia = "", "", ""
        ncm_list, qtd_total, n_itens = [], "", 1
        if not self._clicar_aba("Item"):
            return cnpj, nome, ncm_list, qtd_total, referencia, n_itens
        time.sleep(4)

        # CNPJ do adquirente (coluna da tabela, isolada por id do cabecalho)
        try:
            el_cnpj = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((
                By.XPATH,
                "//th[@id='identificacaoAdquirente.codigo']/ancestor::div[contains(@class,'p-datatable')]"
                "//tbody/tr[1]//td[count(//th[@id='identificacaoAdquirente.codigo']/preceding-sibling::th)+1]//span")))
            cnpj = el_cnpj.text.strip()
        except Exception:
            try:
                el_alt = self.driver.find_element(
                    By.XPATH,
                    "//*[@id='identificacaoAdquirente.codigo']/ancestor::table//tbody/tr[1]"
                    "//span[contains(@class,'p-element') and contains(text(),'/')]")
                cnpj = el_alt.text.strip()
            except Exception:
                cnpj = primeiro_cnpj(self.driver.page_source)

        # Nome do adquirente e referencia interna (informacoes complementares)
        try:
            paragrafo = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((
                By.XPATH,
                "//p[contains(@style,'white-space: pre-wrap') or contains(text(),'ENCOMENDANTE') "
                "or contains(text(),'IMPORTADOR') or contains(text(),'ADQUIRENTE')]")))
            nome, referencia = dados_complementares(paragrafo.text)
        except Exception as e:
            self._log(f"  Nao foi possivel ler informacoes complementares: {e}", "warn")

        try:
            ncm_list = self._extrair_ncms()
        except Exception as e:
            self._log(f"  Nao foi possivel ler NCM: {e}", "warn")

        try:
            qtd_total, n_itens = self._extrair_qtd_n_itens()
        except Exception as e:
            self._log(f"  Nao foi possivel ler quantidade/itens: {e}", "warn")

        return cnpj, nome, ncm_list, qtd_total, referencia, n_itens

    def _processar_duimp(self, numero, data_registro, pdf_dir):
        """Abre a DUIMP, le as abas Carga/Item e (se solicitado) gera o
        PDF do extrato, tudo na mesma visita, sem depender de internals
        do Angular."""
        wait = WebDriverWait(self.driver, 25)
        aba_principal = self.driver.current_window_handle

        link = self._localizar_link_duimp(numero)
        if not link:
            self._log(f"  Link da DUIMP {numero} nao encontrado na listagem.", "warn")
            return None

        self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", link)
        time.sleep(0.3)
        try:
            # Clique nativo do Selenium, tratado pelo Chrome como gesto do
            # usuario. Um clique disparado via JS (element.click() por
            # execute_script) nao e confiavel para o navegador e o
            # bloqueador de pop-up pode descartar o window.open() que o
            # Angular dispara ao abrir a aba de detalhe.
            link.click()
        except Exception:
            try:
                self.driver.execute_script("arguments[0].click();", link)
            except Exception as e:
                self._log(f"  Erro ao clicar em {numero}: {e}", "warn")
                return None
        time.sleep(4)

        for aba in self.driver.window_handles:
            if aba != aba_principal:
                self.driver.switch_to.window(aba)
                break

        # O Chrome bloqueia downloads por padrao em abas com o Selenium
        # anexado, e essa permissao NAO se propaga sozinha para abas novas.
        # Precisa ser reforcada aqui, na aba de detalhe recem-aberta.
        self._permitir_downloads(pdf_dir)

        try:
            self._log(f"  URL da aba: {self.driver.current_url}")
        except Exception:
            pass

        try:
            wait.until(EC.presence_of_element_located((
                By.XPATH,
                "//span[contains(@class,'p-tabview-title')] | "
                "//li[contains(@class,'p-tabview')]//a | "
                "//*[@role='tab']")))
        except TimeoutException:
            achou_carga_item = False
            try:
                texto = self.driver.execute_script("return document.body.innerText") or ""
                achou_carga_item = ("Carga" in texto) or ("Item" in texto)
            except Exception:
                pass
            self._log(
                f"  Detalhe da DUIMP nao carregou "
                f"(texto Carga/Item presente na pagina: {achou_carga_item}).", "warn")
            self._fechar_aba_detalhe(aba_principal)
            return None

        if not data_registro:
            data_registro = self._extrair_data_registro()

        peso_txt = self._extrair_peso_bruto(wait)
        peso = numero_br(peso_txt) if peso_txt else ""
        cnpj, nome, ncm_list, qtd, referencia, n_itens = self._extrair_item_dados(wait)

        if pdf_dir:
            self._gerar_extrato_pdf_atual(numero, pdf_dir, n_itens)

        self._fechar_aba_detalhe(aba_principal)

        return {
            "numero": numero,
            "data_registro": data_registro,
            "qtd": qtd,
            "peso": peso,
            "ncm": ", ".join(ncm_list),
            "cnpj": cnpj,
            "nome": nome,
            "referencia": referencia,
        }

    def _gerar_extrato_pdf_atual(self, numero, pdf_dir, n_itens=1):
        """Gera e baixa o PDF do extrato a partir da aba 'Item' ja aberta
        (nao reabre a busca, reusa a mesma visita a DUIMP). O botao 'Gerar
        Extrato' existe duas vezes na tela (o que abre o dialog e o que
        confirma dentro dele, com o MESMO texto), o primeiro clique marca
        o botao usado para que a busca do botao de confirmacao o ignore."""
        # Reaplica a liberacao de download NESTA aba imediatamente antes do
        # clique (a aba pode ter navegado/trocado de processo desde que foi
        # liberada em _processar_duimp, perdendo a permissao).
        self._permitir_downloads(pdf_dir, verbose=True)

        gerar_ok = self.driver.execute_script("""
            var spans=document.querySelectorAll('.p-button-label,.p-button span,button span');
            for(var i=0;i<spans.length;i++){
                if(spans[i].textContent.trim()==='Gerar Extrato'){
                    var btn=spans[i].closest('button')||spans[i].parentElement;
                    if(btn&&!btn.disabled&&btn.offsetParent!==null){
                        btn.setAttribute('data-robo-abriu','1');
                        btn.click();
                        return true;
                    }
                }
            }
            return false;
        """)
        if not gerar_ok:
            self._log("  Botao Gerar Extrato nao encontrado.", "warn")
            return False
        time.sleep(2)

        # Campo da faixa de itens tem id fixo: <input id="faixa" formcontrolname="faixa">
        n_str = f"1-{max(int(n_itens or 1), 1)}"
        preenchido = self.driver.execute_script("""
            var v = arguments[0];
            var el = document.getElementById('faixa') || document.querySelector('[formcontrolname="faixa"]');
            if (!el || el.offsetParent===null) return false;
            var setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
            setter.call(el, v);
            el.dispatchEvent(new Event('input', {bubbles:true}));
            el.dispatchEvent(new Event('change', {bubbles:true}));
            el.dispatchEvent(new Event('blur', {bubbles:true}));
            return true;
        """, n_str)
        if not preenchido:
            self._log("  Campo #faixa (itens) nao encontrado ou nao visivel.", "warn")
            return False
        time.sleep(0.8)

        confirm_ok = False
        for _ in range(6):
            confirm_ok = self.driver.execute_script("""
                var btns=document.querySelectorAll('button');
                for(var i=0;i<btns.length;i++){
                    var b=btns[i];
                    if(b.hasAttribute('data-robo-abriu')) continue;
                    var txt=(b.textContent||'').trim();
                    if(txt.indexOf('Gerar Extrato')>=0 && !b.disabled && b.offsetParent!==null){
                        b.click();return true;
                    }
                }
                return false;
            """)
            if confirm_ok:
                break
            time.sleep(0.5)
        if not confirm_ok:
            self._log("  Nao confirmou geracao do extrato.", "warn")
            return False
        time.sleep(1)

        # A partir daqui o extrato pode chegar de duas formas: (a) um
        # download real na pasta configurada, ou (b) uma nova aba abrindo
        # com o PDF (visualizador do Chrome). Espera por QUALQUER um dos
        # dois sinais, nunca assume sucesso sem um deles.
        aba_detalhe = self.driver.current_window_handle
        handles_antes = set(self.driver.window_handles)
        safe = numero.replace("/", "-").replace("\\", "-")
        dest = os.path.join(pdf_dir, f"{safe}.pdf")
        digitos_num = numero.replace("-", "")

        # Pastas vigiadas: a pasta temporaria local (para onde mandamos o
        # Chrome baixar), a pasta final escolhida e a pasta Downloads padrao
        # (caso a liberacao via CDP falhe e o Chrome use o padrao dele; ali
        # so valem arquivos cujo nome contem o numero desta DUIMP).
        downloads_padrao = os.path.join(os.path.expanduser("~"), "Downloads")
        pastas = []
        for p in (self._stage_dir, pdf_dir, downloads_padrao):
            if p and p not in pastas and os.path.isdir(p):
                pastas.append(p)
        antes = {p: self._list_pdfs(p) for p in pastas}

        aba_pdf = None
        deadline = time.time() + 90
        while time.time() < deadline:
            time.sleep(1.5)
            for p in pastas:
                try:
                    arquivos = os.listdir(p)
                except Exception:
                    continue
                if any(f.lower().endswith((".crdownload", ".tmp")) for f in arquivos):
                    continue  # ainda baixando
                novos = self._list_pdfs(p) - antes[p]
                if p == downloads_padrao:
                    novos = {f for f in novos if digitos_num in os.path.basename(f).replace("-", "")}
                if not novos:
                    continue
                origem = sorted(novos, key=os.path.getmtime)[-1]
                if self._mover_pdf(origem, dest):
                    self._log(f"  PDF salvo: {safe}.pdf", "pdf")
                    return True
            try:
                extras = set(self.driver.window_handles) - handles_antes
            except Exception:
                extras = set()
            if extras:
                aba_pdf = list(extras)[0]
                break

        if not aba_pdf:
            detalhe = []
            for p in pastas:
                try:
                    detalhe.append(f"{p}: {len(os.listdir(p))} arquivo(s)")
                except Exception:
                    pass
            self._log("  Timeout: nenhum download nem aba de PDF detectados "
                      f"(pastas vigiadas -> {' | '.join(detalhe)}).", "warn")
            return False

        # Aba nova apareceu, trocar para ela e salvar o PDF de verdade
        # (nunca imprimir a aba Item por engano).
        salvo = False
        try:
            self.driver.switch_to.window(aba_pdf)
            url_pdf = self.driver.current_url
        except Exception:
            url_pdf = ""

        if url_pdf:
            try:
                b64 = self.driver.execute_async_script("""
                    var url = arguments[0];
                    var callback = arguments[arguments.length - 1];
                    fetch(url).then(function(r){
                        if (url.indexOf('blob:') === 0) return r.blob();
                        var ct = (r.headers.get('content-type')||'').toLowerCase();
                        if (ct.indexOf('pdf') === -1) return null;
                        return r.blob();
                    }).then(function(blob){
                        if (!blob) { callback(null); return; }
                        var reader = new FileReader();
                        reader.onloadend = function(){ callback(reader.result.split(',')[1]); };
                        reader.onerror = function(){ callback(null); };
                        reader.readAsDataURL(blob);
                    }).catch(function(){ callback(null); });
                """, url_pdf)
                if b64:
                    with open(dest, "wb") as f:
                        f.write(base64.b64decode(b64))
                    self._log(f"  PDF salvo: {safe}.pdf", "pdf")
                    salvo = True
            except Exception as e:
                self._log(f"  Falha ao baixar bytes do PDF: {e}", "warn")

        if not salvo:
            # Fallback: imprime a PROPRIA aba do PDF (ja trocamos para ela
            # acima), nunca a aba Item.
            if self._cdp_print_to_pdf(dest):
                self._log(f"  PDF salvo via impressao da aba do extrato: {safe}.pdf", "pdf")
                salvo = True

        try:
            if len(self.driver.window_handles) > 1:
                self.driver.close()
            self.driver.switch_to.window(aba_detalhe)
        except Exception:
            self._ensure_single_window()

        if not salvo:
            self._log("  Nao foi possivel salvar o PDF do extrato.", "warn")
        return salvo

    def _ensure_single_window(self):
        """Fecha abas extras e garante que o driver esta anexado a uma janela
        valida. Evita que comandos falhem apos a aba de detalhe da DUIMP ser
        fechada/invalidada pelo portal (causa do erro 'Fallback falhou' com
        stacktrace vazio do chromedriver)."""
        try:
            handles = self.driver.window_handles
        except Exception:
            return
        if len(handles) > 1:
            main = handles[0]
            for h in handles[1:]:
                try:
                    self.driver.switch_to.window(h)
                    self.driver.close()
                except Exception:
                    pass
            try:
                self.driver.switch_to.window(main)
            except Exception:
                try:
                    self.driver.switch_to.window(self.driver.window_handles[0])
                except Exception:
                    pass
        elif len(handles) == 1:
            try:
                self.driver.switch_to.window(handles[0])
            except Exception:
                pass

    def _mover_pdf(self, origem, destino):
        """Move o PDF baixado para o destino final (pode ser unidade de
        rede). Espera o tamanho estabilizar e tenta de novo se o arquivo
        ainda estiver travado pelo Chrome."""
        try:
            tam = -1
            for _ in range(10):
                atual = os.path.getsize(origem)
                if atual > 0 and atual == tam:
                    break
                tam = atual
                time.sleep(1)
        except Exception:
            pass
        ultimo_erro = None
        for _ in range(8):
            try:
                if os.path.exists(destino):
                    os.remove(destino)
                shutil.move(origem, destino)
                return True
            except Exception as e:
                ultimo_erro = e
                time.sleep(1)
        self._log(f"  PDF baixado, mas nao foi possivel mover para a pasta final: {ultimo_erro}", "warn")
        return False

    def _list_pdfs(self, folder):
        try:
            return {os.path.join(folder, f)
                    for f in os.listdir(folder) if f.lower().endswith(".pdf")}
        except Exception:
            return set()

    def _cdp_print_to_pdf(self, pdf_path):
        try:
            result = self.driver.execute_cdp_cmd("Page.printToPDF", {
                "printBackground": True, "paperWidth": 8.27, "paperHeight": 11.69,
                "marginTop": 0.4, "marginBottom": 0.4,
                "marginLeft": 0.4, "marginRight": 0.4,
                "scale": 0.85, "displayHeaderFooter": True,
                "headerTemplate": "<div style='font-size:9px;color:#666;width:100%;text-align:center'>"
                                  "Portal Unico Siscomex - DUIMP</div>",
                "footerTemplate": "<div style='font-size:8px;color:#666;width:100%;text-align:center'>"
                                  "Pag <span class='pageNumber'></span>/<span class='totalPages'></span></div>",
            })
            with open(pdf_path, "wb") as f:
                f.write(base64.b64decode(result["data"]))
            return True
        except Exception as e:
            self._log(f"  CDP erro: {e}", "warn")
            return False


# ---------------------------------------------------------------------------
if __name__ == "__main__":
    root = tk.Tk()
    RoboDuimp(root)
    root.mainloop()
