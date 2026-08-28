import asyncio
import io
import os
import re
import sqlite3
import sys
import threading
import time
from collections import deque
from datetime import datetime
from pathlib import Path
from urllib.parse import quote_plus
import warnings

# --- CORREÇÃO DE NAVEGADORES PLAYWRIGHT EM EXECUTÁVEIS (.EXE) ---
if getattr(sys, 'frozen', False):
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = os.path.join(
        os.environ.get("LOCALAPPDATA", ""), "ms-playwright"
    )
else:
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = "0"

# Silencia avisos de DeprecationWarning do Python 3.14
warnings.filterwarnings("ignore", category=DeprecationWarning)

import customtkinter as ctk
import httpx
from bs4 import BeautifulSoup
import openpyxl
from PIL import Image
from playwright.async_api import async_playwright
import psutil

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class LeadHunterApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("LeadHunter - Prospecção Automática de Leads")
        self.geometry("1100x750")
        self.minsize(950, 650)

        # ------------------ GERENCIAMENTO DE PASTAS LOCAIS (.EXE / .PY) ------------------
        if getattr(sys, 'frozen', False):
            # Garante a criação na pasta real do .exe (fora da pasta temporária _MEIxxxx)
            self.base_dir = Path(sys.executable).parent.resolve()
        else:
            self.base_dir = Path(__file__).parent.resolve() if "__file__" in globals() else Path.cwd()

        self.logs_dir = self.base_dir / "logs"
        self.relatorios_dir = self.base_dir / "relatorios"

        # Cria as pastas "logs" e "relatorios" no diretório local
        self.logs_dir.mkdir(parents=True, exist_ok=True)
        self.relatorios_dir.mkdir(parents=True, exist_ok=True)

        # Inicializa arquivo de log da sessão atual
        session_timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        self.log_file_path = self.logs_dir / f"log_{session_timestamp}.txt"
        
        # Variáveis de Controle
        self.is_running = False
        self.allocated_ram_gb = ctk.DoubleVar(value=2.0)
        self.ram_buffer = deque()
        self.db_path = self.base_dir / "leadhunter_storage.db"

        self.init_db()
        self.build_ui()
        self.start_ram_monitor()

        self.log_message(f"[SISTEMA] Diretório local de execução: {self.base_dir}")
        self.log_message(f"[SISTEMA] Pasta de Logs pronta: {self.logs_dir}")
        self.log_message(f"[SISTEMA] Pasta de Relatórios pronta: {self.relatorios_dir}")

    def init_db(self):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS leads (
                id TEXT PRIMARY KEY,
                nome TEXT,
                telefone TEXT,
                email TEXT,
                site TEXT,
                whatsapp TEXT,
                instagram TEXT,
                facebook TEXT,
                linkedin TEXT,
                nicho TEXT,
                cidade TEXT,
                timestamp TEXT
            )
        """)
        conn.commit()
        conn.close()

    def is_lead_processed(self, lead_id):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT 1 FROM leads WHERE id = ?", (lead_id,))
        exists = cursor.fetchone() is not None
        conn.close()
        return exists

    def build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(3, weight=2)

        # Inputs
        self.frame_inputs = ctk.CTkFrame(self, corner_radius=10)
        self.frame_inputs.grid(row=0, column=0, padx=15, pady=15, sticky="nsew")

        ctk.CTkLabel(self.frame_inputs, text="Filtros de Prospecção", font=("Segoe UI", 16, "bold")).pack(anchor="w", padx=15, pady=(12, 5))

        self.entry_nicho = ctk.CTkEntry(self.frame_inputs, placeholder_text="Nicho (Ex: Odontologia)")
        self.entry_nicho.pack(fill="x", padx=15, pady=5)

        self.entry_cidade = ctk.CTkEntry(self.frame_inputs, placeholder_text="Cidades (Ex: Artur Nogueira, Campinas, Sumaré)")
        self.entry_cidade.pack(fill="x", padx=15, pady=5)

        ctk.CTkLabel(self.frame_inputs, text="Alocação de Cache RAM (Heap Size)", font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=15, pady=(10, 2))
        
        self.slider_ram = ctk.CTkSlider(self.frame_inputs, from_=0.5, to=16.0, number_of_steps=31, variable=self.allocated_ram_gb, command=self.update_ram_label)
        self.slider_ram.pack(fill="x", padx=15, pady=2)

        self.lbl_ram_allocated = ctk.CTkLabel(self.frame_inputs, text=f"RAM Alocada para Cache: {self.allocated_ram_gb.get():.1f} GB")
        self.lbl_ram_allocated.pack(anchor="w", padx=15, pady=(0, 5))

        # Mini Player
        self.frame_player = ctk.CTkFrame(self, corner_radius=10)
        self.frame_player.grid(row=0, column=1, padx=15, pady=15, sticky="nsew")

        ctk.CTkLabel(self.frame_player, text="Mini Player - Monitoramento em Tempo Real", font=("Segoe UI", 14, "bold")).pack(anchor="w", padx=10, pady=8)

        self.lbl_mini_player = ctk.CTkLabel(self.frame_player, text="Navegador em Espera", width=360, height=200, fg_color="#1e293b", corner_radius=8)
        self.lbl_mini_player.pack(expand=True, fill="both", padx=10, pady=10)

        # Status RAM
        self.frame_ram_bar = ctk.CTkFrame(self, corner_radius=8)
        self.frame_ram_bar.grid(row=1, column=0, columnspan=2, padx=15, pady=(0, 10), sticky="ew")

        self.lbl_ram_status = ctk.CTkLabel(self.frame_ram_bar, text="Uso de Memória RAM do Sistema: 0%", font=("Segoe UI", 11, "bold"))
        self.lbl_ram_status.pack(anchor="w", padx=15, pady=(5, 2))

        self.progress_ram = ctk.CTkProgressBar(self.frame_ram_bar, height=12)
        self.progress_ram.set(0.0)
        self.progress_ram.pack(fill="x", padx=15, pady=(0, 8))

        # Controles
        self.frame_controls = ctk.CTkFrame(self, corner_radius=8)
        self.frame_controls.grid(row=2, column=0, columnspan=2, padx=15, pady=5, sticky="ew")

        self.btn_start = ctk.CTkButton(self.frame_controls, text="Iniciar Varredura", command=self.toggle_scraping, fg_color="#10b981", hover_color="#059669")
        self.btn_start.pack(side="left", padx=15, pady=10)

        self.lbl_progress_status = ctk.CTkLabel(self.frame_controls, text="Status: Aguardando início...", font=("Segoe UI", 12, "bold"))
        self.lbl_progress_status.pack(side="left", padx=15)

        self.progress_search = ctk.CTkProgressBar(self.frame_controls, width=280)
        self.progress_search.set(0.0)
        self.progress_search.pack(side="right", padx=15, pady=10)

        # Log Textbox
        self.textbox_log = ctk.CTkTextbox(self, corner_radius=10, font=("Consolas", 11))
        self.textbox_log.grid(row=3, column=0, columnspan=2, padx=15, pady=(5, 15), sticky="nsew")

    def log_message(self, text):
        self.textbox_log.insert("end", text + "\n")
        self.textbox_log.see("end")

        try:
            with open(self.log_file_path, "a", encoding="utf-8") as f:
                f.write(text + "\n")
        except Exception:
            pass

    def update_status(self, text):
        self.lbl_progress_status.configure(text=f"Status: {text}")

    def update_ram_label(self, value):
        self.lbl_ram_allocated.configure(text=f"RAM Alocada para Cache: {value:.1f} GB")

    def start_ram_monitor(self):
        def monitor_loop():
            while True:
                mem = psutil.virtual_memory()
                used_pct = mem.percent / 100.0

                if used_pct < 0.65:
                    bar_color = "#34d399"
                elif used_pct < 0.85:
                    bar_color = "#fbbf24"
                else:
                    bar_color = "#f87171"

                self.progress_ram.configure(progress_color=bar_color)
                self.progress_ram.set(used_pct)
                self.lbl_ram_status.configure(text=f"Uso de Memória RAM do Sistema: {mem.percent}% ({mem.used / (1024**3):.1f}GB / {mem.total / (1024**3):.1f}GB)")
                time.sleep(1.5)

        threading.Thread(target=monitor_loop, daemon=True).start()

    def toggle_scraping(self):
        if not self.is_running:
            nicho = self.entry_nicho.get().strip()
            cidades_raw = self.entry_cidade.get().strip()

            if not nicho or not cidades_raw:
                self.log_message("[ALERTA] Preencha os campos de Nicho e Cidade antes de iniciar!")
                return

            lista_cidades = [c.strip() for c in cidades_raw.split(",") if c.strip()]

            self.textbox_log.delete("1.0", "end")

            self.is_running = True
            self.btn_start.configure(text="Parar Varredura", fg_color="#ef4444", hover_color="#dc2626")
            self.update_status("Iniciando motor de busca...")
            self.log_message(f"[{datetime.now().strftime('%H:%M:%S')}] Iniciando lote de {len(lista_cidades)} cidade(s) para o nicho '{nicho}'...")

            threading.Thread(target=self.run_async_batch_scraper, args=(nicho, lista_cidades), daemon=True).start()
        else:
            self.is_running = False
            self.btn_start.configure(text="Iniciar Varredura", fg_color="#10b981", hover_color="#059669")
            self.update_status("Interrompido pelo usuário.")
            self.progress_search.set(0.0)
            self.log_message(f"[{datetime.now().strftime('%H:%M:%S')}] Interrompendo varredura pelo usuário...")

    def run_async_batch_scraper(self, nicho, lista_cidades):
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        try:
            loop.run_until_complete(self.batch_scraper_worker(nicho, lista_cidades))
        except Exception as e:
            self.log_message(f"[ERRO DE EXECUÇÃO] {str(e)}")
            self.update_status("Erro durante execução.")
        finally:
            loop.close()

    async def batch_scraper_worker(self, nicho, lista_cidades):
        total_cidades = len(lista_cidades)

        for idx_c, cidade in enumerate(lista_cidades):
            if not self.is_running:
                break

            self.log_message(f"\n==================================================")
            self.log_message(f"[CIDADE {idx_c + 1}/{total_cidades}] Processando: '{nicho}' em '{cidade}'")
            self.log_message(f"==================================================\n")

            await self.scraper_worker(nicho, cidade)

            if idx_c + 1 < total_cidades and self.is_running:
                self.log_message(f"\n[PAUSA ANTI-BLOQUEIO] Aguardando 10 segundos antes da próxima cidade...")
                for i in range(10, 0, -1):
                    if not self.is_running:
                        break
                    self.update_status(f"Pausa entre cidades: {i}s restantes...")
                    await asyncio.sleep(1)

        self.is_running = False
        self.btn_start.configure(text="Iniciar Varredura", fg_color="#10b981", hover_color="#059669")
        self.progress_search.set(1.0)
        self.update_status("Varredura do Lote Concluída!")
        self.log_message(f"\n[{datetime.now().strftime('%H:%M:%S')}] Todas as cidades do lote foram processadas com sucesso!")

    async def scraper_worker(self, nicho, cidade):
        self.update_status(f"Abrindo Chromium ({cidade})...")
        self.log_message(f"[STATUS] Abrindo navegador para {cidade}...")
        
        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=True)
            context = await browser.new_context(
                viewport={"width": 1280, "height": 800},
                user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
            )
            page = await context.new_page()

            search_query = f"{nicho} em {cidade}"
            self.update_status(f"Buscando em {cidade}...")
            self.log_message(f"[STATUS] Acessando Google Maps para buscar: '{search_query}'...")
            
            maps_url = f"https://www.google.com/maps/search/{quote_plus(search_query)}"
            await page.goto(maps_url, wait_until="domcontentloaded")
            await asyncio.sleep(4)
            await self.update_mini_player(page)

            leads_captured = []
            
            self.update_status("Mapeando perfis...")
            feed_selector = "div[role='feed']"
            
            try:
                await page.wait_for_selector(feed_selector, timeout=8000)
                for _ in range(5):
                    if not self.is_running:
                        break
                    await page.eval_on_selector(feed_selector, "el => el.scrollBy(0, 1000)")
                    await asyncio.sleep(1.5)
                    await self.update_mini_player(page)
            except Exception:
                self.log_message("[AVISO] Painel lateral de rolagem não encontrado. Verificando perfil único...")

            elements = await page.query_selector_all("a[href*='/maps/place/']")
            place_urls = []
            for el in elements:
                href = await el.get_attribute("href")
                if href and href not in place_urls:
                    place_urls.append(href)

            if len(place_urls) == 0:
                current_url = page.url
                if "/maps/place/" in current_url:
                    self.log_message("[INFO] Resultado único detectado. Processando perfil direto...")
                    place_urls.append(current_url)

            total = len(place_urls)
            self.log_message(f"[INFO] Encontrados {total} perfis em {cidade}.")

            start_time = time.time()

            for idx, place_url in enumerate(place_urls):
                if not self.is_running:
                    break

                current_item = idx + 1
                progress_pct = current_item / total if total > 0 else 0.0
                self.progress_search.set(progress_pct)

                elapsed_time = time.time() - start_time
                avg_time_per_item = elapsed_time / current_item if current_item > 0 else 0
                remaining_items = total - current_item
                eta_seconds = int(remaining_items * avg_time_per_item)

                eta_str = f"{eta_seconds // 60:02d}m {eta_seconds % 60:02d}s" if current_item > 1 else "Calculando..."

                status_msg = f"{cidade}: {current_item}/{total} ({int(progress_pct * 100)}%) | ETA: {eta_str}"
                self.update_status(status_msg)

                try:
                    place_id = place_url.split("?")[0]

                    if self.is_lead_processed(place_id):
                        self.log_message(f"[PULADO] Empresa já cadastrada no banco SQLite.")
                        continue

                    if page.url != place_url:
                        await page.goto(place_url, wait_until="domcontentloaded")
                        await asyncio.sleep(2.0)

                    await self.update_mini_player(page)

                    nome = "N/D"
                    nome_el = await page.query_selector("h1.DUwDvf, h1")
                    if nome_el:
                        nome = await nome_el.inner_text()
                    
                    if nome.strip().lower() in ["resultados", "pesquisar neste local", "", "n/d"]:
                        continue

                    telefone = "N/D"
                    tel_el = await page.query_selector("button[data-item-id*='phone:tel:']")
                    if tel_el:
                        item_id = await tel_el.get_attribute("data-item-id")
                        if item_id:
                            telefone = item_id.replace("phone:tel:", "").strip()

                    site = "N/D"
                    site_el = await page.query_selector("a[data-item-id='authority']")
                    if site_el:
                        site = await site_el.get_attribute("href")

                    email, instagram, facebook, linkedin = "N/D", "N/D", "N/D", "N/D"
                    if site != "N/D" and site.startswith("http"):
                        email, instagram, facebook, linkedin = await self.fetch_site_info(site)

                    whatsapp_link = "N/D"
                    if telefone != "N/D":
                        raw_phone = re.sub(r'\D', '', telefone)
                        if raw_phone.startswith('0'):
                            raw_phone = raw_phone[1:]
                        if len(raw_phone) in [10, 11]:
                            whatsapp_link = f"https://wa.me/55{raw_phone}"

                    lead_data = {
                        "id": place_id,
                        "nome": nome,
                        "telefone": telefone,
                        "email": email,
                        "site": site,
                        "whatsapp": whatsapp_link,
                        "instagram": instagram,
                        "facebook": facebook,
                        "linkedin": linkedin,
                        "nicho": nicho,
                        "cidade": cidade,
                        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    }

                    self.ram_buffer.append(lead_data)
                    leads_captured.append(lead_data)

                    self.log_message(f"[CAPTURADO] {nome} | Tel: {telefone} | Email: {email} | Insta: {instagram}")

                except Exception as e:
                    continue

            await browser.close()

            self.flush_buffer_to_db()

            if leads_captured:
                saved_path = self.export_to_excel(nicho, cidade, leads_captured)
                self.log_message(f"[{datetime.now().strftime('%H:%M:%S')}] Relatório de {cidade} salvo em: {saved_path}")
            else:
                self.log_message(f"[{datetime.now().strftime('%H:%M:%S')}] [AVISO] Nenhum lead novo capturado em {cidade}. Nenhum relatório foi gerado.")

    async def fetch_site_info(self, url):
        email, instagram, facebook, linkedin = "N/D", "N/D", "N/D", "N/D"
        try:
            async with httpx.AsyncClient(timeout=4.0, follow_redirects=True, headers={"User-Agent": "Mozilla/5.0"}) as client:
                resp = await client.get(url)
                html = resp.text

                matches = re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', html)
                for e in matches:
                    e_lower = e.lower()
                    if not any(noise in e_lower for noise in ["sentry", "wixpress", "example", ".png", ".jpg", ".jpeg", ".gif", "bootstrap", "schema.org"]):
                        email = e
                        break

                soup = BeautifulSoup(html, "html.parser")
                for a in soup.find_all("a", href=True):
                    href = a["href"].strip()
                    if "instagram.com" in href and instagram == "N/D":
                        instagram = href
                    elif "facebook.com" in href and facebook == "N/D":
                        facebook = href
                    elif "linkedin.com" in href and linkedin == "N/D":
                        linkedin = href

        except Exception:
            pass

        return email, instagram, facebook, linkedin

    async def update_mini_player(self, page):
        try:
            screenshot_bytes = await page.screenshot(type="jpeg", quality=40)
            img = Image.open(io.BytesIO(screenshot_bytes))
            img = img.resize((360, 200))
            ctk_img = ctk.CTkImage(light_image=img, dark_image=img, size=(360, 200))
            self.lbl_mini_player.configure(image=ctk_img, text="")
        except:
            pass

    def flush_buffer_to_db(self):
        if not self.ram_buffer:
            return

        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        while self.ram_buffer:
            item = self.ram_buffer.popleft()
            cursor.execute("""
                INSERT OR IGNORE INTO leads (id, nome, telefone, email, site, whatsapp, instagram, facebook, linkedin, nicho, cidade, timestamp)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                item["id"], item["nome"], item["telefone"], item["email"], item["site"],
                item["whatsapp"], item["instagram"], item["facebook"], item["linkedin"],
                item["nicho"], item["cidade"], item["timestamp"]
            ))

        conn.commit()
        conn.close()

    def export_to_excel(self, nicho, cidade, leads):
        data_hora = datetime.now().strftime("%Y-%m-%d_%H-%M")
        nicho_clean = nicho.replace(" ", "_")
        cidade_clean = cidade.replace(" ", "_")
        
        filename = f"{nicho_clean}_{cidade_clean}_{data_hora}.xlsx"
        filepath = self.relatorios_dir / filename

        wb = openpyxl.Workbook()

        # ABA 1: Visão Geral do Lead
        ws1 = wb.active
        ws1.title = "Leads Geral"
        headers1 = ["Timestamp", "Nome da Empresa", "Telefone", "E-mail", "Website", "Link WhatsApp", "Instagram", "Facebook", "LinkedIn"]
        ws1.append(headers1)

        for lead in leads:
            ws1.append([
                lead["timestamp"],
                lead["nome"],
                lead["telefone"],
                lead["email"],
                lead["site"],
                lead["whatsapp"],
                lead["instagram"],
                lead["facebook"],
                lead["linkedin"]
            ])

        # ABA 2: Tabela Focada em Redes Sociais
        ws2 = wb.create_sheet(title="Redes Sociais")
        headers2 = ["Nome da Empresa", "Instagram", "Facebook", "LinkedIn", "Website", "Link WhatsApp"]
        ws2.append(headers2)

        for lead in leads:
            ws2.append([
                lead["nome"],
                lead["instagram"],
                lead["facebook"],
                lead["linkedin"],
                lead["site"],
                lead["whatsapp"]
            ])

        wb.save(filepath)
        return str(filepath)


if __name__ == "__main__":
    app = LeadHunterApp()
    app.mainloop()