import os
import sys
import time
import json
import threading
import multiprocessing
from datetime import datetime
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox

# Configuracoes visuais do CustomTkinter
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

PRESETS_FILE = Path("tarefas_frequentes.json")

def carregar_presets_disco():
    if PRESETS_FILE.exists():
        try:
            with open(PRESETS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def salvar_presets_disco(presets):
    with open(PRESETS_FILE, "w", encoding="utf-8") as f:
        json.dump(presets, f, indent=4, ensure_ascii=False)

def processar_arquivo_individual(args):
    arq_str, aba_alvo, linha_cabecalho, col_validadora = args
    arq = Path(arq_str)
    
    aba_alvo_normalizada = aba_alvo.strip().lower()

    engine_usado = "calamine"
    try:
        excel_file = pd.ExcelFile(arq, engine="calamine")
    except Exception:
        engine_usado = "openpyxl"
        excel_file = pd.ExcelFile(arq, engine="openpyxl")

    mapa_abas = {sheet.strip().lower(): sheet for sheet in excel_file.sheet_names}

    if aba_alvo_normalizada not in mapa_abas:
        return {"status": "ignorado", "arquivo": arq.name, "motivo": f"Aba '{aba_alvo}' não encontrada"}

    nome_real_aba = mapa_abas[aba_alvo_normalizada]

    df = pd.read_excel(
        arq, 
        sheet_name=nome_real_aba, 
        header=linha_cabecalho,
        engine=engine_usado
    )

    df.dropna(how="all", inplace=True)

    if col_validadora and not df.empty:
        mapa_cols = {str(c).strip().lower(): c for c in df.columns}
        col_norm = col_validadora.strip().lower()
        if col_norm in mapa_cols:
            col_real = mapa_cols[col_norm]
            mask = df[col_real].notna() & (df[col_real].astype(str).str.strip() != "") & (df[col_real].astype(str).str.strip() != "-")
            df = df[mask]

    if not df.empty:
        df.insert(0, "Arquivo_Origem", arq.name)
        return {"status": "ok", "arquivo": arq.name, "df": df, "linhas": len(df)}
    else:
        return {"status": "vazio", "arquivo": arq.name}


def salvar_com_substituicao_segura(df, caminho_base, fmt_csv, fmt_xlsx, log_callback):
    sucesso_geral = True
    erros = []

    formatos = []
    if fmt_csv: formatos.append(("csv", f"{caminho_base}.csv"))
    if fmt_xlsx: formatos.append(("xlsx", f"{caminho_base}.xlsx"))

    for fmt_tipo, caminho_final in formatos:
        caminho_temp = f"{caminho_final}.tmp"
        log_callback(f" -> Gerando {fmt_tipo.upper()}: {caminho_final}...")

        try:
            if fmt_tipo == "csv":
                df.to_csv(caminho_temp, index=False, encoding="utf-8-sig", sep=";")
            elif fmt_tipo == "xlsx":
                df.to_excel(caminho_temp, index=False, engine="openpyxl")

            substituido = False
            for tentativa in range(1, 6):
                try:
                    if os.path.exists(caminho_final):
                        os.replace(caminho_temp, caminho_final)
                    else:
                        os.rename(caminho_temp, caminho_final)
                    substituido = True
                    break
                except PermissionError:
                    log_callback(f" [AVISO] Arquivo '{Path(caminho_final).name}' bloqueado (Power BI lendo?). Tentativa {tentativa}/5...")
                    time.sleep(2)
                except Exception as ex:
                    log_callback(f" [ERRO] {str(ex)}")
                    break

            if not substituido:
                sucesso_geral = False
                erros.append(f"Não foi possível sobrescrever {Path(caminho_final).name}. Verifique se o arquivo está aberto.")
                if os.path.exists(caminho_temp):
                    try: os.remove(caminho_temp)
                    except Exception: pass

        except Exception as e:
            sucesso_geral = False
            erros.append(f"Erro ao salvar {fmt_tipo}: {str(e)}")

    return sucesso_geral, erros


class ExcelAggregatorApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Agregador de Planilhas")
        self.geometry("820x820")
        self.minsize(750, 750)

        self.lista_origens = []
        self.presets = carregar_presets_disco()
        self.auto_loop_ativo = False
        self.timer_thread = None

        self.setup_ui()

    def setup_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(8, weight=1)

        # Banner Título Simplificado
        self.banner_frame = ctk.CTkFrame(self, fg_color=("#1f538d", "#14375e"), corner_radius=0)
        self.banner_frame.grid(row=0, column=0, sticky="ew", padx=0, pady=(0, 10))
        
        self.lbl_title = ctk.CTkLabel(
            self.banner_frame, 
            text="Agregador de Planilhas", 
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color="white"
        )
        self.lbl_title.pack(padx=20, pady=12, anchor="w")

        # Section 1: Tarefas Frequentes (Presets)
        self.preset_frame = ctk.CTkFrame(self)
        self.preset_frame.grid(row=1, column=0, sticky="ew", padx=20, pady=5)
        self.preset_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(self.preset_frame, text="Tarefa Salva:", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=10, pady=8, sticky="w")
        
        self.opt_presets = ctk.CTkOptionMenu(self.preset_frame, values=[], command=self.carregar_preset_selecionado)
        self.opt_presets.grid(row=0, column=1, padx=10, pady=8, sticky="ew")
        self.atualizar_dropdown_presets()

        btn_salvar_preset = ctk.CTkButton(self.preset_frame, text="Salvar Tarefa", width=110, fg_color="#17a2b8", hover_color="#138496", command=self.salvar_preset_dialog)
        btn_salvar_preset.grid(row=0, column=2, padx=5, pady=8)

        btn_del_preset = ctk.CTkButton(self.preset_frame, text="Excluir", width=70, fg_color="#dc3545", hover_color="#bd2130", command=self.excluir_preset)
        btn_del_preset.grid(row=0, column=3, padx=(5, 10), pady=8)

        # Section 2: Origens
        self.origem_frame = ctk.CTkFrame(self)
        self.origem_frame.grid(row=2, column=0, sticky="ew", padx=20, pady=5)
        self.origem_frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(self.origem_frame, text="Origens (Pastas e/ou Arquivos):", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=10, pady=(8, 2), sticky="w")
        
        self.txt_origens = ctk.CTkTextbox(self.origem_frame, height=75, font=ctk.CTkFont(size=11))
        self.txt_origens.grid(row=1, column=0, columnspan=3, padx=10, pady=5, sticky="ew")

        btn_add_folder = ctk.CTkButton(self.origem_frame, text="+ Adicionar Pasta (Recursiva)", command=self.add_pasta)
        btn_add_folder.grid(row=2, column=0, padx=10, pady=(0, 8), sticky="w")

        btn_add_files = ctk.CTkButton(self.origem_frame, text="+ Adicionar Arquivos", command=self.add_arquivos)
        btn_add_files.grid(row=2, column=1, padx=5, pady=(0, 8), sticky="w")

        btn_clear_origem = ctk.CTkButton(self.origem_frame, text="Limpar Origens", fg_color="#6c757d", hover_color="#5a6268", command=self.limpar_origens)
        btn_clear_origem.grid(row=2, column=2, padx=10, pady=(0, 8), sticky="e")

        # Section 3: Destino Final
        self.dest_frame = ctk.CTkFrame(self)
        self.dest_frame.grid(row=3, column=0, sticky="ew", padx=20, pady=5)
        self.dest_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(self.dest_frame, text="Destino Final:", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=10, pady=8, sticky="w")
        self.entry_destino = ctk.CTkEntry(self.dest_frame, placeholder_text="Deixe em branco p/ Desktop ou escolha pasta/nome fixo")
        self.entry_destino.grid(row=0, column=1, padx=10, pady=8, sticky="ew")

        btn_destino = ctk.CTkButton(self.dest_frame, text="Salvar Em...", width=100, command=self.selecionar_destino)
        btn_destino.grid(row=0, column=2, padx=10, pady=8)

        # Section 4: Parâmetros
        self.param_frame = ctk.CTkFrame(self)
        self.param_frame.grid(row=4, column=0, sticky="ew", padx=20, pady=5)
        self.param_frame.grid_columnconfigure(1, weight=1)
        self.param_frame.grid_columnconfigure(3, weight=1)
        self.param_frame.grid_columnconfigure(5, weight=1)

        ctk.CTkLabel(self.param_frame, text="Aba:", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=8, pady=8, sticky="w")
        self.entry_aba = ctk.CTkEntry(self.param_frame, placeholder_text="Ex: Dados")
        self.entry_aba.grid(row=0, column=1, padx=5, pady=8, sticky="ew")

        ctk.CTkLabel(self.param_frame, text="Cabeçalho (Linha):", font=ctk.CTkFont(weight="bold")).grid(row=0, column=2, padx=8, pady=8, sticky="w")
        self.entry_linha = ctk.CTkEntry(self.param_frame, width=50)
        self.entry_linha.insert(0, "2")
        self.entry_linha.grid(row=0, column=3, padx=5, pady=8, sticky="w")

        ctk.CTkLabel(self.param_frame, text="Coluna Validadora:", font=ctk.CTkFont(weight="bold")).grid(row=0, column=4, padx=8, pady=8, sticky="w")
        self.entry_validadora = ctk.CTkEntry(self.param_frame, placeholder_text="Ex: CPF, Data (Opcional)")
        self.entry_validadora.grid(row=0, column=5, padx=8, pady=8, sticky="ew")

        # Section 5: Formatos (CSV e XLSX) & Automação em Background
        self.fmt_frame = ctk.CTkFrame(self)
        self.fmt_frame.grid(row=5, column=0, sticky="ew", padx=20, pady=5)
        
        self.var_csv = ctk.BooleanVar(value=True)
        self.var_xlsx = ctk.BooleanVar(value=False)

        ctk.CTkCheckBox(self.fmt_frame, text="CSV (.csv)", variable=self.var_csv).pack(side="left", padx=10, pady=8)
        ctk.CTkCheckBox(self.fmt_frame, text="Excel (.xlsx)", variable=self.var_xlsx).pack(side="left", padx=10, pady=8)

        self.var_autoloop = ctk.BooleanVar(value=False)
        self.chk_autoloop = ctk.CTkCheckBox(self.fmt_frame, text="Loop Automático a cada:", variable=self.var_autoloop, command=self.toggle_autoloop)
        self.chk_autoloop.pack(side="left", padx=(30, 5), pady=8)

        self.entry_intervalo = ctk.CTkEntry(self.fmt_frame, width=50)
        self.entry_intervalo.insert(0, "15")
        self.entry_intervalo.pack(side="left", padx=2, pady=8)
        ctk.CTkLabel(self.fmt_frame, text="min").pack(side="left", padx=2, pady=8)

        # Botão Ação
        self.btn_processar = ctk.CTkButton(
            self, 
            text="INICIAR AGREGAÇÃO AGORA", 
            font=ctk.CTkFont(size=14, weight="bold"),
            height=40,
            fg_color="#2ba84a",
            hover_color="#1e7e34",
            command=self.iniciar_processamento_manual
        )
        self.btn_processar.grid(row=6, column=0, sticky="ew", padx=20, pady=(8, 4))

        # Status & Progress Bar
        self.status_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.status_frame.grid(row=7, column=0, sticky="ew", padx=20, pady=(0, 2))
        self.status_frame.grid_columnconfigure(0, weight=1)

        self.lbl_eta = ctk.CTkLabel(self.status_frame, text="Pronto para iniciar", font=ctk.CTkFont(size=12, weight="bold"))
        self.lbl_eta.grid(row=0, column=0, sticky="w")

        self.lbl_pct = ctk.CTkLabel(self.status_frame, text="0%", font=ctk.CTkFont(size=12, weight="bold"))
        self.lbl_pct.grid(row=0, column=1, sticky="e")

        self.progress_bar = ctk.CTkProgressBar(self, orientation="horizontal", mode="determinate")
        self.progress_bar.grid(row=8, column=0, sticky="ew", padx=20, pady=(0, 8))
        self.progress_bar.set(0)

        # Console Log
        self.log_frame = ctk.CTkFrame(self)
        self.log_frame.grid(row=9, column=0, sticky="nsew", padx=20, pady=(0, 8))
        self.log_frame.grid_columnconfigure(0, weight=1)
        self.log_frame.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(self.log_frame, text="Log de Execução e Status:", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=10, pady=(5, 0), sticky="w")
        
        self.txt_log = ctk.CTkTextbox(self.log_frame, state="disabled", wrap="word", font=ctk.CTkFont(family="Consolas", size=11))
        self.txt_log.grid(row=1, column=0, sticky="nsew", padx=10, pady=8)

        # Assinatura
        self.lbl_credits = ctk.CTkLabel(
            self, 
            text="by Lucas G. Oliveira", 
            font=ctk.CTkFont(size=12, weight="bold", slant="italic"), 
            text_color=("gray30", "gray70")
        )
        self.lbl_credits.grid(row=10, column=0, pady=(0, 8))

    def log(self, mensagem):
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", mensagem + "\n")
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")
        self.update_idletasks()

    def add_pasta(self):
        pasta = filedialog.askdirectory(title="Selecione a Pasta para Busca Recursiva")
        if pasta and pasta not in self.lista_origens:
            self.lista_origens.append(pasta)
            self.atualizar_txt_origens()

    def add_arquivos(self):
        arquivos = filedialog.askopenfilenames(title="Selecione os Arquivos Excel", filetypes=[("Planilhas Excel", "*.xlsx *.xls")])
        if arquivos:
            for arq in arquivos:
                if arq not in self.lista_origens:
                    self.lista_origens.append(arq)
            self.atualizar_txt_origens()

    def limpar_origens(self):
        self.lista_origens.clear()
        self.atualizar_txt_origens()

    def atualizar_txt_origens(self):
        self.txt_origens.configure(state="normal")
        self.txt_origens.delete("1.0", "end")
        for p in self.lista_origens:
            self.txt_origens.insert("end", p + "\n")
        self.txt_origens.configure(state="disabled")

    def selecionar_destino(self):
        arquivo_saida = filedialog.asksaveasfilename(
            title="Selecione o Local e Nome do Arquivo de Saída (Sem extensão)",
            filetypes=[("Todos os arquivos", "*.*")]
        )
        if arquivo_saida:
            caminho_base = os.path.splitext(arquivo_saida)[0]
            self.entry_destino.delete(0, "end")
            self.entry_destino.insert(0, caminho_base)

    def atualizar_dropdown_presets(self):
        opcoes = ["+ Nova Tarefa (Limpar Campos)"] + list(self.presets.keys())
        self.opt_presets.configure(values=opcoes)
        if self.opt_presets.get() not in opcoes:
            self.opt_presets.set("+ Nova Tarefa (Limpar Campos)")

    def carregar_preset_selecionado(self, escolha):
        if escolha == "+ Nova Tarefa (Limpar Campos)":
            self.limpar_origens()
            self.entry_destino.delete(0, "end")
            self.entry_aba.delete(0, "end")
            self.entry_linha.delete(0, "end")
            self.entry_linha.insert(0, "2")
            self.entry_validadora.delete(0, "end")
            self.var_csv.set(True)
            self.var_xlsx.set(False)
            self.entry_intervalo.delete(0, "end")
            self.entry_intervalo.insert(0, "15")
            self.log("[PRESET] Campos limpos para criação de uma nova tarefa.")
        elif escolha in self.presets:
            p = self.presets[escolha]
            self.lista_origens = p.get("origens", [])
            self.atualizar_txt_origens()

            self.entry_destino.delete(0, "end")
            self.entry_destino.insert(0, p.get("destino", ""))

            self.entry_aba.delete(0, "end")
            self.entry_aba.insert(0, p.get("aba", ""))

            self.entry_linha.delete(0, "end")
            self.entry_linha.insert(0, p.get("linha", "2"))

            self.entry_validadora.delete(0, "end")
            self.entry_validadora.insert(0, p.get("validadora", ""))

            self.var_csv.set(p.get("csv", True))
            self.var_xlsx.set(p.get("xlsx", False))

            self.entry_intervalo.delete(0, "end")
            self.entry_intervalo.insert(0, p.get("intervalo", "15"))

    def salvar_preset_dialog(self):
        dialog = ctk.CTkInputDialog(text="Digite um nome para esta tarefa:", title="Salvar Tarefa Frequente")
        nome = dialog.get_input()
        if nome:
            nome = nome.strip()
            if not nome or nome == "+ Nova Tarefa (Limpar Campos)": return
            
            self.presets[nome] = {
                "origens": self.lista_origens,
                "destino": self.entry_destino.get().strip(),
                "aba": self.entry_aba.get().strip(),
                "linha": self.entry_linha.get().strip(),
                "validadora": self.entry_validadora.get().strip(),
                "csv": self.var_csv.get(),
                "xlsx": self.var_xlsx.get(),
                "intervalo": self.entry_intervalo.get().strip()
            }
            salvar_presets_disco(self.presets)
            self.atualizar_dropdown_presets()
            self.opt_presets.set(nome)
            messagebox.showinfo("Sucesso", f"Tarefa '{nome}' salva com sucesso!")

    def excluir_preset(self):
        escolha = self.opt_presets.get()
        if escolha in self.presets:
            if messagebox.askyesno("Confirmar", f"Deseja excluir a tarefa '{escolha}'?"):
                del self.presets[escolha]
                salvar_presets_disco(self.presets)
                self.atualizar_dropdown_presets()
                self.opt_presets.set("+ Nova Tarefa (Limpar Campos)")
                self.carregar_preset_selecionado("+ Nova Tarefa (Limpar Campos)")

    def toggle_autoloop(self):
        if self.var_autoloop.get():
            self.auto_loop_ativo = True
            self.btn_processar.configure(state="disabled")
            self.timer_thread = threading.Thread(target=self.loop_background_daemon, daemon=True)
            self.timer_thread.start()
            self.log("[BACKGROUND] Modo de execução automática ATIVADO.")
        else:
            self.auto_loop_ativo = False
            self.btn_processar.configure(state="normal")
            self.log("[BACKGROUND] Modo de execução automática DESATIVADO.")

    def loop_background_daemon(self):
        while self.auto_loop_ativo:
            self.processar_dados(eh_automatico=True)
            
            try:
                minutos = int(self.entry_intervalo.get().strip())
                if minutos < 1: minutos = 1
            except ValueError:
                minutos = 15

            segundos_espera = minutos * 60
            self.log(f"\n[BACKGROUND] Próxima execução em {minutos} minuto(s)...")
            
            for _ in range(segundos_espera):
                if not self.auto_loop_ativo: break
                time.sleep(1)

    def iniciar_processamento_manual(self):
        thread = threading.Thread(target=self.processar_dados, args=(False,), daemon=True)
        thread.start()

    def atualizar_progresso(self, concluidos, total, tempo_decorrido):
        progresso = (concluidos / total) * 0.99
        self.progress_bar.set(progresso)
        
        porcentagem = int(progresso * 100)
        self.lbl_pct.configure(text=f"{porcentagem}%")

        if concluidos > 0:
            tempo_medio = tempo_decorrido / concluidos
            restantes = total - concluidos
            tempo_est = restantes * tempo_medio

            if tempo_est < 60:
                eta_str = f"~{int(tempo_est)}s restantes"
            else:
                eta_str = f"~{int(tempo_est // 60)}m {int(tempo_est % 60)}s restantes"

            self.lbl_eta.configure(text=f"Processando ({concluidos}/{total}) - {eta_str}")
        self.update_idletasks()

    def processar_dados(self, eh_automatico=False):
        destino_base_user = self.entry_destino.get().strip()
        aba_alvo = self.entry_aba.get().strip()
        linha_cabecalho_str = self.entry_linha.get().strip()
        col_validadora = self.entry_validadora.get().strip()

        fmt_csv = self.var_csv.get()
        fmt_xlsx = self.var_xlsx.get()

        if not self.lista_origens:
            if not eh_automatico: messagebox.showerror("Erro", "Adicione pelo menos uma pasta ou arquivo de origem!")
            return

        if not aba_alvo:
            if not eh_automatico: messagebox.showerror("Erro", "Informe o nome da aba!")
            return

        try:
            linha_cabecalho = int(linha_cabecalho_str) - 1
            if linha_cabecalho < 0: raise ValueError
        except ValueError:
            if not eh_automatico: messagebox.showerror("Erro", "Linha do cabeçalho inválida!")
            return

        nome_aba_limpo = "".join([c for c in aba_alvo if c.isalnum() or c in (' ', '_', '-')]).strip()
        nome_padrao = f"{nome_aba_limpo}_Consolidado"

        if not destino_base_user:
            destino_base = str(Path.home() / "Desktop" / nome_padrao)
        elif os.path.isdir(destino_base_user):
            destino_base = str(Path(destino_base_user) / nome_padrao)
        else:
            destino_base = destino_base_user

        if not eh_automatico: self.btn_processar.configure(state="disabled")
        self.progress_bar.set(0)
        self.lbl_pct.configure(text="0%")
        self.lbl_eta.configure(text="Mapeando arquivos...")

        try:
            self.log(f"\n=== INICIANDO AGREGAÇÃO ({datetime.now().strftime('%H:%M:%S')}) ===")
            self.log(f"Destino Final: {destino_base}")
            if col_validadora: self.log(f"Coluna Validadora (Filtro): '{col_validadora}'")

            arquivos_excel = []
            for item in self.lista_origens:
                p = Path(item)
                if p.is_file() and p.suffix.lower() in ['.xlsx', '.xls']:
                    arquivos_excel.append(p)
                elif p.is_dir():
                    for sub in p.rglob("*"):
                        if sub.suffix.lower() in ['.xlsx', '.xls'] and not sub.name.startswith("~$"):
                            arquivos_excel.append(sub)

            total_arquivos = len(arquivos_excel)
            self.log(f"Encontrados {total_arquivos} arquivos para leitura.")

            if total_arquivos == 0:
                self.log("[AVISO] Nenhum arquivo Excel encontrado.")
                self.lbl_eta.configure(text="Sem arquivos.")
                return

            dataframes = []
            processados = 0
            inicio_tempo = time.time()

            tarefas = [(str(arq), aba_alvo, linha_cabecalho, col_validadora) for arq in arquivos_excel]
            max_trabalhadores = min(8, total_arquivos)

            with ThreadPoolExecutor(max_workers=max_trabalhadores) as executor:
                futuros = [executor.submit(processar_arquivo_individual, t) for t in tarefas]
                
                for idx, futuro in enumerate(as_completed(futuros), start=1):
                    res = futuro.result()
                    if res["status"] == "ok":
                        dataframes.append(res["df"])
                        processados += 1
                        self.log(f"[{idx}/{total_arquivos}] [OK] '{res['arquivo']}' -> {res['linhas']} linhas.")
                    elif res["status"] == "ignorado":
                        self.log(f"[{idx}/{total_arquivos}] [IGNORADO] '{res['arquivo']}' -> {res['motivo']}")

                    self.atualizar_progresso(idx, total_arquivos, time.time() - inicio_tempo)

            if dataframes:
                self.progress_bar.set(0.99)
                self.lbl_pct.configure(text="99%")
                self.lbl_eta.configure(text="Gravando arquivo final...")
                
                df_consolidado = pd.concat(dataframes, ignore_index=True, sort=False)
                df_consolidado = df_consolidado.astype(object).fillna("-")

                sucesso, erros = salvar_com_substituicao_segura(
                    df_consolidado, destino_base, fmt_csv, fmt_xlsx, self.log
                )

                tempo_total = time.time() - inicio_tempo
                
                if sucesso:
                    agora = datetime.now().strftime("%H:%M:%S")
                    self.progress_bar.set(1.0)
                    self.lbl_pct.configure(text="100%")
                    self.lbl_eta.configure(text=f"Concluído às {agora} ({tempo_total:.1f}s)")
                    self.log(f"=== SUCESSO! {len(df_consolidado)} linhas consolidadas às {agora} (em {tempo_total:.2f}s) ===")
                    if not eh_automatico:
                        messagebox.showinfo("Sucesso", f"Concluído com sucesso às {agora}!\nLinhas: {len(df_consolidado)}")
                else:
                    self.log("\n[ATENÇÃO] Houve falhas ao gravar o arquivo final:")
                    for err in erros: self.log(f" - {err}")
                    if not eh_automatico:
                        messagebox.showwarning("Aviso de Gravação", "\n".join(erros))
            else:
                self.log("\n[ATENÇÃO] Nenhuma linha extraída.")
                self.lbl_eta.configure(text="Nenhum dado extraído.")

        except Exception as e:
            self.log(f"\n[ERRO CRÍTICO] {str(e)}")
        finally:
            if not eh_automatico and not self.auto_loop_ativo:
                self.btn_processar.configure(state="normal")


if __name__ == "__main__":
    multiprocessing.freeze_support()
    app = ExcelAggregatorApp()
    app.mainloop()
