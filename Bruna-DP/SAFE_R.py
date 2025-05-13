import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import pandas as pd
import re
import threading
from PIL import Image, ImageTk
import uuid

# Configurações do tema CustomTkinter
ctk.set_appearance_mode("dark")  # Modos: "System", "Dark", "Light"
ctk.set_default_color_theme("blue")  # Temas: "blue", "green", "dark-blue"

class SafeApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Configuração da janela principal
        self.title("SAFE R - Sistema de Alocação de Elementos Rubrica")
        self.geometry("700x550")  # Tamanho ajustado da interface
        self.resizable(True, True)
        
        # Definir ícone da janela
        try:
            # Caminho para o ícone (mesmo diretório do script)
            icon_path = os.path.join(os.path.dirname(__file__), "safeProgram_icon.ico")
            self.iconbitmap(icon_path)
        except Exception as e:
            print(f"Erro ao carregar o ícone: {e}")
        
        # Variáveis para armazenar caminhos dos arquivos
        self.arquivo_rubricas = tk.StringVar()
        self.arquivo_ativos = tk.StringVar()
        self.arquivo_saida = tk.StringVar(value="ativos_com_elementos.xlsx")
        
        # Status do processamento
        self.status_var = tk.StringVar(value="Aguardando seleção de arquivos...")
        
        # Criação do layout
        self.criar_widgets()
        
    def criar_widgets(self):
        # Frame principal
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Logo e título
        title_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        title_frame.pack(pady=(10, 20))
        
        
        
        # Carregar e exibir a imagem PNG ao lado do título
        try:
            # Caminho para o PNG (mesmo diretório do script)
            img_path = os.path.join(os.path.dirname(__file__), "safeImg.png")
            # Carregar e redimensionar a imagem
            img = Image.open(img_path)
            img = img.resize((50, 50), Image.LANCZOS)  # Redimensionar para 40x40 pixels
            icone = ctk.CTkImage(light_image=img, dark_image=img, size=(50, 50))
            
            # img_tk = ImageTk.PhotoImage(img)
            # Adicionar a imagem como um CTkLabel
            img_label = ctk.CTkLabel(title_frame, image=icone, text="")
            # img_label.image = img_tk  # Manter referência para evitar garbage collection
            img_label.pack(side="left",pady=(0, 10))
        except Exception as e:
            print(f"Erro ao carregar a imagem PNG: {e}")
        
        logo_label = ctk.CTkLabel(title_frame, text="SAFE R", font=("Helvetica", 20, "bold"))
        logo_label.pack(side="left")
        # subtitle = ctk.CTkLabel(title_frame, text="Sistema de Alocação de Elementos", 
        #                         font=ctk.CTkFont(size=14))
        # subtitle.pack(pady=(0, 5))
        
        # Linha separadora
        separator = ctk.CTkFrame(main_frame, height=2, fg_color=("gray70", "gray30"))
        separator.pack(fill="x", pady=(0, 10))
        
        # Frame para seleção de arquivos
        files_frame = ctk.CTkFrame(main_frame)
        files_frame.pack(fill="x", pady=5)
        
        # Arquivo de Rubricas
        rubrica_frame = ctk.CTkFrame(files_frame, fg_color="transparent")
        rubrica_frame.pack(fill="x", pady=5)
        
        rubrica_label = ctk.CTkLabel(rubrica_frame, text="Arquivo de Elementos:", 
                                    font=ctk.CTkFont(size=12, weight="bold"))
        rubrica_label.pack(anchor="w", padx=5, pady=(5, 0))
        
        rubrica_subframe = ctk.CTkFrame(rubrica_frame, fg_color="transparent")
        rubrica_subframe.pack(fill="x", padx=5)
        
        rubrica_entry = ctk.CTkEntry(rubrica_subframe, textvariable=self.arquivo_rubricas, 
                                    placeholder_text="Selecione o arquivo de elementos...", width=400)
        rubrica_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        rubrica_btn = ctk.CTkButton(rubrica_subframe, text="Procurar", 
                                    command=lambda: self.selecionar_arquivo(self.arquivo_rubricas))
        rubrica_btn.pack(side="right")
        
        # Arquivo de Ativos
        ativos_frame = ctk.CTkFrame(files_frame, fg_color="transparent")
        ativos_frame.pack(fill="x", pady=5)
        
        ativos_label = ctk.CTkLabel(ativos_frame, text="Arquivo de Ativos:", 
                                    font=ctk.CTkFont(size=12, weight="bold"))
        ativos_label.pack(anchor="w", padx=5, pady=(5, 0))
        
        ativos_subframe = ctk.CTkFrame(ativos_frame, fg_color="transparent")
        ativos_subframe.pack(fill="x", padx=5)
        
        ativos_entry = ctk.CTkEntry(ativos_subframe, textvariable=self.arquivo_ativos, 
                                    placeholder_text="Selecione o arquivo de ativos...", width=400)
        ativos_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        ativos_btn = ctk.CTkButton(ativos_subframe, text="Procurar", 
                                    command=lambda: self.selecionar_arquivo(self.arquivo_ativos))
        ativos_btn.pack(side="right")
        
        # Arquivo de Saída
        saida_frame = ctk.CTkFrame(files_frame, fg_color="transparent")
        saida_frame.pack(fill="x", pady=5)
        
        saida_label = ctk.CTkLabel(saida_frame, text="Arquivo de Saída:", 
                                   font=ctk.CTkFont(size=12, weight="bold"))
        saida_label.pack(anchor="w", padx=5, pady=(5, 0))
        
        saida_subframe = ctk.CTkFrame(saida_frame, fg_color="transparent")
        saida_subframe.pack(fill="x", padx=5)
        
        saida_entry = ctk.CTkEntry(saida_subframe, textvariable=self.arquivo_saida, 
                                   placeholder_text="Selecione o arquivo de saída...", width=400)
        saida_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        saida_btn = ctk.CTkButton(saida_subframe, text="Procurar", 
                                  command=self.selecionar_arquivo_saida)
        saida_btn.pack(side="right")
        
        # Botão de processamento
        btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        btn_frame.pack(fill="x", pady=10)
        
        self.process_btn = ctk.CTkButton(btn_frame, text="Processar Arquivos", 
                                        font=ctk.CTkFont(size=14, weight="bold"),
                                        height=40, command=self.iniciar_processamento)
        self.process_btn.pack(pady=5)
        
        # Barra de progresso
        self.progress_bar = ctk.CTkProgressBar(main_frame)
        self.progress_bar.pack(fill="x", padx=5, pady=5)
        self.progress_bar.set(0)
        
        # Status
        status_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        status_frame.pack(fill="x", side="bottom", pady=5)
        
        status_label = ctk.CTkLabel(status_frame, textvariable=self.status_var,
                                    font=ctk.CTkFont(size=10))
        status_label.pack(anchor="w", padx=5)
        
        # Rodapé
        footer_frame = ctk.CTkFrame(main_frame, height=20, fg_color=("gray90", "gray10"))
        footer_frame.pack(fill="x", side="bottom", pady=(10, 0))
        
        footer_text = ctk.CTkLabel(footer_frame, text="© 2025 SAFE v1.0 - Desenvolvido por Hugo", 
                                   font=ctk.CTkFont(size=8))
        footer_text.pack(side="right", padx=5, pady=5)
        
    def selecionar_arquivo(self, var):
        filepath = filedialog.askopenfilename(
            filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
        )
        if filepath:
            var.set(filepath)
            self.status_var.set("Arquivo selecionado: " + os.path.basename(filepath))
    
    def selecionar_arquivo_saida(self):
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Arquivos Excel", "*.xlsx")],
            initialfile="ativos_com_elementos.xlsx"
        )
        if filepath:
            self.arquivo_saida.set(filepath)
            self.status_var.set("Arquivo de saída selecionado: " + os.path.basename(filepath))
    
    def iniciar_processamento(self):
        # Verifica se os arquivos foram selecionados
        if not self.arquivo_rubricas.get() or not self.arquivo_ativos.get() or not self.arquivo_saida.get():
            messagebox.showerror("Erro", "Selecione todos os arquivos para continuar.")
            return
        
        # Desabilita o botão durante o processamento
        self.process_btn.configure(state="disabled", text="Processando...")
        self.status_var.set("Iniciando processamento...")
        self.progress_bar.set(0.1)
        
        # Inicia o processamento em uma thread separada
        threading.Thread(target=self.processar_arquivos, daemon=True).start()
    
    def processar_arquivos(self):
        try:
            # Atualiza a interface
            self.status_var.set("Lendo arquivo de elementos...")
            self.progress_bar.set(0.2)
            self.update_idletasks()
            
            # Função de formatação CPF
            def formatar_cpf(valor):
                if pd.isna(valor):
                    return ""
                
                valor = str(valor)
                valor = re.sub(r'\D', '', valor)
                
                if valor.isdigit() and len(valor) <= 11:
                    valor = valor.zfill(11)
                    return valor
                
                return ""
            
            # Função de formatação PIS
            def formatar_pis(valor):
                if pd.isna(valor):
                    return ""
                
                valor = str(valor)
                valor = re.sub(r'\D', '', valor)
                
                if valor.isdigit() and len(valor) <= 11:
                    valor = valor.zfill(11)
                    return valor
                
                return ""
            
            # Lê a planilha de elementos, pulando as primeiras 6 linhas de metadados
            df_rubricas_raw = pd.read_excel(self.arquivo_rubricas.get(), header=None, skiprows=6)
            
            self.status_var.set("Extraindo dados de elementos...")
            self.progress_bar.set(0.3)
            self.update_idletasks()
            
            # Inicializa dados extraídos
            dados = []
            rubrica_atual = None
            
            # Percorre as linhas para identificar elementos e nomes
            for i, row in df_rubricas_raw.iterrows():
                texto_coluna_b = str(row[1]).strip().upper() if pd.notna(row[1]) else ''  # Coluna B (elemento)
                texto_coluna_e = str(row[4]).strip().upper() if pd.notna(row[4]) else ''  # Coluna E (nome)
            
                # Identifica início de grupo de elemento
                if any(p in texto_coluna_b for p in ["INSALUBRIDADE", "PERICULOSIDADE", "ADC PERICULOSIDADE"]):
                    rubrica_atual = texto_coluna_b.split("-", 1)[-1].strip()
                
                # Se encontrou nome na coluna E e elemento está definido
                elif texto_coluna_e and rubrica_atual and not texto_coluna_e.startswith("NOME") and not texto_coluna_e.startswith("SISTEMA LICENCIADO"):
                    dados.append({
                        "Nome": texto_coluna_e.strip(),
                        "Elemento": rubrica_atual
                    })
            
            # Cria DataFrame com nomes e elementos
            df_rubricas = pd.DataFrame(dados)
            
            # Verifica se o DataFrame está vazio
            if df_rubricas.empty:
                raise ValueError("Nenhum dado extraído. Verifique a estrutura do arquivo de elementos.")
            
            self.status_var.set("Lendo arquivo de ativos...")
            self.progress_bar.set(0.5)
            self.update_idletasks()
            
            # Lê a planilha de ativos
            df_ativos = pd.read_excel(self.arquivo_ativos.get())
            
            self.status_var.set("Processando dados...")
            self.progress_bar.set(0.7)
            self.update_idletasks()
            
            df_ativos['Nome'] = df_ativos['Nome'].str.strip().str.upper()
            
            # Junta os dados
            df_ativos_com_rubrica = pd.merge(df_ativos, df_rubricas, on="Nome", how="left")
            
            # Formata CPF e PIS
            self.status_var.set("Formatando CPF e PIS...")
            self.progress_bar.set(0.8)
            self.update_idletasks()
            
            # Aplica formatação de CPF e PIS
            df_ativos_com_rubrica.iloc[:, 10] = df_ativos_com_rubrica.iloc[:, 10].apply(formatar_cpf)  # Coluna K (índice 10)
            df_ativos_com_rubrica.iloc[:, 11] = df_ativos_com_rubrica.iloc[:, 11].astype(str).str.replace('\.0$', '', regex=True).str.zfill(11) # Coluna L (índice 11)
            
            # Insere a nova coluna Elemento na posição da coluna V (índice 21)
            colunas = df_ativos_com_rubrica.columns.tolist()
            if "Elemento" in colunas:
                colunas.remove("Elemento")
                colunas.insert(21, "Elemento")
                df_ativos_com_rubrica = df_ativos_com_rubrica[colunas]
            
            # Usa o caminho completo do arquivo de saída
            caminho_completo = self.arquivo_saida.get()
            nome_saida = os.path.basename(caminho_completo)
            
            self.status_var.set("Salvando resultado...")
            self.progress_bar.set(0.9)
            self.update_idletasks()
            
            # Salva o resultado
            df_ativos_com_rubrica.to_excel(caminho_completo, index=False)
            
            self.status_var.set(f"Concluído! Arquivo '{nome_saida}' gerado com sucesso.")
            self.progress_bar.set(1.0)
            
            messagebox.showinfo("Processamento Concluído", 
                               f"O arquivo '{nome_saida}' foi gerado com sucesso em '{os.path.dirname(caminho_completo)}'!")
        
        except Exception as e:
            self.status_var.set(f"Erro: {str(e)}")
            messagebox.showerror("Erro no Processamento", str(e))
        
        finally:
            # Reabilita o botão
            self.process_btn.configure(state="normal", text="Processar Arquivos")

if __name__ == "__main__":
    app = SafeApp()
    app.mainloop()