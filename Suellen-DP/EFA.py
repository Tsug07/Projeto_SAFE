import re
import pandas as pd
import pdfplumber
import customtkinter as ctk
from tkinter import PhotoImage, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import os
from PIL import Image, ImageTk

class AplicativoModificacaoPDF:
    def recurso_caminho(self, relativo):
            import sys, os
            if hasattr(sys, '_MEIPASS'):
                return os.path.join(sys._MEIPASS, relativo)
            return os.path.join(os.path.abspath("."), relativo)
    
    def __init__(self, master):
        self.master = master
        master.title("E.F.A - Extração de Funcionários Ativos")
        master.geometry("800x700")
        # Adicionando ícone
        try:
            master.iconbitmap(self.recurso_caminho("efa.ico"))
        except Exception as e:
            print(f"Erro ao definir ícone da janela: {e}")
        
        # Configurar aparência
        ctk.set_appearance_mode("Dark")  # Pode ser "System", "Dark", ou "Light"
        ctk.set_default_color_theme("dark-blue")  # Outros temas: "green", "dark-blue"
        
        # Criar container principal
        self.frame_principal = ctk.CTkFrame(master)
        self.frame_principal.pack(fill="both", expand=True, padx=10, pady=10)
        
        # # Título
        # self.label_titulo = ctk.CTkLabel(
        #     self.frame_principal, 
        #     text="E.F.A",
        #     font=("Helvetica", 20, "bold")
        # )
        # self.label_titulo.pack(pady=(10, 20))
        
        # Frame para conter o ícone e o título lado a lado
        self.frame_titulo = ctk.CTkFrame(self.frame_principal)
        self.frame_titulo.pack(pady=(10, 20))

        # Carregar e redimensionar a imagem (ícone)
        try:
            # Carregue sua imagem - substitua pelo caminho correto
            imagem = Image.open(self.recurso_caminho("efa.png"))
            # Redimensione para o tamanho desejado (por exemplo, 32x32 pixels)
            imagem = imagem.resize((32, 32), Image.LANCZOS)
            # Converta para formato compatível com CTkLabel
            self.icone = ctk.CTkImage(light_image=imagem, dark_image=imagem, size=(32, 32))
            
            # Crie o label para o ícone
            self.label_icone = ctk.CTkLabel(
                self.frame_titulo,
                image=self.icone,
                text=""
            )
            self.label_icone.pack(side="left", padx=(0, 10))
        except Exception as e:
            print(f"Erro ao carregar o ícone: {str(e)}")

        # Crie o label para o texto do título
        self.label_titulo = ctk.CTkLabel(
            self.frame_titulo, 
            text="E.F.A",
            font=("Helvetica", 20, "bold")
        )
        self.label_titulo.pack(side="left")
        
        # Seleção do Arquivo PDF
        self.frame_pdf = ctk.CTkFrame(self.frame_principal)
        self.frame_pdf.pack(fill="x", padx=10, pady=5)
        
        self.label_pdf = ctk.CTkLabel(
            self.frame_pdf, 
            text="Selecionar Arquivo PDF:",
            font=("Helvetica", 12)
        )
        self.label_pdf.pack(anchor="w")
        
        self.entrada_pdf = ctk.CTkEntry(self.frame_pdf, width=500)
        self.entrada_pdf.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        self.botao_pdf = ctk.CTkButton(
            self.frame_pdf, 
            text="Procurar", 
            command=self.procurar_pdf,
            width=100
        )
        self.botao_pdf.pack(side="right")
        
        # Seleção do Arquivo Excel
        self.frame_excel = ctk.CTkFrame(self.frame_principal)
        self.frame_excel.pack(fill="x", padx=10, pady=5)
        
        self.label_excel = ctk.CTkLabel(
            self.frame_excel, 
            text="Selecionar Arquivo Excel:",
            font=("Helvetica", 12)
        )
        self.label_excel.pack(anchor="w")
        
        self.entrada_excel = ctk.CTkEntry(self.frame_excel, width=500)
        self.entrada_excel.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        self.botao_excel = ctk.CTkButton(
            self.frame_excel, 
            text="Procurar", 
            command=self.procurar_excel,
            width=100
        )
        self.botao_excel.pack(side="right")
        
        # Seleção do Arquivo de Saída
        self.frame_saida = ctk.CTkFrame(self.frame_principal)
        self.frame_saida.pack(fill="x", padx=10, pady=5)
        
        self.label_saida = ctk.CTkLabel(
            self.frame_saida, 
            text="Arquivo Excel de Saída:",
            font=("Helvetica", 12)
        )
        self.label_saida.pack(anchor="w")
        
        self.entrada_saida = ctk.CTkEntry(self.frame_saida, width=500)
        self.entrada_saida.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        self.botao_saida = ctk.CTkButton(
            self.frame_saida, 
            text="Procurar", 
            command=self.procurar_saida,
            width=100
        )
        self.botao_saida.pack(side="right")
        
        # Botão de Processamento
        self.botao_processar = ctk.CTkButton(
            self.frame_principal, 
            text="Processar Arquivos", 
            command=self.processar_arquivos,
            font=("Helvetica", 14),
            height=40
        )
        self.botao_processar.pack(pady=20)
        
        # Área de Log/Resultados
        self.label_log = ctk.CTkLabel(
            self.frame_principal, 
            text="Log de Processamento:",
            font=("Helvetica", 12)
        )
        self.label_log.pack(anchor="w", padx=10)  
        
        self.area_log = ScrolledText(
            self.frame_principal, 
            wrap="word", 
            width=80, 
            height=15,
            font=("Consolas", 10)
        )
        self.area_log.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.area_log.configure(state="disabled")
        
        # Barra de Status
        self.var_status = ctk.StringVar()
        self.var_status.set("E.F.A - Desenvolvido por Hugo Almeida\nPronto")
        self.barra_status = ctk.CTkLabel(
            self.frame_principal, 
            textvariable=self.var_status,
            font=("Helvetica", 10),
            anchor="w"
        )
        self.barra_status.pack(fill="x", padx=10, pady=(0, 5))
    
    
    
    def procurar_pdf(self):
        caminho_arquivo = filedialog.askopenfilename(
            filetypes=[("Arquivos PDF", "*.pdf"), ("Todos os Arquivos", "*.*")]
        )
        if caminho_arquivo:
            self.entrada_pdf.delete(0, "end")
            self.entrada_pdf.insert(0, caminho_arquivo)
    
    def procurar_excel(self):
        caminho_arquivo = filedialog.askopenfilename(
            filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os Arquivos", "*.*")]
        )
        if caminho_arquivo:
            self.entrada_excel.delete(0, "end")
            self.entrada_excel.insert(0, caminho_arquivo)
    
    def procurar_saida(self):
        caminho_arquivo = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os Arquivos", "*.*")]
        )
        if caminho_arquivo:
            self.entrada_saida.delete(0, "end")
            self.entrada_saida.insert(0, caminho_arquivo)
    
    def registrar_mensagem(self, mensagem):
        self.area_log.configure(state="normal")
        self.area_log.insert("end", mensagem + "\n")
        self.area_log.configure(state="disabled")
        self.area_log.see("end")
        self.master.update()
    
    def atualizar_status(self, mensagem):
        self.var_status.set(mensagem)
        self.master.update()
    
    def extrair_texto_do_pdf(self, caminho_pdf):
        self.registrar_mensagem(f"Extraindo texto do PDF: {caminho_pdf}")
        texto = ""
        try:
            with pdfplumber.open(caminho_pdf) as pdf:
                for pagina in pdf.pages:
                    texto += pagina.extract_text() + "\n"
            return texto
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao ler o arquivo PDF: {str(e)}")
            return None
    
    def extrair_dados_funcionarios(self, texto):
        self.registrar_mensagem("Extraindo dados dos funcionários do texto do PDF...")
        padrao = r"Empr\.: (\d+)([A-ZÀ-ÚÇÑ\s]+) Situação:"
        correspondencias = re.findall(padrao, texto)
        
        codigos_funcionarios = []
        nomes_funcionarios = []
        
        for correspondencia in correspondencias:
            codigo, nome = correspondencia
            codigos_funcionarios.append(codigo)
            nomes_funcionarios.append(nome.strip())
        
        return codigos_funcionarios, nomes_funcionarios
    
    def comparar_e_filtrar_excel(self, codigos_pdf, caminho_excel, caminho_excel_saida):
        self.registrar_mensagem(f"Comparando com o arquivo Excel: {caminho_excel}")
        
        try:
            # Ler arquivo Excel
            df = pd.read_excel(caminho_excel, header=None, skiprows=6)
            codigos_excel = df[0].astype(str).tolist()
            
            # Encontrar códigos correspondentes
            codigos_correspondentes = set(codigos_pdf).intersection(set(codigos_excel))
            
            self.registrar_mensagem("\nCódigos correspondentes entre PDF e Excel:")
            if codigos_correspondentes:
                # for codigo in codigos_correspondentes:
                #     self.registrar_mensagem(f"Código: {codigo}")
                self.registrar_mensagem(f"Total de códigos correspondentes: {len(codigos_correspondentes)}")
            else:
                self.registrar_mensagem("Nenhum código correspondente encontrado.")
            
            # Filtrar DataFrame
            df_filtrado = df[df[0].astype(str).isin(codigos_correspondentes)]
            
            # Formatar colunas antes de salvar
            # CPF (Coluna F, índice 5) - formatar como texto com zeros à esquerda (11 dígitos)
            df_filtrado[5] = df_filtrado[5].astype(str).str.zfill(11)
            
            # PIS (Coluna M, índice 12) - formatar como texto
            df_filtrado[12] = df_filtrado[12].astype(str)
            
            # Datas (Colunas H e O, índices 7 e 14)
            for col in [7, 14]:
                if pd.api.types.is_datetime64_any_dtype(df_filtrado[col]):
                    df_filtrado[col] = df_filtrado[col].dt.strftime('%d/%m/%Y')
                elif pd.api.types.is_numeric_dtype(df_filtrado[col]):
                    # Converter data serial do Excel para datetime
                    df_filtrado[col] = pd.to_datetime(df_filtrado[col], unit='D', origin='1899-12-30').dt.strftime('%d/%m/%Y')
                else:
                    df_filtrado[col] = df_filtrado[col].astype(str).str[:10]  # Pegar apenas a parte da data se for string
            
            # Ler Excel original para o cabeçalho
            df_original = pd.read_excel(caminho_excel, header=None)
            df_cabecalho = df_original.iloc[:6]
            
            # Criar um escritor de Excel usando openpyxl
            with pd.ExcelWriter(caminho_excel_saida, engine='openpyxl') as escritor:
                # Escrever cabeçalho
                df_cabecalho.to_excel(escritor, sheet_name='Sheet1', index=False, header=False)
                
                # Escrever dados filtrados começando da linha 7
                df_filtrado.to_excel(
                    escritor, 
                    sheet_name='Sheet1', 
                    index=False, 
                    header=False, 
                    startrow=6  # Pular linhas de cabeçalho
                )
                
                # Obter objetos workbook e worksheet
                workbook = escritor.book
                worksheet = escritor.sheets['Sheet1']
                
                # Aplicar formato de texto à coluna CPF (F) e coluna PIS (M)
                for idx_col, letra_col in [(6, 'F'), (13, 'M')]:  # F=6, M=13 (baseado em 0)
                    for linha in worksheet.iter_rows(min_row=7, max_row=worksheet.max_row, min_col=idx_col, max_col=idx_col):
                        for celula in linha:
                            celula.number_format = '@'  # Formato de texto
                
                # Aplicar formato de data às colunas de data (H e O)
                estilo_data = 'DD/MM/YYYY'
                for letra_col in ['H', 'O']:
                    idx_col = ord(letra_col) - 64  # Converter letra para índice baseado em 1
                    for linha in worksheet.iter_rows(min_row=7, max_row=worksheet.max_row, min_col=idx_col, max_col=idx_col):
                        for celula in linha:
                            celula.number_format = estilo_data
            
            self.registrar_mensagem(f"\nNovo arquivo Excel salvo em: {caminho_excel_saida}")
            self.registrar_mensagem(f"Total de linhas no novo Excel (excluindo cabeçalho): {len(df_filtrado)}")
            
            return True
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao processar arquivo Excel: {str(e)}")
            return False
    
    def processar_arquivos(self):
        # Limpar área de log
        self.area_log.configure(state="normal")
        self.area_log.delete(1.0, "end")
        self.area_log.configure(state="disabled")
        
        # Obter caminhos dos arquivos
        caminho_pdf = self.entrada_pdf.get()
        caminho_excel = self.entrada_excel.get()
        caminho_saida = self.entrada_saida.get()
        
        # Validar entradas
        if not caminho_pdf or not caminho_excel or not caminho_saida:
            messagebox.showwarning("Aviso", "Por favor, selecione todos os arquivos necessários")
            return
        
        if not os.path.exists(caminho_pdf):
            messagebox.showerror("Erro", "O arquivo PDF não existe")
            return
        
        if not os.path.exists(caminho_excel):
            messagebox.showerror("Erro", "O arquivo Excel não existe")
            return
        
        # Processar arquivos
        self.atualizar_status("Processando...")
        
        try:
            # Extrair texto do PDF
            texto_pdf = self.extrair_texto_do_pdf(caminho_pdf)
            if texto_pdf is None:
                return
            
            # Extrair dados dos funcionários
            codigos, nomes = self.extrair_dados_funcionarios(texto_pdf)
            
            # Registrar resultados
            self.registrar_mensagem(f"\nTotal de funcionários encontrados: {len(codigos)}")
            self.registrar_mensagem("\nPrimeiros 5 funcionários:")
            for i in range(min(5, len(codigos))):
                self.registrar_mensagem(f"{i+1}. {codigos[i]} {nomes[i]}")
            
            # Comparar e filtrar Excel
            sucesso = self.comparar_e_filtrar_excel(codigos, caminho_excel, caminho_saida)
            
            if sucesso:
                self.atualizar_status("E.F.A - Desenvolvido por Hugo Almeida\nProcessamento concluído com sucesso")
                messagebox.showinfo("Sucesso", "Arquivos processados com sucesso!")
            else:
                self.atualizar_status("Processamento falhou")
        except Exception as e:
            self.registrar_mensagem(f"\nErro: {str(e)}")
            self.atualizar_status("Ocorreu um erro")
            messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")

if __name__ == "__main__":
    root = ctk.CTk()
    app = AplicativoModificacaoPDF(root)
    root.mainloop()