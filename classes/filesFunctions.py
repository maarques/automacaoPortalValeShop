import tkinter as tk
from tkinter import filedialog

class FilesFunctions:
    
    def __init__(self, lbl_entrada, lbl_saida, log_text):
        self.lbl_entrada_status = lbl_entrada
        self.lbl_saida_status = lbl_saida
        self.log_text = log_text
        self.arquivos_selecionados = []
        self.nome_arquivo_saida = "Cebraspe_Etiquetas_Geradas.docx"
        self.caminho_arquivo_saida_completo = ""

    def limpar_log(self):
        if self.log_text:
            self.log_text.config(state='normal')
            self.log_text.delete('1.0', tk.END)
            self.log_text.config(state='disabled')

    def selecionar_arquivos(self):
        tipos_arquivo = [
            ("Arquivos de Endereço", "*.docx *.xlsx"),
            ("Todos os arquivos", "*.*")
        ]
        arquivos = filedialog.askopenfilenames(title="Selecione um ou mais arquivos", filetypes=tipos_arquivo)
        if arquivos:
            self.arquivos_selecionados = list(arquivos)
            if self.lbl_entrada_status:
                self.lbl_entrada_status.config(text=f"{len(arquivos)} arquivos selecionados.")
            self.log(f"Entrada definida: {len(arquivos)} arquivos.")

    def selecionar_pasta(self):
        pasta = filedialog.askdirectory(title="Selecione uma pasta com os arquivos")
        if pasta:
            self.arquivos_selecionados = [pasta] 
            if self.lbl_entrada_status:
                self.lbl_entrada_status.config(text=f"Pasta selecionada: {pasta}")
            self.log(f"Entrada definida: Pasta {pasta}")

    def selecionar_saida(self):
        caminho_saida = filedialog.asksaveasfilename(
            title="Salvar arquivo Word como...",
            initialfile=self.nome_arquivo_saida,
            defaultextension=".docx",
            filetypes=[("Documento Word", "*.docx")]
        )
        if caminho_saida:
            self.caminho_arquivo_saida_completo = caminho_saida
            if self.lbl_saida_status:
                self.lbl_saida_status.config(text=f"Será salvo como: {caminho_saida}")
            self.log(f"Saída definida: {caminho_saida}")

    def log(self, mensagem: str):
        if self.log_text:
            self.log_text.config(state='normal')
            self.log_text.insert(tk.END, f"{mensagem}\n")
            self.log_text.see(tk.END) # Auto-scroll
            self.log_text.config(state='disabled')
        else:
            print(mensagem) 