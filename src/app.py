import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import sys
from pathlib import Path

# Adicionar o diretório src ao path para importar o módulo
sys.path.append(str(Path(__file__).parent / "src"))

from extractor.main import gerar_relatorio_consolidado


class CashFlowApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerador de Relatório Cash Flow - Consolidado")
        self.root.geometry("800x600")
        
        # Variáveis
        self.arquivo_entrada = tk.StringVar()
        self.arquivo_saida = tk.StringVar()
        self.sheet_pendencias = tk.StringVar(value="Pendências")
        self.sheet_novas = tk.StringVar(value="Sheet1")
        
        self.criar_interface()
        
    def criar_interface(self):
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configurar grid
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Título
        title_label = ttk.Label(main_frame, text="Gerador de Relatório Cash Flow", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Arquivo de entrada
        ttk.Label(main_frame, text="Arquivo de Entrada (.xlsx):").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.arquivo_entrada, width=50).grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(10, 5))
        ttk.Button(main_frame, text="Procurar", command=self.selecionar_arquivo_entrada).grid(row=1, column=2, padx=(5, 0))
        
        # Sheet de pendências
        ttk.Label(main_frame, text="Nome da Sheet - Pendências:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.sheet_pendencias, width=30).grid(row=2, column=1, sticky=tk.W, padx=(10, 0))
        
        # Sheet de novas transações
        ttk.Label(main_frame, text="Nome da Sheet - Novas Transações:").grid(row=3, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.sheet_novas, width=30).grid(row=3, column=1, sticky=tk.W, padx=(10, 0))
        
        # Arquivo de saída
        ttk.Label(main_frame, text="Arquivo de Saída (.xlsx):").grid(row=4, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.arquivo_saida, width=50).grid(row=4, column=1, sticky=(tk.W, tk.E), padx=(10, 5))
        ttk.Button(main_frame, text="Salvar Como", command=self.selecionar_arquivo_saida).grid(row=4, column=2, padx=(5, 0))
        
        # Botões
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=5, column=0, columnspan=3, pady=20)
        
        ttk.Button(button_frame, text="Gerar Relatório", command=self.gerar_relatorio, 
                  style="Accent.TButton").pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="Limpar", command=self.limpar_campos).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Sair", command=self.root.quit).pack(side=tk.LEFT, padx=10)
        
        # Área de log
        log_frame = ttk.LabelFrame(main_frame, text="Log de Processamento", padding="10")
        log_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(20, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        # Text widget com scrollbar
        text_frame = ttk.Frame(log_frame)
        text_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        text_frame.columnconfigure(0, weight=1)
        text_frame.rowconfigure(0, weight=1)
        
        self.log_text = tk.Text(text_frame, height=15, width=70, wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Configurar peso das linhas para expandir
        main_frame.rowconfigure(6, weight=1)
        
        self.log("Aplicação iniciada. Selecione o arquivo de entrada para começar.")
        
    def selecionar_arquivo_entrada(self):
        arquivo = filedialog.askopenfilename(
            title="Selecionar arquivo de entrada",
            filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")]
        )
        if arquivo:
            self.arquivo_entrada.set(arquivo)
            self.log(f"Arquivo de entrada selecionado: {os.path.basename(arquivo)}")
            
            # Sugerir nome para arquivo de saída
            if not self.arquivo_saida.get():
                caminho_base = os.path.splitext(arquivo)[0]
                arquivo_saida_sugerido = f"{caminho_base}_consolidado.xlsx"
                self.arquivo_saida.set(arquivo_saida_sugerido)
                
    def selecionar_arquivo_saida(self):
        arquivo = filedialog.asksaveasfilename(
            title="Salvar arquivo como",
            defaultextension=".xlsx",
            filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")]
        )
        if arquivo:
            self.arquivo_saida.set(arquivo)
            self.log(f"Arquivo de saída definido: {os.path.basename(arquivo)}")
            
    def limpar_campos(self):
        self.arquivo_entrada.set("")
        self.arquivo_saida.set("")
        self.sheet_pendencias.set("Pendências")
        self.sheet_novas.set("Sheet1")
        self.log_text.delete(1.0, tk.END)
        self.log("Campos limpos.")
        
    def log(self, mensagem):
        """Adiciona uma mensagem ao log"""
        self.log_text.insert(tk.END, f"{mensagem}\n")
        self.log_text.see(tk.END)
        self.root.update()
        
    def gerar_relatorio(self):
        """Gera o relatório consolidado"""
        # Validar campos obrigatórios
        if not self.arquivo_entrada.get():
            messagebox.showerror("Erro", "Selecione o arquivo de entrada!")
            return
            
        if not self.arquivo_saida.get():
            messagebox.showerror("Erro", "Defina o arquivo de saída!")
            return
            
        if not self.sheet_pendencias.get().strip():
            messagebox.showerror("Erro", "Informe o nome da sheet de pendências!")
            return
            
        if not self.sheet_novas.get().strip():
            messagebox.showerror("Erro", "Informe o nome da sheet de novas transações!")
            return
            
        # Verificar se o arquivo de entrada existe
        if not os.path.exists(self.arquivo_entrada.get()):
            messagebox.showerror("Erro", "Arquivo de entrada não encontrado!")
            return
            
        try:
            self.log("=" * 50)
            self.log("Iniciando processamento...")
            self.log(f"Arquivo de entrada: {os.path.basename(self.arquivo_entrada.get())}")
            self.log(f"Sheet de pendências: {self.sheet_pendencias.get()}")
            self.log(f"Sheet de novas transações: {self.sheet_novas.get()}")
            self.log(f"Arquivo de saída: {os.path.basename(self.arquivo_saida.get())}")
            self.log("")
            
            # Desabilitar botão durante processamento
            for widget in self.root.winfo_children():
                if isinstance(widget, ttk.Frame):
                    for child in widget.winfo_children():
                        if isinstance(child, ttk.Button) and child['text'] == 'Gerar Relatório':
                            child.configure(state='disabled')
            
            # Processar
            linhas_consolidadas = gerar_relatorio_consolidado(
                self.arquivo_entrada.get(),
                self.arquivo_saida.get(),
                self.sheet_pendencias.get().strip(),
                self.sheet_novas.get().strip()
            )
            
            self.log("✅ Processamento concluído com sucesso!")
            self.log(f"📊 Total de linhas consolidadas: {linhas_consolidadas}")
            self.log(f"💾 Arquivo salvo em: {self.arquivo_saida.get()}")
            
            messagebox.showinfo("Sucesso", 
                              f"Relatório gerado com sucesso!\n\n"
                              f"Linhas consolidadas: {linhas_consolidadas}\n"
                              f"Arquivo salvo em:\n{self.arquivo_saida.get()}")
            
        except Exception as e:
            error_msg = f"Erro durante o processamento: {str(e)}"
            self.log(f"❌ {error_msg}")
            messagebox.showerror("Erro", error_msg)
            
        finally:
            # Reabilitar botão
            for widget in self.root.winfo_children():
                if isinstance(widget, ttk.Frame):
                    for child in widget.winfo_children():
                        if isinstance(child, ttk.Button) and child['text'] == 'Gerar Relatório':
                            child.configure(state='normal')


def main():
    root = tk.Tk()
    
    # Configurar estilo
    style = ttk.Style()
    style.theme_use('clam')
    
    app = CashFlowApp(root)
    root.mainloop()


if __name__ == "__main__":
    main() 