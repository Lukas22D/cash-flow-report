import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import sys
from pathlib import Path

# Imports da nova arquitetura modular
from extractor.main import gerar_relatorio_consolidado


class CashFlowApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerador de Relatório Cash Flow - Consolidado")
        self.root.geometry("900x700")
        self.root.minsize(800, 600)
        
        # Variáveis
        self.arquivo_rel_sem_tratar = tk.StringVar()
        self.arquivo_pendencias_antigas = tk.StringVar()
        self.sheet_pendencias = tk.StringVar(value="Pendências")
        
        self.criar_interface()
        
    def criar_interface(self):
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="25")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configurar grid
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Título e descrição
        title_label = ttk.Label(main_frame, text="📊 Gerador de Relatório Cash Flow", 
                               font=("Arial", 18, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 5))
        
        subtitle_label = ttk.Label(main_frame, text="Consolidador de Pendências - Arquitetura Modular", 
                                  font=("Arial", 12), foreground="gray")
        subtitle_label.grid(row=1, column=0, columnspan=3, pady=(0, 25))
        
        # Seção 1: Arquivo Rel_sem_tratar (Novas Transações)
        section1_frame = ttk.LabelFrame(main_frame, text="📊 Rel_sem_tratar.xlsx (Novas Transações)", padding="15")
        section1_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        section1_frame.columnconfigure(0, weight=1)
        
        # Entry para mostrar arquivo selecionado
        self.rel_sem_tratar_entry = ttk.Entry(section1_frame, textvariable=self.arquivo_rel_sem_tratar, 
                                             font=("Arial", 10), state="readonly")
        self.rel_sem_tratar_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        
        # Botão procurar
        ttk.Button(section1_frame, text="📂 Procurar", 
                  command=self.selecionar_rel_sem_tratar).grid(row=0, column=1)
        
        # Informação adicional
        ttk.Label(section1_frame, text="ℹ️ Header na linha 2, numeração americana", 
                 font=("Arial", 9), foreground="gray").grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(5, 0))
        
        # Seção 2: Arquivo de Pendências Antigas
        section2_frame = ttk.LabelFrame(main_frame, text="📋 Pendências Antigas (Excel)", padding="15")
        section2_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        section2_frame.columnconfigure(0, weight=1)
        
        # Entry para mostrar arquivo selecionado
        self.pendencias_antigas_entry = ttk.Entry(section2_frame, textvariable=self.arquivo_pendencias_antigas, 
                                                 font=("Arial", 10), state="readonly")
        self.pendencias_antigas_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        
        # Botão procurar
        ttk.Button(section2_frame, text="📂 Procurar", 
                  command=self.selecionar_pendencias_antigas).grid(row=0, column=1)
        
        # Seção 3: Configuração da sheet de pendências
        section3_frame = ttk.LabelFrame(main_frame, text="⚙️ Configuração da Sheet", padding="15")
        section3_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        section3_frame.columnconfigure(1, weight=1)
        
        # Sheet de pendências
        ttk.Label(section3_frame, text="Nome da Sheet - Pendências Antigas:").grid(row=0, column=0, sticky=tk.W)
        pendencias_entry = ttk.Entry(section3_frame, textvariable=self.sheet_pendencias, 
                                    font=("Arial", 10), width=25)
        pendencias_entry.grid(row=0, column=1, sticky=tk.W, padx=(15, 0))
        
        # Seção 4: Ações
        action_frame = ttk.Frame(main_frame)
        action_frame.grid(row=5, column=0, columnspan=3, pady=20)
        
        # Botão principal (maior e mais destacado)
        self.gerar_btn = ttk.Button(action_frame, text="🚀 Gerar Relatório Consolidado", 
                                   command=self.gerar_relatorio, 
                                   style="Accent.TButton")
        self.gerar_btn.pack(side=tk.LEFT, padx=(0, 15), ipadx=20, ipady=5)
        
        # Botões secundários
        ttk.Button(action_frame, text="🧹 Limpar", command=self.limpar_campos).pack(side=tk.LEFT, padx=5)
        ttk.Button(action_frame, text="❌ Sair", command=self.root.quit).pack(side=tk.LEFT, padx=15)
        
        # Seção 5: Log de processamento
        log_frame = ttk.LabelFrame(main_frame, text="📜 Log de Processamento", padding="15")
        log_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(15, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        # Text widget com scrollbar
        text_frame = ttk.Frame(log_frame)
        text_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        text_frame.columnconfigure(0, weight=1)
        text_frame.rowconfigure(0, weight=1)
        
        self.log_text = tk.Text(text_frame, height=12, font=("Consolas", 9), wrap=tk.WORD,
                               bg="#f8f8f8", fg="#333333")
        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Configurar peso das linhas para expandir
        main_frame.rowconfigure(6, weight=1)
        
        self.log("✨ Aplicação iniciada com nova arquitetura modular!")
        self.log("🏗️ Componentes: Entidades | Extrator | Serviços | Saída")
        self.log("📝 Instruções:")
        self.log("   1. Selecione o arquivo 'Rel_sem_tratar.xlsx' (novas transações)")
        self.log("   2. Selecione o arquivo com pendências antigas")
        self.log("   3. Verifique o nome da sheet de pendências")
        self.log("   4. Clique em 'Gerar Relatório' para processar")
        self.log("")
        
    def selecionar_rel_sem_tratar(self):
        """Abre diálogo para selecionar arquivo Rel_sem_tratar.xlsx"""
        arquivo = filedialog.askopenfilename(
            title="📊 Selecionar Rel_sem_tratar.xlsx",
            filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")]
        )
        if arquivo:
            self.arquivo_rel_sem_tratar.set(arquivo)
            nome_arquivo = os.path.basename(arquivo)
            self.log(f"📊 Rel_sem_tratar selecionado: {nome_arquivo}")
    
    def selecionar_pendencias_antigas(self):
        """Abre diálogo para selecionar arquivo de pendências antigas"""
        arquivo = filedialog.askopenfilename(
            title="📋 Selecionar arquivo de pendências antigas",
            filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")]
        )
        if arquivo:
            self.arquivo_pendencias_antigas.set(arquivo)
            nome_arquivo = os.path.basename(arquivo)
            self.log(f"📋 Pendências antigas selecionadas: {nome_arquivo}")
            
    def limpar_campos(self):
        """Limpa todos os campos"""
        self.arquivo_rel_sem_tratar.set("")
        self.arquivo_pendencias_antigas.set("")
        self.sheet_pendencias.set("Pendências")
        self.log_text.delete(1.0, tk.END)
        self.log("🧹 Campos limpos. Pronto para novo processamento!")
        
    def log(self, mensagem):
        """Adiciona uma mensagem ao log com timestamp"""
        from datetime import datetime
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {mensagem}\n")
        self.log_text.see(tk.END)
        self.root.update()
        
    def gerar_relatorio(self):
        """Gera o relatório consolidado usando a nova arquitetura"""
        # Validar campos obrigatórios
        if not self.arquivo_rel_sem_tratar.get():
            messagebox.showerror("❌ Erro", "Por favor, selecione o arquivo Rel_sem_tratar.xlsx!")
            return
            
        if not self.arquivo_pendencias_antigas.get():
            messagebox.showerror("❌ Erro", "Por favor, selecione o arquivo de pendências antigas!")
            return
            
        if not self.sheet_pendencias.get().strip():
            messagebox.showerror("❌ Erro", "Informe o nome da sheet de pendências!")
            return
            
        # Verificar se os arquivos existem
        if not os.path.exists(self.arquivo_rel_sem_tratar.get()):
            messagebox.showerror("❌ Erro", "Arquivo Rel_sem_tratar não encontrado!")
            return
            
        if not os.path.exists(self.arquivo_pendencias_antigas.get()):
            messagebox.showerror("❌ Erro", "Arquivo de pendências antigas não encontrado!")
            return
            
        # Solicitar onde salvar o arquivo (apenas no momento da geração)
        nome_base = os.path.splitext(os.path.basename(self.arquivo_rel_sem_tratar.get()))[0]
        nome_sugerido = f"{nome_base}_consolidado.xlsx"
        
        arquivo_saida = filedialog.asksaveasfilename(
            title="💾 Salvar relatório consolidado como...",
            defaultextension=".xlsx",
            initialfile=nome_sugerido,
            filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")]
        )
        
        if not arquivo_saida:
            self.log("⚠️ Operação cancelada pelo usuário.")
            return
            
        try:
            self.log("=" * 60)
            self.log("🚀 INICIANDO PROCESSAMENTO - ARQUITETURA MODULAR")
            self.log(f"📊 Rel_sem_tratar: {os.path.basename(self.arquivo_rel_sem_tratar.get())}")
            self.log(f"📋 Pendências antigas: {os.path.basename(self.arquivo_pendencias_antigas.get())}")
            self.log(f"📋 Sheet de pendências: '{self.sheet_pendencias.get()}'")
            self.log(f"💾 Arquivo de saída: {os.path.basename(arquivo_saida)}")
            self.log("")
            
            # Desabilitar botão durante processamento
            self.gerar_btn.configure(state='disabled', text="⏳ Processando...")
            self.root.update()
            
            # Processar usando a nova arquitetura
            self.log("🔧 Etapa 1: Extraindo novas transações do Rel_sem_tratar...")
            self.log("🔧 Etapa 2: Extraindo pendências antigas...")
            self.log("🔧 Etapa 3: Executando serviço de conciliação...")
            self.log("🔧 Etapa 4: Salvando arquivo consolidado...")
            
            resultado = gerar_relatorio_consolidado(
                self.arquivo_rel_sem_tratar.get(),
                self.arquivo_pendencias_antigas.get(),
                arquivo_saida,
                self.sheet_pendencias.get().strip()
            )
            
            self.log("✅ PROCESSAMENTO CONCLUÍDO COM SUCESSO!")
            self.log(f"📊 Total de linhas consolidadas: {resultado['total_consolidadas']}")
            self.log(f"📈 Pendências preservadas: {resultado['pendencias_preservadas']}")
            self.log(f"📈 Novas pendências adicionadas: {resultado['novas_pendencias_adicionadas']}")
            self.log(f"📊 Resumo incluído: {'Sim' if resultado['tem_resumo'] else 'Não'}")
            self.log("")
            self.log("📊 RESUMO DE PENDÊNCIAS GERADO:")
            self.log(f"   • Departamentos processados: {resultado['total_departamentos']}")
            self.log(f"   • Pendências D1: {resultado['total_d1']}")
            self.log(f"   • Pendências >D+1: {resultado['total_d_mais_1']}")
            self.log(f"   • Pendências sem vencimento: {resultado['total_vazio']}")
            self.log(f"   • Total geral: {resultado['total_geral_absoluto']}")
            self.log(f"   • Dia útil de referência: {resultado['dia_util_referencia']}")
            self.log(f"💾 Arquivo salvo em: {arquivo_saida}")
            self.log("=" * 60)
            
            # Perguntar se quer abrir o diretório
            resposta = messagebox.askyesno("🎉 Sucesso!", 
                                         f"Relatório gerado com sucesso!\n\n"
                                         f"📊 CONCILIAÇÃO:\n"
                                         f"   • Linhas consolidadas: {resultado['total_consolidadas']}\n"
                                         f"   • Pendências preservadas: {resultado['pendencias_preservadas']}\n"
                                         f"   • Novas pendências: {resultado['novas_pendencias_adicionadas']}\n"
                                         f"   • Resumo original: {'Incluído' if resultado['tem_resumo'] else 'Não encontrado'}\n\n"
                                         f"📊 RESUMO DE PENDÊNCIAS:\n"
                                         f"   • Departamentos: {resultado['total_departamentos']}\n"
                                         f"   • D1: {resultado['total_d1']} | >D+1: {resultado['total_d_mais_1']} | Vazios: {resultado['total_vazio']}\n"
                                         f"   • Total: {resultado['total_geral_absoluto']}\n"
                                         f"   • Dia útil: {resultado['dia_util_referencia']}\n\n"
                                         f"💾 Arquivo salvo em:\n{arquivo_saida}\n\n"
                                         f"Deseja abrir a pasta onde o arquivo foi salvo?")
            
            if resposta:
                self._abrir_pasta(arquivo_saida)
            
        except Exception as e:
            error_msg = f"Erro durante o processamento: {str(e)}"
            self.log(f"❌ {error_msg}")
            messagebox.showerror("❌ Erro", error_msg)
            
        finally:
            # Reabilitar botão
            self.gerar_btn.configure(state='normal', text="🚀 Gerar Relatório Consolidado")
    
    def _abrir_pasta(self, arquivo_saida):
        """Abre a pasta onde o arquivo foi salvo"""
        try:
            import subprocess
            import platform
            
            pasta = os.path.dirname(arquivo_saida)
            sistema = platform.system()
            
            if sistema == "Windows":
                subprocess.run(f'explorer "{pasta}"', shell=True)
            elif sistema == "Darwin":  # macOS
                subprocess.run(f'open "{pasta}"', shell=True)
            else:  # Linux
                subprocess.run(f'xdg-open "{pasta}"', shell=True)
                
            self.log(f"📁 Pasta aberta: {pasta}")
        except Exception as e:
            self.log(f"⚠️ Não foi possível abrir a pasta: {e}")


def criar_aplicacao():
    """Função para criar e executar a aplicação"""
    try:
        # Usar Tk normal
        root = tk.Tk()
        
        # Configurar estilo
        style = ttk.Style()
        style.theme_use('clam')
        
        # Personalizar alguns estilos
        style.configure("Accent.TButton", font=("Arial", 11, "bold"))
        
        app = CashFlowApp(root)
        
        # Centralizar janela na tela
        root.update_idletasks()
        x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
        y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
        root.geometry(f"+{x}+{y}")
        
        return root, app
        
    except Exception as e:
        print(f"❌ Erro ao criar aplicação: {e}")
        return None, None


def main():
    """Função principal"""
    print("🚀 Iniciando Gerador de Relatório Cash Flow...")
    
    root, app = criar_aplicacao()
    if root and app:
        try:
            root.mainloop()
        except Exception as e:
            print(f"❌ Erro durante execução: {e}")
    else:
        print("❌ Não foi possível iniciar a aplicação")


if __name__ == "__main__":
    main() 