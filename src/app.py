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
        self.arquivo_entrada = tk.StringVar()
        self.sheet_pendencias = tk.StringVar(value="Pendências")
        self.sheet_novas = tk.StringVar(value="Sheet1")
        
        # Verificar se drag and drop está disponível
        self.drag_drop_disponivel = self._verificar_drag_drop()
        
        self.criar_interface()
        
    def _verificar_drag_drop(self):
        """Verifica se tkinterdnd2 está disponível para evitar falha de segmentação"""
        try:
            from tkinterdnd2 import DND_FILES, TkinterDnD
            return True
        except ImportError:
            return False
        
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
        
        # Seção 1: Arquivo de entrada
        section1_frame = ttk.LabelFrame(main_frame, text="📁 Arquivo de Entrada", padding="15")
        section1_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        section1_frame.columnconfigure(0, weight=1)
        
        # Área de seleção de arquivo
        if self.drag_drop_disponivel:
            self._criar_area_drag_drop(section1_frame)
        else:
            self._criar_area_simples(section1_frame)
        
        # Entry para mostrar arquivo selecionado
        self.arquivo_entry = ttk.Entry(section1_frame, textvariable=self.arquivo_entrada, 
                                      font=("Arial", 10), state="readonly")
        self.arquivo_entry.grid(row=1, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        
        # Botão procurar
        ttk.Button(section1_frame, text="📂 Procurar", 
                  command=self.selecionar_arquivo_entrada).grid(row=1, column=1)
        
        # Seção 2: Configuração das sheets
        section2_frame = ttk.LabelFrame(main_frame, text="📋 Configuração das Planilhas", padding="15")
        section2_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        section2_frame.columnconfigure(1, weight=1)
        
        # Sheet de pendências
        ttk.Label(section2_frame, text="Nome da Sheet - Pendências Existentes:").grid(row=0, column=0, sticky=tk.W, pady=(0, 10))
        pendencias_entry = ttk.Entry(section2_frame, textvariable=self.sheet_pendencias, 
                                    font=("Arial", 10), width=25)
        pendencias_entry.grid(row=0, column=1, sticky=tk.W, padx=(15, 0), pady=(0, 10))
        
        # Sheet de novas transações
        ttk.Label(section2_frame, text="Nome da Sheet - Novas Transações:").grid(row=1, column=0, sticky=tk.W)
        novas_entry = ttk.Entry(section2_frame, textvariable=self.sheet_novas, 
                               font=("Arial", 10), width=25)
        novas_entry.grid(row=1, column=1, sticky=tk.W, padx=(15, 0))
        
        # Seção 3: Ações
        action_frame = ttk.Frame(main_frame)
        action_frame.grid(row=4, column=0, columnspan=3, pady=20)
        
        # Botão principal (maior e mais destacado)
        self.gerar_btn = ttk.Button(action_frame, text="🚀 Gerar Relatório Consolidado", 
                                   command=self.gerar_relatorio, 
                                   style="Accent.TButton")
        self.gerar_btn.pack(side=tk.LEFT, padx=(0, 15), ipadx=20, ipady=5)
        
        # Botões secundários
        ttk.Button(action_frame, text="🧹 Limpar", command=self.limpar_campos).pack(side=tk.LEFT, padx=5)
        ttk.Button(action_frame, text="❌ Sair", command=self.root.quit).pack(side=tk.LEFT, padx=15)
        
        # Seção 4: Log de processamento
        log_frame = ttk.LabelFrame(main_frame, text="📜 Log de Processamento", padding="15")
        log_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(15, 0))
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
        main_frame.rowconfigure(5, weight=1)
        
        self.log("✨ Aplicação iniciada com nova arquitetura modular!")
        self.log("🏗️ Componentes: Entidades | Extrator | Serviços | Saída")
        if not self.drag_drop_disponivel:
            self.log("⚠️ Drag & Drop não disponível (tkinterdnd2 não instalado)")
        self.log("📝 Instruções:")
        self.log("   1. Selecione o arquivo Excel usando o botão 'Procurar'")
        self.log("   2. Verifique os nomes das sheets (planilhas)")
        self.log("   3. Clique em 'Gerar Relatório' para processar")
        self.log("")
        
    def _criar_area_drag_drop(self, parent):
        """Cria área com funcionalidade de drag and drop"""
        try:
            from tkinterdnd2 import DND_FILES
            
            # Área de drag and drop
            self.drop_frame = tk.Frame(parent, height=100, bg="#f0f0f0", relief="ridge", bd=2)
            self.drop_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
            self.drop_frame.columnconfigure(0, weight=1)
            
            # Configurar drag and drop
            self.drop_frame.drop_target_register(DND_FILES)
            self.drop_frame.dnd_bind('<<Drop>>', self.on_drop)
            
            # Label informativo na área de drop
            self.drop_label = tk.Label(self.drop_frame, 
                                      text="🎯 Arraste o arquivo Excel (.xlsx) aqui\nou clique em 'Procurar' para selecionar",
                                      bg="#f0f0f0", font=("Arial", 11), fg="gray")
            self.drop_label.place(relx=0.5, rely=0.5, anchor="center")
        except Exception as e:
            self.log(f"⚠️ Erro ao configurar drag & drop: {e}")
            self._criar_area_simples(parent)
            
    def _criar_area_simples(self, parent):
        """Cria área simples sem drag and drop"""
        # Área simples
        self.drop_frame = tk.Frame(parent, height=100, bg="#f0f0f0", relief="ridge", bd=2)
        self.drop_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        self.drop_frame.columnconfigure(0, weight=1)
        
        # Label informativo
        self.drop_label = tk.Label(self.drop_frame, 
                                  text="📂 Clique em 'Procurar' para selecionar o arquivo Excel (.xlsx)",
                                  bg="#f0f0f0", font=("Arial", 11), fg="gray")
        self.drop_label.place(relx=0.5, rely=0.5, anchor="center")
    
    def on_drop(self, event):
        """Manipula o evento de arrastar e soltar arquivo"""
        if not self.drag_drop_disponivel:
            return
            
        try:
            arquivos = event.data.split()
            if arquivos:
                arquivo = arquivos[0].strip('{}')  # Remove chaves se houver
                if arquivo.lower().endswith('.xlsx'):
                    self.arquivo_entrada.set(arquivo)
                    nome_arquivo = os.path.basename(arquivo)
                    self.log(f"📁 Arquivo arrastado: {nome_arquivo}")
                    self.atualizar_drop_label(nome_arquivo)
                else:
                    messagebox.showwarning("Aviso", "Por favor, selecione um arquivo Excel (.xlsx)")
        except Exception as e:
            self.log(f"⚠️ Erro no drag & drop: {e}")
                
    def atualizar_drop_label(self, nome_arquivo):
        """Atualiza o label da área de drop com o nome do arquivo"""
        self.drop_label.config(text=f"✅ Arquivo selecionado:\n{nome_arquivo}", 
                              fg="green", font=("Arial", 10, "bold"))
        
    def selecionar_arquivo_entrada(self):
        """Abre diálogo para selecionar arquivo de entrada"""
        arquivo = filedialog.askopenfilename(
            title="📂 Selecionar arquivo Excel",
            filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")]
        )
        if arquivo:
            self.arquivo_entrada.set(arquivo)
            nome_arquivo = os.path.basename(arquivo)
            self.log(f"📁 Arquivo selecionado: {nome_arquivo}")
            self.atualizar_drop_label(nome_arquivo)
            
    def limpar_campos(self):
        """Limpa todos os campos"""
        self.arquivo_entrada.set("")
        self.sheet_pendencias.set("Pendências")
        self.sheet_novas.set("Sheet1")
        self.log_text.delete(1.0, tk.END)
        
        if self.drag_drop_disponivel:
            texto_drop = "🎯 Arraste o arquivo Excel (.xlsx) aqui\nou clique em 'Procurar' para selecionar"
        else:
            texto_drop = "📂 Clique em 'Procurar' para selecionar o arquivo Excel (.xlsx)"
            
        self.drop_label.config(text=texto_drop, fg="gray", font=("Arial", 11))
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
        if not self.arquivo_entrada.get():
            messagebox.showerror("❌ Erro", "Por favor, selecione o arquivo de entrada!")
            return
            
        if not self.sheet_pendencias.get().strip():
            messagebox.showerror("❌ Erro", "Informe o nome da sheet de pendências!")
            return
            
        if not self.sheet_novas.get().strip():
            messagebox.showerror("❌ Erro", "Informe o nome da sheet de novas transações!")
            return
            
        # Verificar se o arquivo de entrada existe
        if not os.path.exists(self.arquivo_entrada.get()):
            messagebox.showerror("❌ Erro", "Arquivo de entrada não encontrado!")
            return
            
        # Solicitar onde salvar o arquivo (apenas no momento da geração)
        nome_base = os.path.splitext(os.path.basename(self.arquivo_entrada.get()))[0]
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
            self.log(f"📂 Arquivo de entrada: {os.path.basename(self.arquivo_entrada.get())}")
            self.log(f"📋 Sheet de pendências: '{self.sheet_pendencias.get()}'")
            self.log(f"📋 Sheet de novas transações: '{self.sheet_novas.get()}'")
            self.log(f"💾 Arquivo de saída: {os.path.basename(arquivo_saida)}")
            self.log("")
            
            # Desabilitar botão durante processamento
            self.gerar_btn.configure(state='disabled', text="⏳ Processando...")
            self.root.update()
            
            # Processar usando a nova arquitetura
            self.log("🔧 Etapa 1: Extraindo dados do Excel...")
            self.log("🔧 Etapa 2: Executando serviço de conciliação...")
            self.log("🔧 Etapa 3: Salvando arquivo consolidado...")
            
            resultado = gerar_relatorio_consolidado(
                self.arquivo_entrada.get(),
                arquivo_saida,
                self.sheet_pendencias.get().strip(),
                self.sheet_novas.get().strip()
            )
            
            self.log("✅ PROCESSAMENTO CONCLUÍDO COM SUCESSO!")
            self.log(f"📊 Total de linhas consolidadas: {resultado['total_consolidadas']}")
            self.log(f"📈 Pendências preservadas: {resultado['pendencias_preservadas']}")
            self.log(f"📈 Novas pendências adicionadas: {resultado['novas_pendencias_adicionadas']}")
            self.log(f"📊 Resumo incluído: {'Sim' if resultado['tem_resumo'] else 'Não'}")
            self.log(f"💾 Arquivo salvo em: {arquivo_saida}")
            self.log("=" * 60)
            
            # Perguntar se quer abrir o diretório
            resposta = messagebox.askyesno("🎉 Sucesso!", 
                                         f"Relatório gerado com sucesso!\n\n"
                                         f"📊 Linhas consolidadas: {resultado['total_consolidadas']}\n"
                                         f"📈 Pendências preservadas: {resultado['pendencias_preservadas']}\n"
                                         f"📈 Novas pendências: {resultado['novas_pendencias_adicionadas']}\n"
                                         f"📊 Resumo: {'Incluído' if resultado['tem_resumo'] else 'Não encontrado'}\n"
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
        # Tentar usar TkinterDnD se disponível
        try:
            from tkinterdnd2 import TkinterDnD
            root = TkinterDnD.Tk()
        except ImportError:
            # Fallback para Tk normal
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