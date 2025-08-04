#!/usr/bin/env python3
"""
Script de inicialização para o Gerador de Relatório Cash Flow.
Execute este arquivo para iniciar a aplicação gráfica.
"""

import sys
import os
from pathlib import Path

# Adicionar src ao path
src_path = Path(__file__).parent / "src"
sys.path.insert(0, str(src_path))

def main():
    """Função principal de inicialização"""
    print("🚀 Gerador de Relatório Cash Flow - Consolidado")
    print("📁 Iniciando aplicação...")
    
    try:
        # Verificar se estamos no diretório correto
        if not src_path.exists():
            print("❌ Erro: Diretório 'src' não encontrado!")
            print("   Execute este script a partir do diretório raiz do projeto.")
            return
        
        # Importar e executar a aplicação
        from app import main as app_main
        app_main()
        
    except ImportError as e:
        print(f"❌ Erro de importação: {e}")
        print("   Verifique se todas as dependências estão instaladas:")
        print("   pip install pandas openpyxl tkinterdnd2")
    except Exception as e:
        print(f"❌ Erro inesperado: {e}")

if __name__ == "__main__":
    main() 