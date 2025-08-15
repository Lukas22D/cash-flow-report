import os
from typing import Dict, Any
from extractor.excel_reader import extrair_pendencias, extrair_transacoes, extrair_resumo
from extractor.depara_reader import extrair_responsaveis, extrair_departamentos
from services.conciliacao_service import ConciliacaoService
from output.excel_writer import ExcelWriter


def gerar_relatorio_consolidado(caminho_arquivo_entrada: str, 
                               caminho_arquivo_saida: str, 
                               sheet_pendencias: str = 'Pendências', 
                               sheet_novas: str = 'Sheet1') -> Dict[str, Any]:
    """
    Função principal que orquestra o processo de geração do relatório consolidado.
    
    Args:
        caminho_arquivo_entrada: Caminho completo para o arquivo de entrada (.xlsx)
        caminho_arquivo_saida: Caminho completo para salvar o arquivo gerado (.xlsx)
        sheet_pendencias: Nome da aba com as pendências existentes (padrão: 'Pendências')
        sheet_novas: Nome da aba com as novas transações (padrão: 'Sheet1')
        
    Returns:
        Dict[str, Any]: Dicionário com estatísticas do processamento
        
    Raises:
        FileNotFoundError: Se o arquivo de entrada não for encontrado
        ValueError: Se as abas especificadas não existirem
        PermissionError: Se não conseguir salvar o arquivo de saída
        Exception: Outros erros durante o processamento
    """
    
    # 1. EXTRAÇÃO: Ler dados do Excel
    pendencias_existentes = extrair_pendencias(caminho_arquivo_entrada, sheet_pendencias)
    novas_transacoes = extrair_transacoes(caminho_arquivo_entrada, sheet_novas)
    df_resumo = extrair_resumo(caminho_arquivo_entrada)
    
    # 1.1. EXTRAÇÃO: Ler dados do DePara (caminho fixo)
    responsaveis = []
    departamentos = []
    
    # Definir caminho fixo para o arquivo DePara
    diretorio_atual = os.path.dirname(os.path.abspath(__file__))
    caminho_depara = os.path.join(diretorio_atual, "depara", "DePara-CashFlow.xlsx")
    
    try:
        responsaveis = extrair_responsaveis(caminho_depara)
        departamentos = extrair_departamentos(caminho_depara)
        print(f"✅ DePara carregado: {len(responsaveis)} responsáveis, {len(departamentos)} departamentos")
    except (FileNotFoundError, ValueError) as e:
        print(f"⚠️ Aviso: Não foi possível carregar DePara de '{caminho_depara}' ({e}). Continuando sem enriquecimento de dados.")
    
    # 2. PROCESSAMENTO: Consolidar pendências usando a lógica de negócio
    pendencias_consolidadas = ConciliacaoService.consolidar_pendencias(
        pendencias_existentes, 
        novas_transacoes,
        responsaveis,
        departamentos
    )
    
    # 3. SAÍDA: Salvar arquivo consolidado
    ExcelWriter.salvar_relatorio_consolidado(
        pendencias_consolidadas,
        df_resumo,
        caminho_arquivo_saida
    )
    
    # 4. ESTATÍSTICAS: Calcular e retornar estatísticas do processamento
    estatisticas = ConciliacaoService.obter_estatisticas_consolidacao(
        pendencias_existentes,
        novas_transacoes, 
        pendencias_consolidadas
    )
    
    # Adicionar informações dos arquivos
    estatisticas.update({
        'arquivo_entrada': caminho_arquivo_entrada,
        'arquivo_saida': caminho_arquivo_saida,
        'sheet_pendencias': sheet_pendencias,
        'sheet_novas': sheet_novas,
        'tem_resumo': not df_resumo.empty
    })
    
    return estatisticas



if __name__ == '__main__':
    # Exemplo de uso
    try:
        # DePara é carregado automaticamente do caminho fixo
        resultado = gerar_relatorio_consolidado("Rel.xlsx", "Rel_cons.xlsx")
        
        print("=" * 60)
        print("🎉 RELATÓRIO GERADO COM SUCESSO!")
        print("=" * 60)
        print(f"📂 Arquivo de entrada: {resultado['arquivo_entrada']}")
        print(f"📋 Sheet pendências: {resultado['sheet_pendencias']}")
        print(f"📋 Sheet novas transações: {resultado['sheet_novas']}")
        print(f"💾 Arquivo de saída: {resultado['arquivo_saida']}")
        print(f"📊 Resumo incluído: {'Sim' if resultado['tem_resumo'] else 'Não'}")
        print()
        print("📈 ESTATÍSTICAS:")
        print(f"   • Pendências existentes: {resultado['total_pendencias_existentes']}")
        print(f"   • Novas transações: {resultado['total_novas_transacoes']}")
        print(f"   • Total consolidadas: {resultado['total_consolidadas']}")
        print(f"   • Pendências preservadas: {resultado['pendencias_preservadas']}")
        print(f"   • Novas pendências adicionadas: {resultado['novas_pendencias_adicionadas']}")
        print("=" * 60)
        
    except Exception as e:
        print(f"❌ Erro: {e}")

