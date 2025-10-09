import os
from typing import Dict, Any
from extractor.excel_reader import extrair_pendencias, extrair_resumo
from extractor.rel_sem_tratar_reader import extrair_novas_transacoes_rel_sem_tratar
from extractor.depara_reader import extrair_responsaveis, extrair_departamentos
from services.conciliacao_service import ConciliacaoService
from services.resumo_service import ResumoService
from output.excel_writer import ExcelWriter


def gerar_relatorio_consolidado(caminho_rel_sem_tratar: str,
                               caminho_pendencias_antigas: str,
                               caminho_arquivo_saida: str,
                               sheet_pendencias: str = 'Pendências') -> Dict[str, Any]:
    """
    Função principal que orquestra o processo de geração do relatório consolidado.
    
    Args:
        caminho_rel_sem_tratar: Caminho para o arquivo Rel_sem_tratar.xlsx (novas transações)
        caminho_pendencias_antigas: Caminho para o arquivo com pendências antigas
        caminho_arquivo_saida: Caminho completo para salvar o arquivo gerado (.xlsx)
        sheet_pendencias: Nome da aba com as pendências existentes (padrão: 'Pendências')
        
    Returns:
        Dict[str, Any]: Dicionário com estatísticas do processamento
        
    Raises:
        FileNotFoundError: Se algum arquivo não for encontrado
        ValueError: Se as abas especificadas não existirem
        PermissionError: Se não conseguir salvar o arquivo de saída
        Exception: Outros erros durante o processamento
    """
    
    # 1. EXTRAÇÃO: Ler dados dos arquivos Excel
    # 1.1. Extrair novas transações do Rel_sem_tratar.xlsx
    novas_transacoes = extrair_novas_transacoes_rel_sem_tratar(caminho_rel_sem_tratar)
    
    # 1.2. Extrair pendências antigas do arquivo separado
    pendencias_existentes = extrair_pendencias(caminho_pendencias_antigas, sheet_pendencias)
    
    # 1.3. Tentar extrair resumo do arquivo de pendências antigas (opcional)
    try:
        df_resumo = extrair_resumo(caminho_pendencias_antigas)
    except:
        # Se não conseguir extrair resumo, usar DataFrame vazio
        import pandas as pd
        df_resumo = pd.DataFrame()
    
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
    
    # 2.1. PROCESSAMENTO: Gerar resumo das pendências consolidadas
    resumo_consolidado = ResumoService.gerar_resumo(pendencias_consolidadas)
    print(f"📊 Resumo gerado: {len(resumo_consolidado.itens)} departamentos processados")
    print(f"📅 Dia útil de referência: {resumo_consolidado.dia_util_referencia.strftime('%d/%m/%Y')}")
    
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
    
    # Adicionar estatísticas do resumo
    estatisticas_resumo = {
        'resumo_pendencias_gerado': True,
        'total_departamentos': len(resumo_consolidado.itens),
        'total_d1': resumo_consolidado.total_d1,
        'total_d_mais_1': resumo_consolidado.total_d_mais_1,
        'total_geral_absoluto': resumo_consolidado.total_geral_absoluto,
        'dia_util_referencia': resumo_consolidado.dia_util_referencia.strftime('%d/%m/%Y')
    }
    
    # Adicionar informações dos arquivos
    estatisticas.update({
        'arquivo_rel_sem_tratar': caminho_rel_sem_tratar,
        'arquivo_pendencias_antigas': caminho_pendencias_antigas,
        'arquivo_saida': caminho_arquivo_saida,
        'sheet_pendencias': sheet_pendencias,
        'tem_resumo': not df_resumo.empty,
        **estatisticas_resumo
    })
    
    return estatisticas



if __name__ == '__main__':
    # Exemplo de uso
    try:
        # DePara é carregado automaticamente do caminho fixo
        resultado = gerar_relatorio_consolidado(
            "docs/Rel_sem_tratar.xlsx", 
            "docs/Pendencias_Antigas_Exemplo.xlsx",
            "Rel_cons.xlsx"
        )
        
        print("=" * 60)
        print("🎉 RELATÓRIO GERADO COM SUCESSO!")
        print("=" * 60)
        print(f"📊 Rel_sem_tratar: {resultado['arquivo_rel_sem_tratar']}")
        print(f"📋 Pendências antigas: {resultado['arquivo_pendencias_antigas']}")
        print(f"📋 Sheet pendências: {resultado['sheet_pendencias']}")
        print(f"💾 Arquivo de saída: {resultado['arquivo_saida']}")
        print(f"📊 Resumo incluído: {'Sim' if resultado['tem_resumo'] else 'Não'}")
        print()
        print("📈 ESTATÍSTICAS DE CONCILIAÇÃO:")
        print(f"   • Pendências existentes: {resultado['total_pendencias_existentes']}")
        print(f"   • Novas transações: {resultado['total_novas_transacoes']}")
        print(f"   • Total consolidadas: {resultado['total_consolidadas']}")
        print(f"   • Pendências preservadas: {resultado['pendencias_preservadas']}")
        print(f"   • Novas pendências adicionadas: {resultado['novas_pendencias_adicionadas']}")
        print()
        print("📊 RESUMO DE PENDÊNCIAS:")
        print(f"   • Departamentos processados: {resultado['total_departamentos']}")
        print(f"   • Pendências D1: {resultado['total_d1']}")
        print(f"   • Pendências >D+1: {resultado['total_d_mais_1']}")
        print(f"   • Total geral: {resultado['total_geral_absoluto']}")
        print(f"   • Dia útil de referência: {resultado['dia_util_referencia']}")
        print("=" * 60)
        
    except Exception as e:
        print(f"❌ Erro: {e}")

