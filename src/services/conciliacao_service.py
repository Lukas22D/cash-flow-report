from typing import List, Dict
from entities.pendencia import Pendencia


class ConciliacaoService:
    """
    Serviço responsável pela conciliação entre pendências existentes e novas transações.
    
    Implementa a lógica de negócio:
    - Para cada transação nova, verifica se existe pendência com mesma chave
    - Se existe: usa a pendência existente (preserva dados extras como Responsável, etc.)
    - Se não existe: usa a nova transação como nova pendência
    """
    
    @staticmethod
    def consolidar_pendencias(pendencias_existentes: List[Pendencia], 
                            novas_transacoes: List[Pendencia]) -> List[Pendencia]:
        """
        Consolida pendências seguindo a lógica de negócio.
        
        Args:
            pendencias_existentes: Lista de pendências já existentes
            novas_transacoes: Lista de novas transações a serem processadas
            
        Returns:
            List[Pendencia]: Lista consolidada de pendências
        """
        # Criar dicionário de pendências existentes por chave para busca rápida
        pendencias_dict = ConciliacaoService._criar_dicionario_pendencias(pendencias_existentes)
        
        # Lista resultado
        pendencias_consolidadas = []
        
        # Para cada nova transação, decidir se usar pendência existente ou nova
        for transacao in novas_transacoes:
            chave = transacao.get_chave_reconciliacao()
            
            if chave in pendencias_dict:
                # Se existe pendência com a mesma chave, usar a pendência existente
                pendencia_existente = pendencias_dict[chave]
                pendencias_consolidadas.append(pendencia_existente)
            else:
                # Se não existe, usar a nova transação como nova pendência
                pendencias_consolidadas.append(transacao)
        
        return pendencias_consolidadas
    
    @staticmethod
    def _criar_dicionario_pendencias(pendencias: List[Pendencia]) -> Dict[str, Pendencia]:
        """
        Cria um dicionário de pendências indexado pela chave de reconciliação.
        
        Args:
            pendencias: Lista de pendências
            
        Returns:
            Dict[str, Pendencia]: Dicionário chave -> pendência
        """
        pendencias_dict = {}
        
        for pendencia in pendencias:
            chave = pendencia.get_chave_reconciliacao()
            # Se houver chaves duplicadas, manter a primeira
            if chave not in pendencias_dict:
                pendencias_dict[chave] = pendencia
        
        return pendencias_dict
    
    @staticmethod
    def obter_estatisticas_consolidacao(pendencias_existentes: List[Pendencia],
                                      novas_transacoes: List[Pendencia],
                                      pendencias_consolidadas: List[Pendencia]) -> Dict[str, int]:
        """
        Calcula estatísticas do processo de consolidação.
        
        Args:
            pendencias_existentes: Lista de pendências existentes
            novas_transacoes: Lista de novas transações
            pendencias_consolidadas: Lista consolidada resultado
            
        Returns:
            Dict[str, int]: Estatísticas do processo
        """
        # Criar sets de chaves para análise
        chaves_existentes = {p.get_chave_reconciliacao() for p in pendencias_existentes}
        chaves_novas = {t.get_chave_reconciliacao() for t in novas_transacoes}
        
        # Calcular intersecções
        chaves_comuns = chaves_existentes & chaves_novas
        chaves_apenas_novas = chaves_novas - chaves_existentes
        
        return {
            'total_pendencias_existentes': len(pendencias_existentes),
            'total_novas_transacoes': len(novas_transacoes),
            'total_consolidadas': len(pendencias_consolidadas),
            'pendencias_preservadas': len(chaves_comuns),
            'novas_pendencias_adicionadas': len(chaves_apenas_novas),
            'chaves_unicas_existentes': len(chaves_existentes),
            'chaves_unicas_novas': len(chaves_novas)
        } 