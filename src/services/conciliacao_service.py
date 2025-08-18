from typing import List, Dict
from datetime import datetime, date, timedelta
from entities.pendencia import Pendencia
from entities.responsavel import Responsavel
from entities.departamento import Departamento


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
                            novas_transacoes: List[Pendencia],
                            responsaveis: List[Responsavel] = None,
                            departamentos: List[Departamento] = None) -> List[Pendencia]:
        """
        Consolida pendências seguindo a lógica de negócio.
        
        Args:
            pendencias_existentes: Lista de pendências já existentes
            novas_transacoes: Lista de novas transações a serem processadas
            responsaveis: Lista de responsáveis do DePara-CashFlow (opcional)
            departamentos: Lista de departamentos do DePara-CashFlow (opcional)
            
        Returns:
            List[Pendencia]: Lista consolidada de pendências
        """
        # Criar dicionário de pendências existentes por chave para busca rápida
        pendencias_dict = ConciliacaoService._criar_dicionario_pendencias(pendencias_existentes)
        
        # Criar dicionários de lookup para responsáveis e departamentos
        responsaveis_dict = ConciliacaoService._criar_dicionario_responsaveis(responsaveis or [])
        departamentos_dict = ConciliacaoService._criar_dicionario_departamentos(departamentos or [])
        
        # Lista resultado
        pendencias_consolidadas = []
        
        # Para cada nova transação, decidir se usar pendência existente ou nova
        for transacao in novas_transacoes:
            chave = transacao.get_chave_reconciliacao()
            
            if chave in pendencias_dict:
                # Se existe pendência com a mesma chave, usar a pendência existente
                pendencia_existente = pendencias_dict[chave]
                pendencia_final = pendencia_existente
            else:
                # Se não existe, usar a nova transação como nova pendência
                pendencia_final = transacao
            
            # Aplicar regras de negócio para preencher RESPONSAVEL, DEPARTAMENTO e VENCIMENTO
            pendencia_final = ConciliacaoService._aplicar_regras_negocio(
                pendencia_final, responsaveis_dict, departamentos_dict
            )
            
            pendencias_consolidadas.append(pendencia_final)
        
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
    def _criar_dicionario_responsaveis(responsaveis: List[Responsavel]) -> Dict[str, Responsavel]:
        """
        Cria um dicionário de responsáveis indexado pela chave de identificação.
        Chave: NOME_BANCO + INFORMACAO_ADICIONAL + TIPO_TRANSACAO
        
        Args:
            responsaveis: Lista de responsáveis
            
        Returns:
            Dict[str, Responsavel]: Dicionário chave -> responsável
        """
        responsaveis_dict = {}
        
        for responsavel in responsaveis:
            chave = responsavel.get_chave_identificacao()
            # Se houver chaves duplicadas, manter a primeira
            if chave not in responsaveis_dict:
                responsaveis_dict[chave] = responsavel
        
        return responsaveis_dict
    
    @staticmethod
    def _criar_dicionario_departamentos(departamentos: List[Departamento]) -> Dict[str, Departamento]:
        """
        Cria um dicionário de departamentos indexado pelo responsável.
        
        Args:
            departamentos: Lista de departamentos
            
        Returns:
            Dict[str, Departamento]: Dicionário responsável -> departamento
        """
        departamentos_dict = {}
        
        for departamento in departamentos:
            responsavel = departamento.RESPONSAVEL
            if responsavel and responsavel not in departamentos_dict:
                departamentos_dict[responsavel] = departamento
        
        return departamentos_dict
    
    @staticmethod
    def _aplicar_regras_negocio(pendencia: Pendencia, 
                               responsaveis_dict: Dict[str, Responsavel],
                               departamentos_dict: Dict[str, Departamento]) -> Pendencia:
        """
        Aplica as regras de negócio para preencher RESPONSAVEL, DEPARTAMENTO e VENCIMENTO da pendência.
        
        Regras:
        1. RESPONSAVEL: baseado na combinação NOME_BANCO + INFORMACAO_ADICIONAL + TIPO_TRANSACAO
        2. DEPARTAMENTO: baseado no RESPONSAVEL encontrado
        3. VENCIMENTO: baseado na comparação entre DATA_EXTRATO e último dia útil anterior
           - Se DATA_EXTRATO < último dia útil anterior: VENCIMENTO = ">D+1"
           - Se DATA_EXTRATO >= último dia útil anterior: VENCIMENTO = "D1"
           - Considera apenas dias úteis (Segunda a Sexta)
        
        Args:
            pendencia: Pendência a ser processada
            responsaveis_dict: Dicionário de responsáveis
            departamentos_dict: Dicionário de departamentos
            
        Returns:
            Pendencia: Pendência com campos atualizados
        """
        # 1ª Regra: Definir RESPONSAVEL
        # Criar chave de busca baseada nos campos da pendência
        nome_banco = str(pendencia.NOME_BANCO) if pendencia.NOME_BANCO is not None else ""
        info_adicional = str(pendencia.INFORMACAO_ADICIONAL) if pendencia.INFORMACAO_ADICIONAL is not None else ""
        tipo_transacao = str(pendencia.TIPO_TRANSACAO) if pendencia.TIPO_TRANSACAO is not None else ""
        
        chave_responsavel = f"{nome_banco}{info_adicional}{tipo_transacao}"
        
        if chave_responsavel in responsaveis_dict:
            responsavel_encontrado = responsaveis_dict[chave_responsavel]
            pendencia.RESPONSAVEL = responsavel_encontrado.RESPONSAVEL
        
        # 2ª Regra: Definir DEPARTAMENTO
        # Só processa se o RESPONSAVEL foi definido
        if pendencia.RESPONSAVEL and pendencia.RESPONSAVEL in departamentos_dict:
            departamento_encontrado = departamentos_dict[pendencia.RESPONSAVEL]
            pendencia.DEPARTAMENTO = departamento_encontrado.AREA
        
        # 3ª Regra: Definir VENCIMENTO
        # Baseado na comparação entre DATA_EXTRATO e data atual
        pendencia.VENCIMENTO = ConciliacaoService._calcular_vencimento(pendencia.DATA_EXTRATO)
        
        return pendencia
    
    @staticmethod
    def _calcular_vencimento(data_extrato) -> str:
        """
        Calcula o vencimento baseado na comparação entre DATA_EXTRATO e (último dia útil anterior).
        
        Regras:
        - Se DATA_EXTRATO < (último dia útil anterior): VENCIMENTO = ">D+1"
        - Se DATA_EXTRATO >= (último dia útil anterior): VENCIMENTO = "D1"
        
        Dias úteis: Segunda a Sexta (exclui sábados e domingos)
        Exemplos:
        - Se executado na Segunda: compara com Sexta anterior
        - Se executado na Terça: compara com Segunda anterior
        - Se executado no Sábado: compara com Sexta anterior
        
        Args:
            data_extrato: Data do extrato (pode ser datetime, date ou None)
            
        Returns:
            str: String do vencimento ("D1" ou ">D+1")
        """
        if data_extrato is None:
            # Se não há data do extrato, considera como vencido
            return ">D+1"
        
        # Obter data de referência (último dia útil anterior)
        data_referencia = ConciliacaoService._obter_ultimo_dia_util_anterior()
        
        # Converter data_extrato para date se necessário
        if isinstance(data_extrato, datetime):
            data_extrato = data_extrato.date()
        elif isinstance(data_extrato, str):
            try:
                # Tentar converter string para date (formato comum: YYYY-MM-DD)
                data_extrato = datetime.strptime(data_extrato, '%Y-%m-%d').date()
            except ValueError:
                try:
                    # Tentar formato brasileiro: DD/MM/YYYY
                    data_extrato = datetime.strptime(data_extrato, '%d/%m/%Y').date()
                except ValueError:
                    # Se não conseguir converter, considera como vencido
                    return ">D+1"
        
        # Aplicar regra de negócio
        if data_extrato < data_referencia:
            return ">D+1"  # Data do extrato é anterior ao último dia útil anterior
        else:
            return "D1"    # Data do extrato é igual ou posterior ao último dia útil anterior
    
    @staticmethod
    def _obter_ultimo_dia_util_anterior() -> date:
        """
        Obtém o último dia útil anterior à data atual.
        
        Lógica:
        - Segunda-feira (0): retorna Sexta anterior (3 dias atrás)
        - Terça-feira (1): retorna Segunda anterior (1 dia atrás)
        - Quarta-feira (2): retorna Terça anterior (1 dia atrás)
        - Quinta-feira (3): retorna Quarta anterior (1 dia atrás)
        - Sexta-feira (4): retorna Quinta anterior (1 dia atrás)
        - Sábado (5): retorna Sexta anterior (1 dia atrás)
        - Domingo (6): retorna Sexta anterior (2 dias atrás)
        
        Returns:
            date: Data do último dia útil anterior
        """
        data_atual = date.today()
        dia_semana = data_atual.weekday()  # 0=Segunda, 1=Terça, ..., 6=Domingo
        
        if dia_semana == 0:  # Segunda-feira
            # Voltar para sexta-feira anterior (3 dias atrás)
            dias_para_voltar = 3
        elif dia_semana == 6:  # Domingo
            # Voltar para sexta-feira anterior (2 dias atrás)
            dias_para_voltar = 2
        else:  # Terça a Sábado
            # Voltar 1 dia (dia útil anterior)
            dias_para_voltar = 1
        
        return data_atual - timedelta(days=dias_para_voltar)
    
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