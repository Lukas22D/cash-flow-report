from typing import List, Dict, Tuple
from datetime import datetime, date, timedelta
from dataclasses import dataclass
from entities.pendencia import Pendencia


@dataclass
class ResumoItem:
    """
    Representa um item do resumo por departamento.
    Conforme especificação do README_Resumo_Pivot.md
    """
    departamento: str
    d1: int  # Pendências com vencimento D1
    d_mais_1: int  # Pendências com vencimento >D+1
    total_geral: int  # Total horizontal (D1 + >D+1)


@dataclass
class ResumoConsolidado:
    """
    Representa o resumo consolidado completo.
    Conforme especificação do README_Resumo_Pivot.md
    """
    itens: List[ResumoItem]
    total_d1: int  # Total vertical D1 de todos os departamentos
    total_d_mais_1: int  # Total vertical >D+1 de todos os departamentos
    total_geral_absoluto: int  # Total geral absoluto
    dia_util_referencia: date  # Dia útil usado como referência (max de DATA_EXTRATO)


class ResumoService:
    """
    Serviço responsável por gerar resumo estruturado das pendências após conciliação.
    
    Baseado na especificação do README_Resumo_Pivot.md
    
    Gera relatório com header: Departamento | D1 | >D+1 | Total Geral
    
    Regras de negócio:
    1. Filtro de status: considerar apenas linhas com STATUS == "Não Reconciliada"
    2. Categorias de vencimento: considerar explicitamente as colunas D1 e >D+1. Se uma categoria não existir no dia, preencher com 0 (zero)
    3. Agregação: contagem de linhas (pode contar qualquer coluna não nula; no exemplo usamos STATUS)
    4. Totais:
       - Total por linha: soma horizontal (D1 + >D+1)
       - Total Geral (linha extra): soma vertical de cada coluna e do total
    5. Ordenação das linhas: Cash, Contas a Pagar, Contas a Receber, Tesouraria, seguidas da linha Total Geral
    6. Data "Dia útil": utilizar o máximo de DATA_EXTRATO presente na aba Pendências
    """

    @staticmethod
    def gerar_resumo(pendencias: List[Pendencia]) -> ResumoConsolidado:
        """
        Gera o resumo consolidado das pendências.
        
        Conforme especificação do README_Resumo_Pivot.md:
        - Apenas STATUS == "Não Reconciliada"
        - Apenas colunas D1 e >D+1
        - Ordem específica dos departamentos
        - Dia útil = max(DATA_EXTRATO)
        
        Args:
            pendencias: Lista de pendências consolidadas
            
        Returns:
            ResumoConsolidado: Resumo estruturado com estatísticas
        """
        # 1. Filtrar apenas STATUS == "Não Reconciliada"
        pendencias_nao_reconciliadas = [
            p for p in pendencias 
            if p.STATUS and p.STATUS.strip() == "Não Reconciliada"
        ]
        
        # 2. Obter dia útil de referência (max de DATA_EXTRATO)
        dia_util_referencia = ResumoService._obter_dia_util_de_data_extrato(pendencias_nao_reconciliadas)
        
        # 3. Agrupar pendências por departamento
        departamentos_stats = ResumoService._agrupar_por_departamento(pendencias_nao_reconciliadas)
        
        # 4. Ordem específica dos departamentos
        ordem_departamentos = ["Cash", "Contas a Pagar", "Contas a Receber", "Tesouraria"]
        
        # 5. Criar itens do resumo na ordem especificada
        itens_resumo = []
        total_d1 = 0
        total_d_mais_1 = 0
        
        for departamento in ordem_departamentos:
            stats = departamentos_stats.get(departamento, {'D1': 0, '>D+1': 0})
            d1 = stats['D1']
            d_mais_1 = stats['>D+1']
            total_geral = d1 + d_mais_1
            
            item = ResumoItem(
                departamento=departamento,
                d1=d1,
                d_mais_1=d_mais_1,
                total_geral=total_geral
            )
            
            itens_resumo.append(item)
            
            # Acumular totais verticais
            total_d1 += d1
            total_d_mais_1 += d_mais_1
        
        total_geral_absoluto = total_d1 + total_d_mais_1
        
        return ResumoConsolidado(
            itens=itens_resumo,
            total_d1=total_d1,
            total_d_mais_1=total_d_mais_1,
            total_geral_absoluto=total_geral_absoluto,
            dia_util_referencia=dia_util_referencia
        )

    @staticmethod
    def _agrupar_por_departamento(pendencias: List[Pendencia]) -> Dict[str, Dict[str, int]]:
        """
        Agrupa pendências por departamento e conta por tipo de vencimento.
        
        Conforme README_Resumo_Pivot.md: apenas D1 e >D+1
        
        Args:
            pendencias: Lista de pendências
            
        Returns:
            Dict[str, Dict[str, int]]: {departamento: {'D1': count, '>D+1': count}}
        """
        departamentos_stats = {}
        
        for pendencia in pendencias:
            # Determinar departamento
            departamento = pendencia.DEPARTAMENTO or "(não definido)"
            
            # Determinar tipo de vencimento (apenas D1 ou >D+1)
            vencimento = pendencia.VENCIMENTO or ""
            tipo_vencimento = ResumoService._classificar_vencimento(vencimento)
            
            # Inicializar estrutura se necessário
            if departamento not in departamentos_stats:
                departamentos_stats[departamento] = {
                    'D1': 0,
                    '>D+1': 0
                }
            
            # Incrementar contador apenas se for D1 ou >D+1
            if tipo_vencimento in ['D1', '>D+1']:
                departamentos_stats[departamento][tipo_vencimento] += 1
        
        return departamentos_stats

    @staticmethod
    def _classificar_vencimento(vencimento: str) -> str:
        """
        Classifica o vencimento nas categorias do resumo.
        
        Conforme README_Resumo_Pivot.md: apenas D1 e >D+1
        
        Args:
            vencimento: String do vencimento
            
        Returns:
            str: Categoria ('D1' ou '>D+1', ou None se não se enquadrar)
        """
        if not vencimento or vencimento.strip() == "":
            return None
        elif vencimento.strip() == "D1":
            return 'D1'
        elif vencimento.strip() == ">D+1":
            return '>D+1'
        else:
            # Para qualquer outro valor, retornar None (não contar)
            return None

    @staticmethod
    def _obter_dia_util_de_data_extrato(pendencias: List[Pendencia]) -> date:
        """
        Obtém o dia útil de referência a partir do máximo de DATA_EXTRATO.
        
        Conforme README_Resumo_Pivot.md: utilizar o máximo de DATA_EXTRATO presente na aba Pendências
        
        Args:
            pendencias: Lista de pendências
            
        Returns:
            date: Data de referência (max de DATA_EXTRATO)
        """
        datas_validas = []
        
        for pendencia in pendencias:
            if pendencia.DATA_EXTRATO:
                # Converter para date se for datetime
                if isinstance(pendencia.DATA_EXTRATO, datetime):
                    datas_validas.append(pendencia.DATA_EXTRATO.date())
                elif isinstance(pendencia.DATA_EXTRATO, date):
                    datas_validas.append(pendencia.DATA_EXTRATO)
        
        # Retornar o máximo, ou data atual se não houver datas válidas
        if datas_validas:
            return max(datas_validas)
        else:
            return date.today()

    @staticmethod
    def gerar_relatorio_texto(resumo: ResumoConsolidado) -> str:
        """
        Gera relatório em formato texto para visualização.
        
        Conforme README_Resumo_Pivot.md
        
        Args:
            resumo: Resumo consolidado
            
        Returns:
            str: Relatório formatado em texto
        """
        linhas = []
        
        # Cabeçalho
        linhas.append("=" * 70)
        linhas.append("RESUMO DE PENDÊNCIAS POR DEPARTAMENTO")
        linhas.append("=" * 70)
        linhas.append(f"Dia útil: {resumo.dia_util_referencia.strftime('%d/%m/%Y')}")
        linhas.append("")
        
        # Header da tabela
        header = f"{'Departamento':<30} {'D1':>10} {'>D+1':>10} {'Total Geral':>15}"
        linhas.append(header)
        linhas.append("-" * len(header))
        
        # Dados por departamento
        for item in resumo.itens:
            linha = f"{item.departamento:<30} {item.d1:>10} {item.d_mais_1:>10} {item.total_geral:>15}"
            linhas.append(linha)
        
        # Linha separadora
        linhas.append("-" * len(header))
        
        # Total geral (linha vertical)
        total_linha = f"{'Total Geral':<30} {resumo.total_d1:>10} {resumo.total_d_mais_1:>10} {resumo.total_geral_absoluto:>15}"
        linhas.append(total_linha)
        
        return "\n".join(linhas)

    @staticmethod
    def gerar_dados_excel(resumo: ResumoConsolidado) -> List[Dict]:
        """
        Gera dados estruturados para exportação Excel.
        
        Conforme README_Resumo_Pivot.md
        
        Args:
            resumo: Resumo consolidado
            
        Returns:
            List[Dict]: Lista de dicionários para exportação
        """
        dados = []
        
        # Dados por departamento
        for item in resumo.itens:
            dados.append({
                'Departamento': item.departamento,
                'D1': item.d1,
                '>D+1': item.d_mais_1,
                'Total Geral': item.total_geral
            })
        
        # Total geral
        dados.append({
            'Departamento': 'Total Geral',
            'D1': resumo.total_d1,
            '>D+1': resumo.total_d_mais_1,
            'Total Geral': resumo.total_geral_absoluto
        })
        
        return dados 