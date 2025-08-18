from typing import List, Dict, Tuple
from datetime import datetime, date, timedelta
from dataclasses import dataclass
from entities.pendencia import Pendencia


@dataclass
class ResumoItem:
    """
    Representa um item do resumo por departamento.
    """
    departamento: str
    d1: int  # Pendências com vencimento D1
    d_mais_1: int  # Pendências com vencimento >D+1
    vazio: int  # Pendências sem vencimento (None ou vazio)
    total_geral: int  # Total horizontal (D1 + D+1 + vazio)


@dataclass
class ResumoConsolidado:
    """
    Representa o resumo consolidado completo.
    """
    itens: List[ResumoItem]
    total_d1: int  # Total vertical D1 de todos os departamentos
    total_d_mais_1: int  # Total vertical >D+1 de todos os departamentos
    total_vazio: int  # Total vertical vazios de todos os departamentos
    total_geral_absoluto: int  # Total geral absoluto
    dia_util_referencia: date  # Dia útil usado como referência


class ResumoService:
    """
    Serviço responsável por gerar resumo estruturado das pendências após conciliação.
    
    Gera relatório com header: Departamento | D1 | >D+1 | (vazio) | Total Geral
    
    Regras:
    - Departamento: todos os departamentos presentes nas pendências
    - D1/D+1/(vazio): contagem baseada no campo vencimento das pendências
    - Total Geral horizontal: soma D1 + D+1 + vazio por departamento
    - Total Geral vertical: soma por tipo de vencimento de todos os departamentos
    - Dia útil: dia atual - 1 dia útil (considerando fins de semana)
    """

    @staticmethod
    def gerar_resumo(pendencias: List[Pendencia]) -> ResumoConsolidado:
        """
        Gera o resumo consolidado das pendências.
        
        Args:
            pendencias: Lista de pendências consolidadas
            
        Returns:
            ResumoConsolidado: Resumo estruturado com estatísticas
        """
        # Obter dia útil de referência
        dia_util_referencia = ResumoService._obter_dia_util_anterior()
        
        # Agrupar pendências por departamento
        departamentos_stats = ResumoService._agrupar_por_departamento(pendencias)
        
        # Criar itens do resumo
        itens_resumo = []
        total_d1 = 0
        total_d_mais_1 = 0
        total_vazio = 0
        
        for departamento, stats in sorted(departamentos_stats.items()):
            d1 = stats['D1']
            d_mais_1 = stats['>D+1']
            vazio = stats['vazio']
            total_geral = d1 + d_mais_1 + vazio
            
            item = ResumoItem(
                departamento=departamento,
                d1=d1,
                d_mais_1=d_mais_1,
                vazio=vazio,
                total_geral=total_geral
            )
            
            itens_resumo.append(item)
            
            # Acumular totais verticais
            total_d1 += d1
            total_d_mais_1 += d_mais_1
            total_vazio += vazio
        
        total_geral_absoluto = total_d1 + total_d_mais_1 + total_vazio
        
        return ResumoConsolidado(
            itens=itens_resumo,
            total_d1=total_d1,
            total_d_mais_1=total_d_mais_1,
            total_vazio=total_vazio,
            total_geral_absoluto=total_geral_absoluto,
            dia_util_referencia=dia_util_referencia
        )

    @staticmethod
    def _agrupar_por_departamento(pendencias: List[Pendencia]) -> Dict[str, Dict[str, int]]:
        """
        Agrupa pendências por departamento e conta por tipo de vencimento.
        
        Args:
            pendencias: Lista de pendências
            
        Returns:
            Dict[str, Dict[str, int]]: {departamento: {vencimento: quantidade}}
        """
        departamentos_stats = {}
        
        for pendencia in pendencias:
            # Determinar departamento (usar string vazia se não houver)
            departamento = pendencia.DEPARTAMENTO or "(não definido)"
            
            # Determinar tipo de vencimento
            vencimento = pendencia.VENCIMENTO or ""
            tipo_vencimento = ResumoService._classificar_vencimento(vencimento)
            
            # Inicializar estrutura se necessário
            if departamento not in departamentos_stats:
                departamentos_stats[departamento] = {
                    'D1': 0,
                    '>D+1': 0,
                    'vazio': 0
                }
            
            # Incrementar contador
            departamentos_stats[departamento][tipo_vencimento] += 1
        
        return departamentos_stats

    @staticmethod
    def _classificar_vencimento(vencimento: str) -> str:
        """
        Classifica o vencimento nas categorias do resumo.
        
        Args:
            vencimento: String do vencimento
            
        Returns:
            str: Categoria ('D1', '>D+1' ou 'vazio')
        """
        if not vencimento or vencimento.strip() == "":
            return 'vazio'
        elif vencimento.strip() == "D1":
            return 'D1'
        elif vencimento.strip() == ">D+1":
            return '>D+1'
        else:
            # Para qualquer outro valor, classificar como vazio
            return 'vazio'

    @staticmethod
    def _obter_dia_util_anterior() -> date:
        """
        Obtém o último dia útil anterior à data atual.
        
        Lógica:
        - Se hoje é Segunda-feira: retorna Sexta anterior (3 dias atrás)
        - Se hoje é Terça a Sexta: retorna dia anterior (1 dia atrás)
        - Se hoje é Sábado: retorna Sexta anterior (1 dia atrás)
        - Se hoje é Domingo: retorna Sexta anterior (2 dias atrás)
        
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
    def gerar_relatorio_texto(resumo: ResumoConsolidado) -> str:
        """
        Gera relatório em formato texto para visualização.
        
        Args:
            resumo: Resumo consolidado
            
        Returns:
            str: Relatório formatado em texto
        """
        linhas = []
        
        # Cabeçalho
        linhas.append("=" * 80)
        linhas.append("RESUMO DE PENDÊNCIAS POR DEPARTAMENTO")
        linhas.append("=" * 80)
        linhas.append(f"Dia útil de referência: {resumo.dia_util_referencia.strftime('%d/%m/%Y')}")
        linhas.append("")
        
        # Header da tabela
        header = f"{'Departamento':<30} {'D1':>8} {'>D+1':>8} {'(vazio)':>8} {'Total Geral':>12}"
        linhas.append(header)
        linhas.append("-" * len(header))
        
        # Dados por departamento
        for item in resumo.itens:
            linha = f"{item.departamento:<30} {item.d1:>8} {item.d_mais_1:>8} {item.vazio:>8} {item.total_geral:>12}"
            linhas.append(linha)
        
        # Linha separadora
        linhas.append("-" * len(header))
        
        # Total geral (linha vertical)
        total_linha = f"{'Total Geral':<30} {resumo.total_d1:>8} {resumo.total_d_mais_1:>8} {resumo.total_vazio:>8} {resumo.total_geral_absoluto:>12}"
        linhas.append(total_linha)
        
        return "\n".join(linhas)

    @staticmethod
    def gerar_dados_excel(resumo: ResumoConsolidado) -> List[Dict]:
        """
        Gera dados estruturados para exportação Excel.
        
        Args:
            resumo: Resumo consolidado
            
        Returns:
            List[Dict]: Lista de dicionários para exportação
        """
        dados = []
        
        # Adicionar linha de cabeçalho informativo
        dados.append({
            'Departamento': f'Dia útil: {resumo.dia_util_referencia.strftime("%d/%m/%Y")}',
            'D1': '',
            '>D+1': '',
            '(vazio)': '',
            'Total Geral': ''
        })
        
        # Linha vazia para separação
        dados.append({
            'Departamento': '',
            'D1': '',
            '>D+1': '',
            '(vazio)': '',
            'Total Geral': ''
        })
        
        # Dados por departamento
        for item in resumo.itens:
            dados.append({
                'Departamento': item.departamento,
                'D1': item.d1,
                '>D+1': item.d_mais_1,
                '(vazio)': item.vazio,
                'Total Geral': item.total_geral
            })
        
        # Total geral
        dados.append({
            'Departamento': 'Total Geral',
            'D1': resumo.total_d1,
            '>D+1': resumo.total_d_mais_1,
            '(vazio)': resumo.total_vazio,
            'Total Geral': resumo.total_geral_absoluto
        })
        
        return dados 