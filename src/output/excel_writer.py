import pandas as pd
from typing import List
from entities.pendencia import Pendencia
from services.resumo_service import ResumoService


class ExcelWriter:
    """
    Responsável por escrever dados consolidados em arquivo Excel.
    """
    
    @staticmethod
    def salvar_relatorio_consolidado(pendencias_consolidadas: List[Pendencia],
                                   df_resumo: pd.DataFrame,
                                   caminho_saida: str) -> None:
        """
        Salva o relatório consolidado em arquivo Excel.
        
        Args:
            pendencias_consolidadas: Lista de pendências consolidadas
            df_resumo: DataFrame com dados do resumo (pode estar vazio)
            caminho_saida: Caminho onde salvar o arquivo
            
        Raises:
            PermissionError: Se não conseguir escrever no arquivo
            Exception: Outros erros durante a escrita
        """
        try:
            # Converter pendências para DataFrame
            df_pendencias = ExcelWriter._pendencias_para_dataframe(pendencias_consolidadas)
            
            # Gerar resumo das pendências
            resumo_consolidado = ResumoService.gerar_resumo(pendencias_consolidadas)
            df_resumo_pendencias = ExcelWriter._resumo_para_dataframe(resumo_consolidado)
            
            # Escrever no arquivo Excel
            with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
                # Aba de Pendências
                df_pendencias.to_excel(writer, sheet_name='Pendências', index=False)
                
                # Aba de Resumo (se existir dados do arquivo original)
                if not df_resumo.empty:
                    df_resumo.to_excel(writer, sheet_name='Resumo', index=False)
                
                # Aba de Resumo de Pendências (sempre gerada)
                df_resumo_pendencias.to_excel(writer, sheet_name='Resumo Pendências', index=False)
                    
        except PermissionError:
            raise PermissionError(f"Não foi possível salvar o arquivo. "
                                f"Verifique se o arquivo não está aberto em outro programa: {caminho_saida}")
        except Exception as e:
            raise Exception(f"Erro ao salvar arquivo Excel: {str(e)}")
    
    @staticmethod
    def _pendencias_para_dataframe(pendencias: List[Pendencia]) -> pd.DataFrame:
        """
        Converte lista de pendências para DataFrame pandas.
        
        Args:
            pendencias: Lista de objetos Pendencia
            
        Returns:
            pd.DataFrame: DataFrame com os dados das pendências
        """
        if not pendencias:
            return pd.DataFrame()
        
        # Converter cada pendência para dicionário
        dados = [pendencia.to_dict() for pendencia in pendencias]
        
        # Criar DataFrame
        df = pd.DataFrame(dados)
        
        # Garantir ordem das colunas (mesma do arquivo original)
        colunas_ordenadas = [
            'STATUS',
            'UNIDADE_NEGOCIO', 
            'EMPRESA',
            'NOME_BANCO',
            'NOME_CONTA',
            'DATA_EXTRATO',
            'NUMERO_CONTA',
            'INFORMACAO_ADICIONAL',
            'NUMERO_EXTRATO',
            'TIPO_TRANSACAO',
            'VALOR',
            'Responsável',
            'Observação',
            'Departamento', 
            'Vencimento'
        ]
        
        # Reordenar colunas, mantendo apenas as que existem
        colunas_existentes = [col for col in colunas_ordenadas if col in df.columns]
        df = df[colunas_existentes]
        
        return df
    
    @staticmethod
    def _resumo_para_dataframe(resumo_consolidado) -> pd.DataFrame:
        """
        Converte o resumo consolidado para DataFrame pandas.
        
        Args:
            resumo_consolidado: Objeto ResumoConsolidado
            
        Returns:
            pd.DataFrame: DataFrame com os dados do resumo
        """
        # Usar o método do ResumoService para gerar dados estruturados
        dados_excel = ResumoService.gerar_dados_excel(resumo_consolidado)
        
        # Criar DataFrame
        df = pd.DataFrame(dados_excel)
        
        return df
    
    @staticmethod
    def validar_caminho_saida(caminho_saida: str) -> bool:
        """
        Valida se é possível escrever no caminho especificado.
        
        Args:
            caminho_saida: Caminho do arquivo de saída
            
        Returns:
            bool: True se é possível escrever, False caso contrário
        """
        try:
            import os
            
            # Verificar se o diretório existe
            diretorio = os.path.dirname(caminho_saida)
            if diretorio and not os.path.exists(diretorio):
                return False
            
            # Tentar criar arquivo temporário para testar permissões
            with open(caminho_saida, 'w') as f:
                pass
            os.remove(caminho_saida)
            return True
            
        except (PermissionError, OSError):
            return False 