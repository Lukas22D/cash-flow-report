import pandas as pd
from typing import List
from entities.pendencia import Pendencia
from services.resumo_service import ResumoService
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo


class ExcelWriter:
    """
    Responsável por escrever dados consolidados em arquivo Excel.
    """
    
    @staticmethod
    def salvar_relatorio_consolidado(pendencias_consolidadas: List[Pendencia],
                                   df_resumo: pd.DataFrame,
                                   caminho_saida: str) -> None:
        """
        Salva o relatório consolidado em arquivo Excel com formatação profissional.
        
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
            
            # Criar workbook com formatação customizada
            wb = Workbook()
            
            # Remover sheet padrão
            if 'Sheet' in wb.sheetnames:
                wb.remove(wb['Sheet'])
            
            # Criar e formatar aba de Pendências
            ExcelWriter._criar_aba_pendencias(wb, df_pendencias)
            
            # Criar aba de Resumo original (se existir)
            if not df_resumo.empty:
                ExcelWriter._criar_aba_resumo_original(wb, df_resumo)
            
            # Criar e formatar aba de Resumo de Pendências
            ExcelWriter._criar_aba_resumo_pendencias(wb, df_resumo_pendencias)
            
            # Salvar arquivo
            wb.save(caminho_saida)
                    
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
    def _criar_aba_pendencias(wb: Workbook, df_pendencias: pd.DataFrame) -> None:
        """
        Cria e formata a aba de Pendências como uma tabela profissional.
        """
        ws = wb.create_sheet("Pendências", 0)  # Primeira aba
        
        # Adicionar dados do DataFrame
        for r in dataframe_to_rows(df_pendencias, index=False, header=True):
            ws.append(r)
        
        # Aplicar formatação
        ExcelWriter._formatar_aba_como_tabela(ws, df_pendencias, "Pendências")
        
        # Ajustar larguras específicas para pendências
        larguras_pendencias = {
            'A': 20,  # STATUS
            'B': 25,  # UNIDADE_NEGOCIO
            'C': 30,  # EMPRESA
            'D': 20,  # NOME_BANCO
            'E': 20,  # NOME_CONTA
            'F': 15,  # DATA_EXTRATO
            'G': 15,  # NUMERO_CONTA
            'H': 30,  # INFORMACAO_ADICIONAL
            'I': 20,  # NUMERO_EXTRATO
            'J': 15,  # TIPO_TRANSACAO
            'K': 15,  # VALOR
            'L': 20,  # Responsável
            'M': 25,  # Observação
            'N': 20,  # Departamento
            'O': 12   # Vencimento
        }
        
        for col, largura in larguras_pendencias.items():
            ws.column_dimensions[col].width = largura
    
    @staticmethod
    def _criar_aba_resumo_original(wb: Workbook, df_resumo: pd.DataFrame) -> None:
        """
        Cria e formata a aba de Resumo original.
        """
        ws = wb.create_sheet("Resumo Original")
        
        # Adicionar dados do DataFrame
        for r in dataframe_to_rows(df_resumo, index=False, header=True):
            ws.append(r)
        
        # Aplicar formatação
        ExcelWriter._formatar_aba_como_tabela(ws, df_resumo, "ResumoOriginal")
        
        # Ajustar larguras automaticamente
        ExcelWriter._ajustar_larguras_automaticas(ws)
    
    @staticmethod
    def _criar_aba_resumo_pendencias(wb: Workbook, df_resumo_pendencias: pd.DataFrame) -> None:
        """
        Cria e formata a aba de Resumo de Pendências.
        """
        ws = wb.create_sheet("Resumo Pendências")
        
        # Adicionar dados do DataFrame
        for r in dataframe_to_rows(df_resumo_pendencias, index=False, header=True):
            ws.append(r)
        
        # Aplicar formatação
        ExcelWriter._formatar_aba_como_tabela(ws, df_resumo_pendencias, "ResumoPendencias")
        
        # Ajustar larguras específicas para resumo
        larguras_resumo = {
            'A': 25,  # Departamento
            'B': 15,  # D1
            'C': 15,  # >D+1
            'D': 15,  # Vazio
            'E': 15   # Total
        }
        
        for col, largura in larguras_resumo.items():
            if col in [c.column_letter for c in ws[1]]:  # Verificar se coluna existe
                ws.column_dimensions[col].width = largura
    
    @staticmethod
    def _formatar_aba_como_tabela(ws, df: pd.DataFrame, nome_tabela: str) -> None:
        """
        Formata uma aba como tabela profissional do Excel.
        """
        if df.empty:
            return
        
        # Definir range da tabela
        max_row = len(df) + 1  # +1 para o header
        max_col = len(df.columns)
        table_range = f"A1:{ws.cell(row=max_row, column=max_col).coordinate}"
        
        # Criar tabela
        table = Table(displayName=nome_tabela, ref=table_range)
        
        # Estilo da tabela
        style = TableStyleInfo(
            name="TableStyleMedium9",  # Azul profissional
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False
        )
        table.tableStyleInfo = style
        
        # Adicionar tabela à worksheet
        ws.add_table(table)
        
        # Formatação do cabeçalho
        header_font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
        header_fill = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
        header_alignment = Alignment(horizontal='center', vertical='center')
        
        # Aplicar formatação ao cabeçalho
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment
        
        # Formatação das células de dados
        data_font = Font(name='Calibri', size=10)
        data_alignment = Alignment(horizontal='left', vertical='center')
        
        # Aplicar formatação aos dados
        for row in ws.iter_rows(min_row=2, max_row=max_row, max_col=max_col):
            for cell in row:
                cell.font = data_font
                cell.alignment = data_alignment
                
                # Formatação específica para valores numéricos
                if isinstance(cell.value, (int, float)) and cell.column_letter in ['K']:  # Coluna VALOR
                    cell.number_format = '#,##0.00'
                    cell.alignment = Alignment(horizontal='right', vertical='center')
        
        # Congelar painéis (primeira linha)
        ws.freeze_panes = 'A2'
    
    @staticmethod
    def _ajustar_larguras_automaticas(ws) -> None:
        """
        Ajusta automaticamente as larguras das colunas baseado no conteúdo.
        """
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            
            # Definir largura com limite máximo
            adjusted_width = min(max_length + 2, 50)
            ws.column_dimensions[column_letter].width = adjusted_width
    
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