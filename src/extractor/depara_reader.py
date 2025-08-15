import pandas as pd
from typing import List
from entities.responsavel import Responsavel
from entities.departamento import Departamento


def extrair_responsaveis(caminho: str, aba: str = 'responsaveis') -> List[Responsavel]:
    """
    Extrai dados da sheet 'responsaveis' do arquivo DePara-CashFlow.
    
    Args:
        caminho: Caminho para o arquivo Excel DePara-CashFlow
        aba: Nome da aba/sheet a ser lida (padrão: 'responsaveis')
        
    Returns:
        List[Responsavel]: Lista de objetos Responsavel
        
    Raises:
        FileNotFoundError: Se o arquivo não for encontrado
        ValueError: Se a aba não existir ou houver erro de formato
    """
    try:
        df = pd.read_excel(caminho, sheet_name=aba, engine='openpyxl')
        return _dataframe_para_responsaveis(df)
    except FileNotFoundError:
        raise FileNotFoundError(f"Arquivo DePara não encontrado: {caminho}")
    except ValueError as e:
        if "Worksheet" in str(e):
            raise ValueError(f"Aba '{aba}' não encontrada no arquivo DePara")
        raise ValueError(f"Erro ao processar dados da aba '{aba}': {str(e)}")


def extrair_departamentos(caminho: str, aba: str = 'departamentos') -> List[Departamento]:
    """
    Extrai dados da sheet 'departamentos' do arquivo DePara-CashFlow.
    
    Args:
        caminho: Caminho para o arquivo Excel DePara-CashFlow
        aba: Nome da aba/sheet a ser lida (padrão: 'departamentos')
        
    Returns:
        List[Departamento]: Lista de objetos Departamento
        
    Raises:
        FileNotFoundError: Se o arquivo não for encontrado
        ValueError: Se a aba não existir ou houver erro de formato
    """
    try:
        df = pd.read_excel(caminho, sheet_name=aba, engine='openpyxl')
        return _dataframe_para_departamentos(df)
    except FileNotFoundError:
        raise FileNotFoundError(f"Arquivo DePara não encontrado: {caminho}")
    except ValueError as e:
        if "Worksheet" in str(e):
            raise ValueError(f"Aba '{aba}' não encontrada no arquivo DePara")
        raise ValueError(f"Erro ao processar dados da aba '{aba}': {str(e)}")


def _dataframe_para_responsaveis(df: pd.DataFrame) -> List[Responsavel]:
    """
    Converte um DataFrame pandas para lista de objetos Responsavel.
    
    Args:
        df: DataFrame com os dados da sheet responsaveis
        
    Returns:
        List[Responsavel]: Lista de objetos Responsavel
    """
    responsaveis = []
    
    for _, row in df.iterrows():
        responsavel = Responsavel(
            NOME_BANCO=_get_valor_coluna(row, 'NOME_BANCO'),
            INFORMACAO_ADICIONAL=_get_valor_coluna(row, 'INFORMACAO_ADICIONAL'),
            TIPO_TRANSACAO=_get_valor_coluna(row, 'TIPO_TRANSACAO'),
            RESPONSAVEL=_get_valor_coluna(row, ['RESPONSAVEL ', 'RESPONSAVEL']),  # Com e sem espaço
            OBSERVACAO=_get_valor_coluna(row, ['OBSERVAÇÃO', 'OBSERVACAO'])
        )
        
        responsaveis.append(responsavel)
    
    return responsaveis


def _dataframe_para_departamentos(df: pd.DataFrame) -> List[Departamento]:
    """
    Converte um DataFrame pandas para lista de objetos Departamento.
    
    Args:
        df: DataFrame com os dados da sheet departamento
        
    Returns:
        List[Departamento]: Lista de objetos Departamento
    """
    departamentos = []
    
    for _, row in df.iterrows():
        departamento = Departamento(
            RESPONSAVEL=_get_valor_coluna(row, ['Responsável', 'RESPONSAVEL']),
            AREA=_get_valor_coluna(row, ['Área', 'AREA'])
        )
        
        departamentos.append(departamento)
    
    return departamentos


def _get_valor_coluna(row, nomes_coluna):
    """
    Obtém o valor de uma coluna do DataFrame, tentando diferentes variações do nome.
    
    Args:
        row: Linha do DataFrame
        nomes_coluna: Nome da coluna ou lista de possíveis nomes
        
    Returns:
        Valor da coluna ou None se não encontrada
    """
    if isinstance(nomes_coluna, str):
        nomes_coluna = [nomes_coluna]
    
    for nome in nomes_coluna:
        if nome in row.index:
            valor = row[nome]
            # Retorna None para valores NaN ou vazios
            if pd.isna(valor) or (isinstance(valor, str) and valor.strip() == ''):
                return None
            return valor
    
    return None 