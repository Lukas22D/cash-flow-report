import pandas as pd
from typing import List
from entities.pendencia import Pendencia


def extrair_pendencias(caminho: str, aba: str) -> List[Pendencia]:
    """
    Extrai pendências existentes de uma planilha Excel.
    
    Args:
        caminho: Caminho para o arquivo Excel
        aba: Nome da aba/sheet a ser lida
        
    Returns:
        List[Pendencia]: Lista de objetos Pendencia
        
    Raises:
        FileNotFoundError: Se o arquivo não for encontrado
        ValueError: Se a aba não existir ou houver erro de formato
    """
    try:
        df = pd.read_excel(caminho, sheet_name=aba, engine='openpyxl')
        return _dataframe_para_pendencias(df)
    except FileNotFoundError:
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho}")
    except ValueError as e:
        if "Worksheet" in str(e):
            raise ValueError(f"Aba '{aba}' não encontrada no arquivo")
        raise ValueError(f"Erro ao processar dados da aba '{aba}': {str(e)}")


def extrair_transacoes(caminho: str, aba: str) -> List[Pendencia]:
    """
    Extrai novas transações de uma planilha Excel e converte para objetos Pendencia.
    
    Args:
        caminho: Caminho para o arquivo Excel
        aba: Nome da aba/sheet a ser lida
        
    Returns:
        List[Pendencia]: Lista de objetos Pendencia (novas transações)
        
    Raises:
        FileNotFoundError: Se o arquivo não for encontrado
        ValueError: Se a aba não existir ou houver erro de formato
    """
    try:
        df = pd.read_excel(caminho, sheet_name=aba, engine='openpyxl')
        return _dataframe_para_pendencias(df)
    except FileNotFoundError:
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho}")
    except ValueError as e:
        if "Worksheet" in str(e):
            raise ValueError(f"Aba '{aba}' não encontrada no arquivo")
        raise ValueError(f"Erro ao processar dados da aba '{aba}': {str(e)}")


def extrair_resumo(caminho: str) -> pd.DataFrame:
    """
    Extrai a planilha de resumo para preservação no arquivo de saída.
    
    Args:
        caminho: Caminho para o arquivo Excel
        
    Returns:
        pd.DataFrame: DataFrame com os dados do resumo
        
    Raises:
        FileNotFoundError: Se o arquivo não for encontrado
        ValueError: Se a aba 'Resumo' não existir
    """
    try:
        return pd.read_excel(caminho, sheet_name='Resumo', engine='openpyxl')
    except FileNotFoundError:
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho}")
    except ValueError:
        # Se não existe aba Resumo, retorna DataFrame vazio
        return pd.DataFrame()


def _dataframe_para_pendencias(df: pd.DataFrame) -> List[Pendencia]:
    """
    Converte um DataFrame pandas para lista de objetos Pendencia.
    
    Args:
        df: DataFrame com os dados
        
    Returns:
        List[Pendencia]: Lista de objetos Pendencia
    """
    pendencias = []
    
    for _, row in df.iterrows():
        # Mapear colunas do DataFrame para atributos da Pendencia
        # Lidar com possíveis diferenças nos nomes das colunas
        pendencia = Pendencia(
            STATUS=_get_valor_coluna(row, 'STATUS'),
            UNIDADE_NEGOCIO=_get_valor_coluna(row, 'UNIDADE_NEGOCIO'),
            EMPRESA=_get_valor_coluna(row, 'EMPRESA'),
            NOME_BANCO=_get_valor_coluna(row, 'NOME_BANCO'),
            NOME_CONTA=_get_valor_coluna(row, 'NOME_CONTA'),
            DATA_EXTRATO=_get_valor_coluna(row, 'DATA_EXTRATO'),
            NUMERO_CONTA=_get_valor_coluna(row, 'NUMERO_CONTA'),
            INFORMACAO_ADICIONAL=_get_valor_coluna(row, 'INFORMACAO_ADICIONAL'),
            NUMERO_EXTRATO=_get_valor_coluna(row, 'NUMERO_EXTRATO'),
            TIPO_TRANSACAO=_get_valor_coluna(row, 'TIPO_TRANSACAO'),
            VALOR=_get_valor_coluna(row, 'VALOR'),
            RESPONSAVEL=_get_valor_coluna(row, ['Responsável', 'RESPONSAVEL']),
            OBSERVACAO=_get_valor_coluna(row, ['Observação', 'OBSERVACAO']),
            DEPARTAMENTO=_get_valor_coluna(row, ['Departamento', 'DEPARTAMENTO']),
            VENCIMENTO=_get_valor_coluna(row, ['Vencimento', 'VENCIMENTO'])
        )
        
        pendencias.append(pendencia)
    
    return pendencias


def _get_valor_coluna(row, nomes_coluna):
    """
    Obtém o valor de uma coluna, testando diferentes nomes possíveis.
    
    Args:
        row: Linha do DataFrame
        nomes_coluna: Nome da coluna ou lista de nomes possíveis
        
    Returns:
        Valor da coluna ou None se não encontrada
    """
    if isinstance(nomes_coluna, str):
        nomes_coluna = [nomes_coluna]
    
    for nome in nomes_coluna:
        if nome in row.index:
            valor = row[nome]
            # Retornar None para valores NaN/NaT
            if pd.isna(valor):
                return None
            return valor
    
    return None 