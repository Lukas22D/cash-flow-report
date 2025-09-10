import pandas as pd
from typing import List
from entities.pendencia import Pendencia


def extrair_novas_transacoes_rel_sem_tratar(caminho: str) -> List[Pendencia]:
    """
    Extrai novas transações do arquivo Rel_sem_tratar.xlsx.
    
    Características específicas:
    - Header começa na linha 2 (índice 1)
    - Numeração americana (vírgula como separador de milhares, ponto para decimal)
    
    Args:
        caminho: Caminho para o arquivo Rel_sem_tratar.xlsx
        
    Returns:
        List[Pendencia]: Lista de objetos Pendencia (novas transações)
        
    Raises:
        FileNotFoundError: Se o arquivo não for encontrado
        ValueError: Se houver erro de formato
    """
    try:
        # Ler com header na linha 2 (índice 1)
        df = pd.read_excel(caminho, header=1, engine='openpyxl')
        
        # Limpar nomes das colunas (remover espaços extras, quebras de linha)
        df.columns = df.columns.str.strip().str.replace('\n', ' ').str.replace('\r', '')
        
        return _dataframe_para_pendencias_rel_sem_tratar(df)
        
    except FileNotFoundError:
        raise FileNotFoundError(f"Arquivo Rel_sem_tratar não encontrado: {caminho}")
    except Exception as e:
        raise ValueError(f"Erro ao processar arquivo Rel_sem_tratar: {str(e)}")


def _dataframe_para_pendencias_rel_sem_tratar(df: pd.DataFrame) -> List[Pendencia]:
    """
    Converte um DataFrame do Rel_sem_tratar para lista de objetos Pendencia.
    
    Args:
        df: DataFrame com os dados do Rel_sem_tratar
        
    Returns:
        List[Pendencia]: Lista de objetos Pendencia
    """
    pendencias = []
    
    for _, row in df.iterrows():
        # Tratar valor numérico (já vem correto do pandas)
        valor = row.get('VALOR', 0)
        if pd.isna(valor):
            valor = 0
            
        # Tratar data do extrato
        data_extrato = row.get('DATA_EXTRATO')
        if pd.isna(data_extrato):
            data_extrato = None
        
        # Criar objeto Pendencia
        pendencia = Pendencia(
            STATUS=_get_valor_seguro(row, 'STATUS'),
            UNIDADE_NEGOCIO=_get_valor_seguro(row, 'UNIDADE_NEGOCIO'),
            EMPRESA=_get_valor_seguro(row, 'EMPRESA'),
            NOME_BANCO=_get_valor_seguro(row, 'NOME_BANCO'),
            NOME_CONTA=_get_valor_seguro(row, 'NOME_CONTA'),
            DATA_EXTRATO=data_extrato,
            NUMERO_CONTA=_get_valor_seguro(row, 'NUMERO_CONTA'),
            INFORMACAO_ADICIONAL=_get_valor_seguro(row, 'INFORMACAO_ADICIONAL'),
            NUMERO_EXTRATO=_get_valor_seguro(row, 'NUMERO_EXTRATO'),
            TIPO_TRANSACAO=_get_valor_seguro(row, 'TIPO_TRANSACAO'),
            VALOR=valor,
            # Campos que serão preenchidos pelas regras de negócio
            RESPONSAVEL=None,
            OBSERVACAO=None,
            DEPARTAMENTO=None,
            VENCIMENTO=None
        )
        
        pendencias.append(pendencia)
    
    return pendencias


def _get_valor_seguro(row, nome_coluna):
    """
    Obtém valor de uma coluna de forma segura, tratando valores nulos.
    
    Args:
        row: Linha do DataFrame
        nome_coluna: Nome da coluna
        
    Returns:
        Valor da coluna ou None se não existir/for nulo
    """
    valor = row.get(nome_coluna)
    if pd.isna(valor):
        return None
    return valor 