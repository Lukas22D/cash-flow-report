from dataclasses import dataclass
from typing import Optional
from datetime import datetime


@dataclass
class Pendencia:
    """
    Entidade que representa uma pendência financeira.
    
    Atributos obrigatórios para chave de reconciliação:
    - VALOR: valor da transação
    - INFORMACAO_ADICIONAL: informação adicional da transação
    - NOME_CONTA: nome da conta
    """
    STATUS: Optional[str] = None
    UNIDADE_NEGOCIO: Optional[str] = None
    EMPRESA: Optional[str] = None
    NOME_BANCO: Optional[str] = None
    NOME_CONTA: Optional[str] = None
    DATA_EXTRATO: Optional[datetime] = None
    NUMERO_CONTA: Optional[str] = None
    INFORMACAO_ADICIONAL: Optional[str] = None
    NUMERO_EXTRATO: Optional[str] = None
    TIPO_TRANSACAO: Optional[str] = None
    VALOR: Optional[float] = None
    RESPONSAVEL: Optional[str] = None
    OBSERVACAO: Optional[str] = None
    DEPARTAMENTO: Optional[str] = None
    VENCIMENTO: Optional[datetime] = None

    def get_chave_reconciliacao(self) -> str:
        """
        Gera a chave de reconciliação baseada em VALOR + INFORMACAO_ADICIONAL + NOME_CONTA.
        
        Returns:
            str: Chave única para identificar a pendência
        """
        valor_str = str(self.VALOR) if self.VALOR is not None else ""
        info_str = str(self.INFORMACAO_ADICIONAL) if self.INFORMACAO_ADICIONAL is not None else ""
        conta_str = str(self.NOME_CONTA) if self.NOME_CONTA is not None else ""
        
        return f"{valor_str}{info_str}{conta_str}"

    def to_dict(self) -> dict:
        """
        Converte a pendência para dicionário para exportação para Excel.
        
        Returns:
            dict: Dicionário com os dados da pendência
        """
        return {
            'STATUS': self.STATUS,
            'UNIDADE_NEGOCIO': self.UNIDADE_NEGOCIO,
            'EMPRESA': self.EMPRESA,
            'NOME_BANCO': self.NOME_BANCO,
            'NOME_CONTA': self.NOME_CONTA,
            'DATA_EXTRATO': self.DATA_EXTRATO,
            'NUMERO_CONTA': self.NUMERO_CONTA,
            'INFORMACAO_ADICIONAL': self.INFORMACAO_ADICIONAL,
            'NUMERO_EXTRATO': self.NUMERO_EXTRATO,
            'TIPO_TRANSACAO': self.TIPO_TRANSACAO,
            'VALOR': self.VALOR,
            'Responsável': self.RESPONSAVEL,  # Mantém nome original da coluna
            'Observação': self.OBSERVACAO,   # Mantém nome original da coluna
            'Departamento': self.DEPARTAMENTO,
            'Vencimento': self.VENCIMENTO
        } 