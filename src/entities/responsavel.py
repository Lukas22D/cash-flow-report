from dataclasses import dataclass
from typing import Optional


@dataclass
class Responsavel:
    """
    Entidade que representa um registro da sheet 'responsaveis' do arquivo DePara-CashFlow.
    
    Headers da sheet:
    - NOME_BANCO: Nome do banco
    - INFORMACAO_ADICIONAL: Informação adicional da transação
    - TIPO_TRANSACAO: Tipo da transação
    - RESPONSAVEL: Nome do responsável
    - OBSERVAÇÃO: Observações adicionais
    """
    NOME_BANCO: Optional[str] = None
    INFORMACAO_ADICIONAL: Optional[str] = None
    TIPO_TRANSACAO: Optional[str] = None
    RESPONSAVEL: Optional[str] = None
    OBSERVACAO: Optional[str] = None

    def to_dict(self) -> dict:
        """
        Converte o responsável para dicionário para exportação.
        
        Returns:
            dict: Dicionário com os dados do responsável
        """
        return {
            'NOME_BANCO': self.NOME_BANCO,
            'INFORMACAO_ADICIONAL': self.INFORMACAO_ADICIONAL,
            'TIPO_TRANSACAO': self.TIPO_TRANSACAO,
            'RESPONSAVEL': self.RESPONSAVEL,
            'OBSERVAÇÃO': self.OBSERVACAO
        }

    def get_chave_identificacao(self) -> str:
        """
        Gera uma chave de identificação baseada nos campos principais.
        
        Returns:
            str: Chave única para identificar o registro de responsável
        """
        banco_str = str(self.NOME_BANCO) if self.NOME_BANCO is not None else ""
        info_str = str(self.INFORMACAO_ADICIONAL) if self.INFORMACAO_ADICIONAL is not None else ""
        tipo_str = str(self.TIPO_TRANSACAO) if self.TIPO_TRANSACAO is not None else ""
        
        return f"{banco_str}{info_str}{tipo_str}" 