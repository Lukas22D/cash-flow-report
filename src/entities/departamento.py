from dataclasses import dataclass
from typing import Optional


@dataclass
class Departamento:
    """
    Entidade que representa um registro da sheet 'departamento' do arquivo DePara-CashFlow.
    
    Headers da sheet:
    - RESPONSAVEL: Nome do responsável
    - AREA: Área/departamento do responsável
    """
    RESPONSAVEL: Optional[str] = None
    AREA: Optional[str] = None

    def to_dict(self) -> dict:
        """
        Converte o departamento para dicionário para exportação.
        
        Returns:
            dict: Dicionário com os dados do departamento
        """
        return {
            'RESPONSAVEL': self.RESPONSAVEL,
            'AREA': self.AREA
        }

    def get_chave_identificacao(self) -> str:
        """
        Gera uma chave de identificação baseada no responsável.
        
        Returns:
            str: Chave única para identificar o registro de departamento
        """
        responsavel_str = str(self.RESPONSAVEL) if self.RESPONSAVEL is not None else ""
        return responsavel_str 