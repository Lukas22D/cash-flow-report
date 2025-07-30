from dataclasses import dataclass
from datetime import date
from decimal import Decimal
from typing import Optional


@dataclass
class CashModel:
    """
    Entity representing a cash flow transaction record.
    All fields use CamelCase naming convention as requested.
    """
    
    # Basic transaction information
    status: str
    unidadeNegocio: str
    nomeBanco: str
    nomeConta: str
    dataExtrato: date
    informacaoAdicional: Optional[str] = None
    
    # Transaction details
    tipoTransacao: str
    valor: Decimal
    
    # Responsibility and tracking
    responsavel: Optional[str] = None
    observacao: Optional[str] = None
    departamento: str
    
    # Payment information
    vencimento: Optional[date] = None
    
    def __post_init__(self):
        """Validate data after initialization"""
        if not self.status:
            raise ValueError("Status is required")
        if not self.unidadeNegocio:
            raise ValueError("Unidade de Negócio is required")
        if not self.nomeBanco:
            raise ValueError("Nome do Banco is required")
        if not self.nomeConta:
            raise ValueError("Nome da Conta is required")
        if not self.tipoTransacao:
            raise ValueError("Tipo de Transação is required")
        if not self.responsavel:
            raise ValueError("Responsável is required")
        if not self.departamento:
            raise ValueError("Departamento is required")
