# 📊 Gerador de Relatório Cash Flow - Consolidado

Sistema modular para consolidação de pendências financeiras com interface gráfica intuitiva.

## 🏗️ Nova Arquitetura Modular

A aplicação foi refatorada seguindo princípios de **Clean Architecture** com separação clara de responsabilidades:

### 📁 Estrutura de Diretórios

```
cash-flow-report/
├── src/
│   ├── entities/           # 🎯 Entidades de domínio
│   │   ├── __init__.py
│   │   └── pendencia.py    # Classe Pendencia com @dataclass
│   ├── extractor/          # 📥 Extração de dados do Excel
│   │   ├── __init__.py
│   │   ├── excel_reader.py # Funções de leitura
│   │   └── main.py         # Orquestrador principal
│   ├── services/           # 🔧 Lógica de negócio
│   │   ├── __init__.py
│   │   └── conciliacao_service.py # Serviço de conciliação
│   ├── output/             # 💾 Saída de dados
│   │   ├── __init__.py
│   │   └── excel_writer.py # Escrita para Excel
│   └── app.py              # 🖥️ Interface gráfica
├── docs/                   # 📄 Arquivos de exemplo
├── run_app.py              # 🚀 Script de inicialização
└── README.md              # 📖 Esta documentação
```

## 🎯 Componentes da Arquitetura

### 1. **Entidades** (`src/entities/`)
- **`Pendencia`**: Classe que representa uma pendência financeira
- Todos os atributos tipados conforme especificação
- Métodos para chave de reconciliação e conversão

### 2. **Extrator** (`src/extractor/`)
- **`extrair_pendencias()`**: Extrai pendências existentes
- **`extrair_transacoes()`**: Extrai novas transações 
- **`extrair_resumo()`**: Extrai planilha de resumo
- Conversão automática para objetos `Pendencia`

### 3. **Serviços** (`src/services/`)
- **`ConciliacaoService`**: Implementa lógica de negócio
- Consolida pendências usando chave: `VALOR + INFORMACAO_ADICIONAL + NOME_CONTA`
- Preserva dados extras (Responsável, Observação, etc.)
- Fornece estatísticas detalhadas

### 4. **Saída** (`src/output/`)
- **`ExcelWriter`**: Responsável por salvar arquivos Excel
- Mantém formatação e ordem das colunas
- Inclui validações de permissão de escrita

## 🚀 Como Usar

### Interface Gráfica

**Opção 1: Script de inicialização (recomendado)**
```bash
python3 run_app.py
```

**Opção 2: Execução direta**
```bash
cd src
python3 app.py
```

**Opção 3: Ativar ambiente virtual**
```bash
source .venv/bin/activate
python run_app.py
```

### Funcionalidades da Interface:
- ✨ **Drag & Drop**: Arraste arquivos Excel diretamente (se tkinterdnd2 estiver instalado)
- 🎯 **Interface Intuitiva**: Campos organizados em seções
- 📊 **Estatísticas Detalhadas**: Feedback completo do processamento
- 💾 **Seleção de Destino**: Escolha onde salvar apenas no momento da geração
- 🔒 **Tratamento de Erros**: Interface robusta com fallbacks

### Via Código
```python
import sys
sys.path.append('src')

from extractor.main import gerar_relatorio_consolidado

resultado = gerar_relatorio_consolidado(
    caminho_arquivo_entrada="arquivo.xlsx",
    caminho_arquivo_saida="consolidado.xlsx", 
    sheet_pendencias="Pendências",
    sheet_novas="Sheet1"
)

print(f"Consolidadas: {resultado['total_consolidadas']} linhas")
```

## 🧪 Lógica de Negócio

### Chave de Reconciliação
A identificação única de cada transação é formada por:
```
VALOR + INFORMACAO_ADICIONAL + NOME_CONTA
```

### Algoritmo de Consolidação
1. **Para cada transação em Sheet1**:
   - Se existe pendência com mesma chave → usa pendência existente (preserva dados extras)
   - Se não existe → usa nova transação como nova pendência

2. **Resultado**: Lista consolidada com exatamente o número de linhas de Sheet1

## 📈 Benefícios da Nova Arquitetura

### ✅ **Separação de Responsabilidades**
- Cada módulo tem uma função específica
- Fácil manutenção e extensão
- Testes unitários independentes

### ✅ **Reutilização**
- Componentes podem ser usados individualmente
- API limpa entre módulos
- Fácil integração com outras aplicações

### ✅ **Testabilidade**
- Cada componente é testável isoladamente
- Mocks e stubs simples de implementar
- Cobertura de testes mais abrangente

### ✅ **Extensibilidade**
- Novos formatos de entrada (CSV, JSON)
- Diferentes estratégias de conciliação
- Múltiplos formatos de saída

### ✅ **Robustez**
- Tratamento de erros em cada camada
- Fallbacks para funcionalidades opcionais
- Interface adaptável a diferentes ambientes

## 🔧 Dependências

### Obrigatórias
```bash
pip install pandas openpyxl
```

### Opcionais (para Drag & Drop)
```bash
pip install tkinterdnd2
```

> **Nota**: Se `tkinterdnd2` não estiver instalado, a interface funcionará normalmente sem a funcionalidade de arrastar e soltar.

## 📊 Resultados Esperados

- ✅ **307 linhas consolidadas** (mesmo número que Sheet1)
- ✅ **196 pendências preservadas** (com dados extras)
- ✅ **111 novas pendências adicionadas**
- ✅ **Abas**: "Pendências" + "Resumo" (se existir)

## 🛠️ Solução de Problemas

### Erro "Comando 'python' não encontrado"
```bash
# Use python3 em vez de python
python3 run_app.py
```

### Erro de "Falha de segmentação" 
```bash
# Instale tkinterdnd2 ou use a versão sem drag & drop
pip install tkinterdnd2
```

### Erro de parâmetro "-initialvalue"
✅ **Corrigido** na versão atual - agora usa `-initialfile`

## 🎉 Status

✅ **Arquitetura implementada e testada**  
✅ **Interface gráfica funcional**  
✅ **Compatibilidade com arquivos existentes**  
✅ **Resultados validados**  
✅ **Tratamento de erros robusto**  
✅ **Fallbacks para dependências opcionais**  

---

*Desenvolvido com foco em qualidade, manutenibilidade e experiência do usuário.* 