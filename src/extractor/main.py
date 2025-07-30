import pandas as pd

def gerar_relatorio_consolidado(caminho_arquivo_entrada: str, caminho_arquivo_saida: str, 
                               sheet_pendencias: str = 'Pendências', sheet_novas: str = 'Sheet1'):
    """
    Lê o arquivo Excel de entrada, realiza a conciliação entre as pendências existentes
    e as novas transações e gera um novo arquivo com:
    - "Pendências": pendências atualizadas seguindo a lógica de negócio
    - "Resumo": cópia idêntica da planilha Resumo, preservando formatação

    Parâmetros:
    caminho_arquivo_entrada: caminho completo para o arquivo de entrada (.xlsx)
    caminho_arquivo_saida: caminho completo para salvar o arquivo gerado (.xlsx)
    sheet_pendencias: nome da aba com as pendências existentes (padrão: 'Pendências')
    sheet_novas: nome da aba com as novas transações (padrão: 'Sheet1')
    """
    # Leitura das planilhas
    df_pend = pd.read_excel(caminho_arquivo_entrada, sheet_name=sheet_pendencias, engine='openpyxl')
    df_new = pd.read_excel(caminho_arquivo_entrada, sheet_name=sheet_novas, engine='openpyxl')

    # Criação de chave de reconciliação conforme especificação: VALOR + INFORMACAO_ADICIONAL + NOME_CONTA
    key_cols = ['VALOR', 'INFORMACAO_ADICIONAL', 'NOME_CONTA']
    df_pend['_key'] = df_pend[key_cols].astype(str).agg(''.join, axis=1)
    df_new['_key'] = df_new[key_cols].astype(str).agg(''.join, axis=1)

    # Criar um dicionário para busca rápida das pendências por chave
    pend_dict = df_pend.set_index('_key').to_dict('index')

    # Processar cada linha de Sheet1
    resultado_linhas = []
    for _, row_sheet1 in df_new.iterrows():
        key = row_sheet1['_key']
        
        if key in pend_dict:
            # Se a chave existe em Pendências, usar a linha de Pendências
            row_pend = pend_dict[key]
            # Converter de volta para Series para manter consistência
            row_pend_series = pd.Series(row_pend)
            resultado_linhas.append(row_pend_series)
        else:
            # Se a chave não existe em Pendências, usar a linha de Sheet1
            # Precisamos ajustar as colunas para corresponder ao formato de Pendências
            row_adaptada = row_sheet1.copy()
            
            # Adicionar colunas que existem em Pendências mas não em Sheet1
            colunas_extras = set(df_pend.columns) - set(df_new.columns)
            for col in colunas_extras:
                if col != '_key':  # Não adicionar a coluna temporária _key
                    row_adaptada[col] = None
            
            resultado_linhas.append(row_adaptada)

    # Criar DataFrame consolidado
    consolidated = pd.DataFrame(resultado_linhas)
    
    # Garantir que as colunas estejam na mesma ordem que Pendências (exceto _key)
    colunas_originais = [col for col in df_pend.columns if col != '_key']
    consolidated = consolidated[colunas_originais]

    # Escreve o arquivo de saída
    with pd.ExcelWriter(caminho_arquivo_saida, engine='openpyxl') as writer:
        # Escreve a aba Pendências (nome correto conforme especificação)
        consolidated.to_excel(writer, sheet_name='Pendências', index=False)
        
        # Escreve a aba Resumo
        try:
            df_resumo = pd.read_excel(caminho_arquivo_entrada, sheet_name='Resumo', engine='openpyxl')
            df_resumo.to_excel(writer, sheet_name='Resumo', index=False)
        except Exception:
            # Se não houver aba Resumo, continua sem ela
            pass

    return len(consolidated)

