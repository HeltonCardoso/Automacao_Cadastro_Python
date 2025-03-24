import pandas as pd

def verificar_medidas_e_pesos(planilha):
    """
    Verifica produtos com:
    - Altura, largura e profundidade devem estar em 'cm'
    - Peso e peso suportado devem estar em 'kg'
    - Ignora o atributo "Com altura ajustável"
    
    Args:
        planilha (str): Caminho do arquivo Excel ou CSV com os dados
        
    Returns:
        DataFrame: Produtos com divergências nas medidas
    """
    
    # Carregar a planilha
    try:
        if planilha.endswith('.xlsx') or planilha.endswith('.xls'):
            df = pd.read_excel(planilha)
        else:
            df = pd.read_csv(planilha, delimiter=';')
    except Exception as e:
        print(f"Erro ao carregar arquivo: {e}")
        return None
    
    # Verificar se existe coluna SKU
    if 'SKU' not in df.columns:
        print("Coluna 'SKU' não encontrada na planilha")
        return None
    
    # Encontrar colunas de atributos
    atributos_nomes = [col for col in df.columns if col.startswith('NomeAtributo')]
    atributos_valores = [col for col in df.columns if col.startswith('ValorAtributo')]
    
    if not atributos_nomes or not atributos_valores:
        print("Estrutura de atributos não encontrada (esperado colunas NomeAtributoX/ValorAtributoX)")
        return None
    
    # Dicionário para armazenar os resultados
    produtos_com_erros = []
    
    # Mapeamento de atributos e unidades esperadas
    regras_validacao = {
        'altura': 'cm',
        'largura': 'cm',
        'profundidade': 'cm',
        'peso': 'kg',
        'peso suportado': 'kg',
        'peso_suportado': 'kg'
    }
    
    # Atributos para ignorar
    atributos_ignorar = ['com altura ajustável', 'altura ajustável']
    
    # Analisar cada produto
    for idx, row in df.iterrows():
        sku = row['SKU']
        erros_produto = []
        
        # Verificar cada par de atributo/valor
        for nome_col, valor_col in zip(atributos_nomes, atributos_valores):
            nome_atributo = str(row[nome_col]).lower().strip() if pd.notna(row[nome_col]) else ''
            valor_atributo = str(row[valor_col]).lower().strip() if pd.notna(row[valor_col]) else ''
            
            # Pular atributos na lista de ignorados
            if any(ignorar in nome_atributo for ignorar in atributos_ignorar):
                continue
            
            # Verificar se o atributo está nas regras de validação
            for atributo_chave, unidade_correta in regras_validacao.items():
                if atributo_chave in nome_atributo:
                    # Verificar unidade
                    if unidade_correta not in valor_atributo:
                        # Verificar se tem unidade errada
                        unidade_errada = 'kg' if unidade_correta == 'cm' else 'cm'
                        if unidade_errada in valor_atributo:
                            erros_produto.append({
                                'Atributo': row[nome_col],  # Nome original (não lower)
                                'Valor': row[valor_col],    # Valor original (não lower)
                                'Erro': f"Unidade {unidade_errada} (deveria ser {unidade_correta})"
                            })
                        elif valor_atributo.replace('.', '').isdigit():
                            erros_produto.append({
                                'Atributo': row[nome_col],
                                'Valor': row[valor_col],
                                'Erro': f"Unidade faltando (adicionar '{unidade_correta}')"
                            })
                    break
        
        # Adicionar ao resultado se houver erros
        if erros_produto:
            for erro in erros_produto:
                produtos_com_erros.append({
                    'SKU': sku,
                    'Atributo': erro['Atributo'],
                    'Valor': erro['Valor'],
                    'Erro': erro['Erro']
                })
    
    return pd.DataFrame(produtos_com_erros)

# (Funções exportar_relatorio e __main__ continuam iguais ao script anterior)

# Função para exportar relatório organizado
def exportar_relatorio(resultado):
    if not resultado.empty:
        # Agrupar erros por SKU
        relatorio = resultado.groupby('SKU').apply(
            lambda x: "\n".join(f"{row['Atributo']}: {row['Valor']} ({row['Erro']})" 
                              for _, row in x.iterrows())
        ).reset_index()
        relatorio.columns = ['SKU', 'Erros']
        return relatorio
    return pd.DataFrame()

# Uso do script
if __name__ == "__main__":
    caminho_planilha = input("Digite o caminho da planilha: ")
    produtos_com_erros = verificar_medidas_e_pesos(caminho_planilha)
    
    if produtos_com_erros is not None:
        if not produtos_com_erros.empty:
            print("\nProdutos com divergências nas medidas:")
            print(produtos_com_erros)
            
            # Relatório organizado
            relatorio = exportar_relatorio(produtos_com_erros)
            print("\nRelatório consolidado:")
            print(relatorio)
            
            # Opção para exportar resultados
            exportar = input("\nDeseja exportar os resultados? (s/n): ").lower()
            if exportar == 's':
                nome_erros = "produtos_com_erros.xlsx"
                produtos_com_erros.to_excel(nome_erros, index=False)
                
                nome_relatorio = "relatorio_consolidado.xlsx"
                relatorio.to_excel(nome_relatorio, index=False)
                
                print(f"Resultados exportados:\n- Detalhado: {nome_erros}\n- Consolidado: {nome_relatorio}")
        else:
            print("Nenhum produto com divergências nas medidas encontrado.")