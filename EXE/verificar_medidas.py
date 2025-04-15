import pandas as pd
import os
from tkinter import filedialog
from tkinter import Tk

def verificar_medidas_e_pesos(planilha):
    """
    Verifica produtos com:
    - Medidas em 'cm' (altura, largura, profundidade)
    - Pesos em 'kg' (peso, peso suportado)
    - Tamanho de TV em Número" (ex: 70")
    - Gera relatório de "Formato do Móvel" com Nome do Produto
    
    Args:
        planilha (str): Caminho do arquivo Excel ou CSV
        
    Returns:
        tuple: (DataFrame com erros, DataFrame com formato dos móveis)
    """
    try:
        # Carregar a planilha
        if planilha.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(planilha)
        else:
            df = pd.read_csv(planilha, delimiter=';')
    except Exception as e:
        print(f"Erro ao carregar arquivo: {e}")
        return None, None

    # Verificar colunas obrigatórias
    colunas_obrigatorias = ['SKU', 'Nome']
    for col in colunas_obrigatorias:
        if col not in df.columns:
            print(f"Coluna '{col}' não encontrada na planilha")
            return None, None

    # Inicializar listas de resultados
    produtos_com_erros = []
    produtos_com_formato = []

    # Procurar colunas de atributos
    cols_nome_atrib = [c for c in df.columns if c.startswith('NomeAtributo')]
    cols_valor_atrib = [c for c in df.columns if c.startswith('ValorAtributo')]

    if not cols_nome_atrib or not cols_valor_atrib:
        print("Colunas de atributos não encontradas (NomeAtributoX/ValorAtributoX)")
        return None, None

    # Processar cada produto
    for _, row in df.iterrows():
        sku = row['SKU']
        nome_produto = row['Nome']

        # Procurar por "Formato do Móvel"
        for nome_col, valor_col in zip(cols_nome_atrib, cols_valor_atrib):
            nome_atrib = str(row[nome_col]).lower().strip()
            valor_atrib = str(row[valor_col]).strip() if pd.notna(row[valor_col]) else ''

            if 'formato do móvel' in nome_atrib and valor_atrib:
                produtos_com_formato.append({
                    'SKU': sku,
                    'Nome do Produto': nome_produto,
                    'Atributo': row[nome_col],  # Nome original
                    'Valor': valor_atrib
                })

    # Converter para DataFrames
    df_erros = pd.DataFrame(produtos_com_erros) if produtos_com_erros else pd.DataFrame()
    df_formato = pd.DataFrame(produtos_com_formato) if produtos_com_formato else pd.DataFrame()

    return df_erros, df_formato

def exportar_relatorios(erros, formato):
    """Exporta todos os relatórios para Excel"""
    if erros.empty and formato.empty:
        print("Nenhum dado para exportar")
        return

    root = Tk()
    root.withdraw()
    pasta = filedialog.askdirectory(title="Selecione a pasta para salvar os relatórios")
    
    if not pasta:
        print("Exportação cancelada")
        return

    # Exportar relatório de formato do móvel (sempre que houver dados)
    if not formato.empty:
        caminho_formato = os.path.join(pasta, "relatorio_formato_moveis.xlsx")
        formato.to_excel(caminho_formato, index=False)
        print(f"\nRelatório de formato salvo em:\n{caminho_formato}")

    # Exportar relatórios de erros (se existirem)
    if not erros.empty:
        caminho_erros = os.path.join(pasta, "produtos_com_erros.xlsx")
        erros.to_excel(caminho_erros, index=False)
        
        # Relatório consolidado
        relatorio = erros.groupby('SKU').apply(
            lambda x: "\n".join(f"{row['Atributo']}: {row['Valor']} ({row['Erro']})" 
                              for _, row in x.iterrows())
        ).reset_index(name='Erros')
        
        caminho_consolidado = os.path.join(pasta, "relatorio_consolidado.xlsx")
        relatorio.to_excel(caminho_consolidado, index=False)
        
        print(f"\nRelatórios de erros salvos em:\n{caminho_erros}\n{caminho_consolidado}")

if __name__ == "__main__":
    caminho = input("Digite o caminho da planilha: ")
    df_erros, df_formato = verificar_medidas_e_pesos(caminho)

    # Mostrar resultados
    if df_formato is not None:
        if not df_formato.empty:
            print("\n=== Produtos com Formato do Móvel ===")
            print(df_formato[['SKU', 'Nome do Produto', 'Atributo', 'Valor']])
        else:
            print("\nNenhum produto com 'Formato do Móvel' encontrado")

        if not df_erros.empty:
            print("\n=== Produtos com Erros ===")
            print(df_erros)
        else:
            print("\nNenhum erro encontrado nas medidas")

        # Exportar
        if input("\nExportar resultados? (s/n): ").lower() == 's':
            exportar_relatorios(df_erros, df_formato)