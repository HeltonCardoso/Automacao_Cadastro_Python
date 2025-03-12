import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from tkinter.ttk import Progressbar
import pandas as pd
import os
import csv
import re
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from datetime import datetime  # Para adicionar timestamps


class TelaComparacaoPrazos:
    def __init__(self, root):
        self.root = root
        self.root.title("VERIFICAÇÃO DE PRAZOS V1.0")
        self.centralizar_janela(800, 650)

        # Adicionar ícone à janela
        try:
            self.root.iconbitmap("icone.ico")  # Substitua "icone.ico" pelo caminho do seu ícone
        except:
            pass  # Ignora se o ícone não for encontrado

        # Frame principal
        frame = tk.Frame(self.root, padx=10, pady=10)
        frame.pack(fill=tk.BOTH, expand=True)

        # Planilha OnClick
        tk.Label(frame, text="PLANILHA ONCLICK:", fg="gray", font=("Arial", 11)).grid(row=0, column=0, sticky=tk.W, pady=5)
        self.entrada_erp = tk.Entry(frame, width=70, font=("Arial", 10))
        self.entrada_erp.grid(row=0, column=1, padx=5, pady=5)
        btn_buscar_erp = tk.Button(frame, text="Buscar", command=lambda: self.selecionar_arquivo(self.entrada_erp), bg="#008CBA", fg="white", font=("Arial", 12, "bold"), bd=0, relief=tk.FLAT)
        btn_buscar_erp.grid(row=0, column=2, padx=5, pady=5)
        btn_buscar_erp.bind("<Enter>", lambda e: btn_buscar_erp.config(bg="#45a049"))
        btn_buscar_erp.bind("<Leave>", lambda e: btn_buscar_erp.config(bg="#008CBA"))

        # Planilha Marketplace
        tk.Label(frame, text="PLANILHA MARKETPLACE:", fg="gray", font=("Arial", 11)).grid(row=1, column=0, sticky=tk.W, pady=5)
        self.entrada_marketplace = tk.Entry(frame, width=70, font=("Arial", 10))
        self.entrada_marketplace.grid(row=1, column=1, padx=5, pady=5)
        btn_buscar_marketplace = tk.Button(frame, text="Buscar", command=lambda: self.selecionar_arquivo(self.entrada_marketplace), bg="#008CBA", fg="white", font=("Arial", 12, "bold"), bd=0, relief=tk.FLAT)
        btn_buscar_marketplace.grid(row=1, column=2, padx=5, pady=5)
        btn_buscar_marketplace.bind("<Enter>", lambda e: btn_buscar_marketplace.config(bg="#45a049"))
        btn_buscar_marketplace.bind("<Leave>", lambda e: btn_buscar_marketplace.config(bg="#008CBA"))

        # Label do Marketplace
        self.label_marketplace = tk.Label(self.root, text="Marketplace: Não identificado", fg="gray", font=("Arial", 11))
        self.label_marketplace.pack(pady=5)

        # Botão Comparar Prazos
        btn_comparar = tk.Button(
            self.root, 
            text="COMPARAR PRAZOS", 
            command=self.abrir_mapeamento_colunas, 
            bg="#008CBA", 
            fg="white", 
            font=("Arial", 12, "bold"), 
            bd=0, 
            relief=tk.FLAT
        )
        btn_comparar.pack(pady=10)
        btn_comparar.bind("<Enter>", lambda e: btn_comparar.config(bg="#45a049"))
        btn_comparar.bind("<Leave>", lambda e: btn_comparar.config(bg="#008CBA"))

        # Barra de Progresso
        self.progresso = Progressbar(self.root, orient="horizontal", length=600, mode="determinate")
        self.progresso.pack(pady=5)

        # Label de Status
        self.status_label = tk.Label(self.root, text="Tela de log", fg="gray", font=("Arial", 10))
        self.status_label.pack(pady=5)

        # Área de Log
        self.log_area = scrolledtext.ScrolledText(self.root, width=80, height=12, bg="white", state="disabled", font=("Courier", 10))
        self.log_area.pack(pady=10)

        # Configuração de tags para cores
        self.log_area.tag_config("info", foreground="blue", font=("Courier", 11, "bold"))
        self.log_area.tag_config("erro", foreground="red", font=("Courier", 11, "bold"))
        self.log_area.tag_config("sucesso", foreground="green", font=("Courier", 11, "bold"))
        self.log_area.tag_config("aviso", foreground="orange", font=("Courier", 11, "bold"))
        self.log_area.tag_config("divergencia", foreground="red", font=("Courier", 11, "bold"))

        # Botão Limpar Log
        btn_limpar = tk.Button(
            self.root, 
            text="LIMPAR LOG", 
            command=self.limpar_log, 
            bg="#008CBA", 
            fg="white", 
            font=("Arial", 12, "bold"), 
            bd=0, 
            relief=tk.FLAT
        )
        btn_limpar.pack(pady=5)
        btn_limpar.bind("<Enter>", lambda e: btn_limpar.config(bg="#45a049"))
        btn_limpar.bind("<Leave>", lambda e: btn_limpar.config(bg="#008CBA"))

        # Botão Fechar
        btn_fechar = tk.Button(
            self.root, 
            text="FECHAR", 
            command=self.root.destroy, 
            bg="#f44336", 
            fg="white", 
            font=("Arial", 12, "bold"), 
            bd=0, 
            relief=tk.FLAT
        )
        btn_fechar.pack(pady=10)
        btn_fechar.bind("<Enter>", lambda e: btn_fechar.config(bg="#e53935"))
        btn_fechar.bind("<Leave>", lambda e: btn_fechar.config(bg="#f44336"))

        # Rodapé
        frame_rodape = tk.Frame(root)
        frame_rodape.pack(side=tk.BOTTOM, fill=tk.X, pady=10)

        # Relógio no rodapé
        self.relogio = tk.Label(frame_rodape, font=("Arial", 10), fg="gray")
        self.relogio.pack(side=tk.RIGHT, padx=10)
        self.atualizar_relogio()

    # Restante do código...

        # Dicionário de mapeamento de colunas por marketplace
        self.mapa_marketplaces = {
            "Wake": {
                "cod_barra": "EAN",  # Coluna do marketplace
                "prazo": "Prazo Manuseio (Dias)",  # Coluna do prazo
                "chave_comparacao": "EAN",  # Chave de comparação (EAN ou SellerSku)
                "prazo_erp": "DIAS P/ ENTREGA"  # Coluna de prazo no ERP
            },
            "Tray": {
                "cod_barra": "EAN",
                "prazo": "Disponibilidade",
                "chave_comparacao": "EAN",
                "prazo_erp": "SITE_DISPONIBILIDADE"  # Coluna de prazo no ERP
            },
            "Shoppe": {
                "cod_barra": "EAN_shoppe",
                "prazo": "Disponibilidade_shoppe",
                "chave_comparacao": "EAN",
                "prazo_erp": "SITE_DISPONIBILIDADE"  # Coluna de prazo no ERP
            },
            "Mobly": {
                "cod_barra": "SellerSku",  # Mobly usa SellerSku
                "prazo": "SupplierDeliveryTime",  # Coluna de prazo no marketplace
                "chave_comparacao": "SellerSku",  # Chave de comparação é SellerSku
                "prazo_erp": "SITE_DISPONIBILIDADE"  # Coluna de prazo no ERP para Mobly
            },
            "MadeiraMadeira": {
                "cod_barra": "EAN",  # Coluna do EAN no marketplace
                "prazo": "Prazo expedição",  # Coluna do prazo no marketplace
                "chave_comparacao": "EAN",  # Chave de comparação (EAN)
                "prazo_erp": "SITE_DISPONIBILIDADE"  # Coluna de prazo no ERP
            },
            "WebContinental": {  # Novo marketplace
                "cod_barra": "EAN",  # Coluna do EAN no WebContinental
                "prazo": "Crossdoc",  # Coluna do prazo no WebContinental
                "chave_comparacao": "EAN",  # Chave de comparação (EAN)
                "prazo_erp": "SITE_DISPONIBILIDADE"  # Coluna de prazo no ERP
            }
        }

    
    def atualizar_relogio(self):
        agora = datetime.now()
        data_hora = agora.strftime("%d/%m/%Y %H:%M:%S")
        self.relogio.config(text=data_hora)
        self.root.after(1000, self.atualizar_relogio)  # Atualiza a cada 1 segundo

    def centralizar_janela(self, largura, altura):
        largura_tela = self.root.winfo_screenwidth()
        altura_tela = self.root.winfo_screenheight()
        pos_x = (largura_tela // 2) - (largura // 2)
        pos_y = (altura_tela // 2) - (altura // 2)
        self.root.geometry(f"{largura}x{altura}+{pos_x}+{pos_y}")

    def selecionar_arquivo(self, entrada):
        self.root.attributes('-topmost', False)
        caminho = filedialog.askopenfilename(
            parent=self.root,
            filetypes=(("Arquivos Excel", "*.xlsx *.xls *.csv"),("Todos os arquivos", "*.*"))
        )
        self.root.attributes('-topmost', True)

        if caminho:
            entrada.delete(0, tk.END)
            entrada.insert(0, caminho)

            if entrada == self.entrada_marketplace:
                self.identificar_marketplace(caminho)

    def identificar_marketplace(self, caminho):
        try:
            df = self.ler_arquivo(caminho)
            colunas = df.columns.tolist()

            for marketplace, colunas_mapeadas in self.mapa_marketplaces.items():
                if colunas_mapeadas["prazo"] in colunas:
                    self.label_marketplace.config(text=f"Marketplace: {marketplace}")
                    return

            # Verifica se é o MadeiraMadeira (caso o nome da coluna seja diferente)
            if "EAN" in colunas and "Prazo de Entrega" in colunas:
                self.label_marketplace.config(text="Marketplace: MadeiraMadeira")
                return

            self.label_marketplace.config(text="Marketplace: Não identificado")
            self.log(f"[AVISO] Não foi possível identificar o marketplace. Verifique as colunas da planilha.", "aviso")
        except Exception as e:
            self.label_marketplace.config(text="Marketplace: Erro ao identificar")
            self.log(f"[ERRO] Erro ao identificar o marketplace: {e}", "erro")

    def ler_arquivo(self, caminho):
        try:
            if not os.path.exists(caminho):
                raise ValueError(f"O arquivo '{caminho}' não existe.")
            
            if caminho.endswith('.csv'):
                # Detectar o delimitador
                delimitador = self.detectar_delimitador(caminho)
                
                # Ler o arquivo CSV, ignorando linhas mal formatadas
                return pd.read_csv(caminho, delimiter=delimitador, encoding='latin1', on_bad_lines="skip")
            
            elif caminho.endswith(('.xls', '.xlsx')):
                return pd.read_excel(caminho, engine='openpyxl' if caminho.endswith('.xlsx') else 'xlrd')
            
            else:
                raise ValueError("Formato de arquivo não suportado. Use .xls, .xlsx ou .csv.")
        except Exception as e:
            raise ValueError(f"Erro ao ler o arquivo: {e}")

    def detectar_delimitador(self, caminho):
        with open(caminho, 'r', encoding='latin1') as file:
            sniffer = csv.Sniffer()
            dialect = sniffer.sniff(file.readline())
            return dialect.delimiter

    def extrair_numeros(self, texto):
        """
        Extrai números de uma string usando regex.
        Exemplo: "Disponível em 60 dias úteis" -> 60
        """
        if pd.isna(texto):  # Verifica se o valor é NaN
            return 0
        numeros = re.findall(r'\d+', str(texto))  # Encontra todos os números na string
        return int(numeros[0]) if numeros else 0  # Retorna o primeiro número ou 0 se não houver

    def log(self, mensagem, tipo="info"):
        """
        Adiciona uma mensagem à área de logs com formatação e cores.
        Tipos: info (azul), erro (vermelho), sucesso (verde), aviso (laranja), divergencia (vermelho e negrito).
        """
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        mensagem_formatada = f"[{timestamp}] {mensagem}\n"

        self.log_area.config(state="normal")
        self.log_area.insert(tk.END, mensagem_formatada, tipo)
        self.log_area.config(state="disabled")
        self.log_area.yview(tk.END)
        self.root.update_idletasks()

    def aplicar_estilo_botoes(self):
        # Configuração de estilo para os botões
        self.style.configure("TButton", 
                            font=("Arial", 10, "bold"), 
                            padding=10, 
                            relief=tk.FLAT, 
                            background="#008CBA", 
                            foreground="white")
        self.style.map("TButton",
                    background=[("active", "#45a049")],  # Cor ao passar o mouse
                    foreground=[("active", "white")])
        
        # Estilo específico para o botão "Fechar"
        self.style.configure("Fechar.TButton", 
                            background="#f44336", 
                            foreground="white")
        self.style.map("Fechar.TButton",
                    background=[("active", "#e53935")],  # Cor ao passar o mouse
                    foreground=[("active", "white")])

    def limpar_log(self):
        """Limpa a área de logs."""
        self.log_area.config(state="normal")
        self.log_area.delete(1.0, tk.END)
        self.log_area.config(state="disabled")

    def gerar_relatorio_detalhado(self, df_comparacao, divergencias, marketplace, coluna_chave_erp):
        try:
            self.log("\n Gerando relatório detalhado...", "aviso")
            
            # Passo 1: Obter o caminho do diretório atual (onde o script está sendo executado)
            caminho_atual = os.path.dirname(os.path.abspath(__file__))
            
            # Passo 2: Subir um nível no diretório
            caminho_pai = os.path.dirname(caminho_atual)
            
            # Passo 3: Criar a pasta RELATORIO_DETALHADO no nível superior
            pasta_relatorio = os.path.join(caminho_pai, "RELATORIO_DETALHADO")
            if not os.path.exists(pasta_relatorio):
                os.makedirs(pasta_relatorio)
            
            # Passo 4: Dividir as divergências em lotes menores
            tamanho_lote = 500  # Número de produtos por PDF
            lotes = [divergencias[i:i + tamanho_lote] for i in range(0, len(divergencias), tamanho_lote)]
            
            # Passo 5: Gerar um PDF para cada lote
            for indice, lote in enumerate(lotes):
                caminho_relatorio = os.path.join(pasta_relatorio, f"Relatorio_Detalhado_Parte_{indice + 1}.pdf")
                
                with PdfPages(caminho_relatorio) as pdf:
                    # Página 1: Estatísticas Gerais
                    plt.figure(figsize=(8.27, 11.69))  # Formato A4 em retrato (largura x altura em polegadas)
                    plt.axis('off')
                    plt.title(f"Relatório Detalhado de Comparação de Prazos - Parte {indice + 1}", fontsize=16, pad=20)
                    
                    total_itens = len(df_comparacao)
                    total_divergencias = len(divergencias)
                    percentual_divergencias = (total_divergencias / total_itens) * 100 if total_itens > 0 else 0
                    
                    estatisticas = (
                        f"Total de produtos verificados: {total_itens}\n"
                        f"Total de produtos com divergências: {total_divergencias}\n"
                        f"Percentual de divergências: {percentual_divergencias:.2f}%\n"
                        f"Produtos nesta parte: {len(lote)}\n"
                    )
                    
                    plt.text(0.1, 0.8, estatisticas, fontsize=12, ha="left", va="top", wrap=True)
                    pdf.savefig()
                    plt.close()
                    
                    # Página 2: Gráfico de Barras (Divergências)
                    plt.figure(figsize=(8.27, 11.69))  # Formato A4 em retrato
                    plt.bar(["Com Divergências", "Sem Divergências"], [len(lote), total_itens - len(lote)], color=["red", "green"])
                    plt.title("Distribuição de Divergências", fontsize=16)
                    plt.ylabel("Quantidade de Produtos")
                    pdf.savefig()
                    plt.close()
                    
                    # Páginas 3 em diante: Detalhes das Divergências
                    detalhes_por_pagina = []
                    detalhes = ""
                    for index, row in lote.iterrows():
                        detalhes_produto = (
                            f"Código: {row[coluna_chave_erp]}\n"
                            f"Prazo ERP: {row['DIAS P/ ENTREGA_ERP']}\n"
                            f"Prazo Marketplace: {row['DIAS P/ ENTREGA_MARKETPLACE']}\n"
                            f"Diferença: {row['DIFERENCA_PRAZO']}\n"
                            f"{'-' * 40}\n"
                        )
                        
                        # Se o texto ultrapassar o limite da página, criar uma nova página
                        if len(detalhes) + len(detalhes_produto) > 1000:  # Aumentei o limite para 3000 caracteres
                            detalhes_por_pagina.append(detalhes)
                            detalhes = detalhes_produto  # Reinicia o texto para a próxima página
                        else:
                            detalhes += detalhes_produto
                    
                    # Adicionar os detalhes restantes (última página)
                    if detalhes:
                        detalhes_por_pagina.append(detalhes)
                    
                    # Adicionar cada página de detalhes ao PDF
                    for i, pagina in enumerate(detalhes_por_pagina):
                        plt.figure(figsize=(8.27, 11.69))  # Formato A4 em retrato
                        plt.axis('off')
                        plt.title(f"Detalhes das Divergências - Página {i + 1}", fontsize=16, pad=20)
                        plt.text(0.1, 0.8, pagina, fontsize=10, ha="left", va="top", wrap=True)
                        pdf.savefig()
                        plt.close()
                
                self.log(f"Relatório detalhado salvo em: {caminho_relatorio}", "sucesso")
            
            messagebox.showinfo("Relatório Gerado", f"Relatórios detalhados salvos em: {pasta_relatorio}", parent=self.root)
        
        except Exception as e:
            self.log(f"Erro ao gerar relatório detalhado: {e}", "erro")
            messagebox.showerror("Erro", f"Erro ao gerar relatório detalhado: {e}", parent=self.root)    
    
    

    def comparar_prazos(self, planilha_erp, planilha_marketplace):
        try:
            self.log("Processando as planilhas...", "info")
            self.progresso["value"] = 10
            self.root.update_idletasks()

            # Ler as planilhas
            df_erp = self.ler_arquivo(planilha_erp)
            df_marketplace = self.ler_arquivo(planilha_marketplace)

            self.progresso["value"] = 20
            self.root.update_idletasks()

            # Identificar o marketplace
            marketplace = self.label_marketplace.cget("text").replace("Marketplace: ", "")
            if marketplace not in self.mapa_marketplaces:
                raise ValueError("Marketplace não identificado ou não suportado.")

            mapeamento = self.mapa_marketplaces[marketplace]

            # Verificar se as colunas necessárias existem
            if mapeamento["cod_barra"] not in df_marketplace.columns or mapeamento["prazo"] not in df_marketplace.columns:
                raise ValueError(f"Colunas do marketplace não encontradas. Verifique se as colunas '{mapeamento['cod_barra']}' e '{mapeamento['prazo']}' existem.")

            if mapeamento["prazo_erp"] not in df_erp.columns:
                raise ValueError(f"Coluna de prazo do ERP não encontrada. Verifique se a coluna '{mapeamento['prazo_erp']}' existe.")

            # Renomear colunas
            df_marketplace.rename(columns={
                mapeamento["cod_barra"]: "COD_COMPARACAO",
                mapeamento["prazo"]: "DIAS P/ ENTREGA_MARKETPLACE"
            }, inplace=True)

            df_erp.rename(columns={
                mapeamento["prazo_erp"]: "DIAS P/ ENTREGA_ERP"
            }, inplace=True)

            # Tratar o Tray
            if marketplace == "Tray":
                # Extrair números da coluna de prazo
                df_marketplace["DIAS P/ ENTREGA_MARKETPLACE"] = df_marketplace["DIAS P/ ENTREGA_MARKETPLACE"].apply(self.extrair_numeros)
                self.log(f"Valores de prazo no Tray: {df_marketplace['DIAS P/ ENTREGA_MARKETPLACE'].unique()}", "info")

            # Remover .0 dos valores de COD_COMPARACAO
            df_marketplace["COD_COMPARACAO"] = df_marketplace["COD_COMPARACAO"].astype(str).str.replace(r"\.0$", "", regex=True)

            # Remover valores 'nan'
            df_marketplace = df_marketplace[df_marketplace["COD_COMPARACAO"] != "nan"]

            # Converter códigos de barras para string e remover espaços
            coluna_chave_erp = "COD AUXILIAR" if marketplace == "Mobly" else "COD BARRA"
            df_erp[coluna_chave_erp] = df_erp[coluna_chave_erp].astype(str).str.strip()
            df_marketplace["COD_COMPARACAO"] = df_marketplace["COD_COMPARACAO"].astype(str).str.strip()

            # Realizar o merge
            df_comparacao = pd.merge(df_erp, df_marketplace, left_on=coluna_chave_erp, right_on="COD_COMPARACAO", suffixes=("_ERP", "_MARKETPLACE"))

            # Converter as colunas de prazo para números
            df_comparacao["DIAS P/ ENTREGA_ERP"] = pd.to_numeric(df_comparacao["DIAS P/ ENTREGA_ERP"], errors="coerce").fillna(0)
            df_comparacao["DIAS P/ ENTREGA_MARKETPLACE"] = pd.to_numeric(df_comparacao["DIAS P/ ENTREGA_MARKETPLACE"], errors="coerce").fillna(0)

            self.log("Calculando diferenças de prazos...", "info")
            self.progresso["value"] = 70
            self.root.update_idletasks()

            df_comparacao["DIFERENCA_PRAZO"] = df_comparacao["DIAS P/ ENTREGA_MARKETPLACE"] - df_comparacao["DIAS P/ ENTREGA_ERP"]

            divergencias = df_comparacao[df_comparacao["DIFERENCA_PRAZO"] != 0]

            pasta_destino = os.path.join(os.getcwd(), "COMPARACAO_PRAZOS")
            if not os.path.exists(pasta_destino):
                os.makedirs(pasta_destino)

            caminho_arquivo = os.path.join(pasta_destino, "Comparacao_Prazos.xlsx")
            self.log("Salvando divergências...", "info")
            self.progresso["value"] = 90
            self.root.update_idletasks()

            divergencias.to_excel(caminho_arquivo, index=False)

            self.progresso["value"] = 100
            self.root.update_idletasks()

            self.log("Produtos com divergência:", "aviso")
            for index, row in divergencias.iterrows():
                # Exibe apenas o EAN, sem timestamp
                self.log_area.config(state="normal")
                self.log_area.insert(tk.END, f"{row[coluna_chave_erp]}\n", "divergencia")
                self.log_area.config(state="disabled")
                self.log_area.yview(tk.END)
                self.root.update_idletasks()

            total_itens = len(df_comparacao)
            total_divergencias = len(divergencias)
            self.log(f"Total de EAN/SKUS verificados: {total_itens}", "info")
            self.log(f"Total de EAN/SKUS com divergência: {total_divergencias}", "info")

            # Gerar relatório detalhado
            self.gerar_relatorio_detalhado(df_comparacao, divergencias, marketplace, coluna_chave_erp)

        except Exception as e:
            self.log(f"[ERRO] Ocorreu um erro durante a comparação: {e}", "erro")
            messagebox.showerror("Erro", f"Ocorreu um erro durante a comparação: {e}", parent=self.root)
            self.progresso["value"] = 0
            self.root.update_idletasks()

    def abrir_mapeamento_colunas(self):
        self.comparar_prazos(self.entrada_erp.get(), self.entrada_marketplace.get())

if __name__ == "__main__":
    root = tk.Tk()
    app = TelaComparacaoPrazos(root)
    root.mainloop()