import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from tkinter.ttk import Progressbar
import pandas as pd
import os
import csv
import re
from datetime import datetime
from PIL import Image, ImageTk

class TelaComparacaoPrazos:
    def __init__(self, root):
        self.root = root
        self.root.title("VERIFICAÇÃO DE PRAZOS V1.0")
        self.root.resizable(False, False)
        self.centralizar_janela(800, 680)
    
        # Frame para marketplace identificado e imagem
        frame_marketplace = tk.Frame(self.root)
        frame_marketplace.pack(pady=5)
        
        # Label para imagem do marketplace 
        self.imagem_marketplace = tk.Label(frame_marketplace)
        self.imagem_marketplace.pack(side=tk.LEFT, padx=5)
        self.imagens_carregadas = {}
        try:
            self.root.iconbitmap("IMG/icone.ico")
        except:
            pass

        # Configurar interface principal
        self.configurar_interface()
        self.imagens_marketplaces = self.carregar_imagens_marketplaces()

    def configurar_interface(self):
        # Frame principal
        frame = tk.Frame(self.root, padx=10, pady=10)
        frame.pack(fill=tk.BOTH, expand=True)

        # Planilha OnClick
        tk.Label(frame, text="PLANILHA ONCLICK:", fg="gray", font=("Arial", 10)).grid(row=0, column=0, sticky=tk.W, pady=5)
        self.entrada_erp = tk.Entry(frame, width=65, font=("Arial", 10))
        self.entrada_erp.grid(row=0, column=1, padx=3, pady=3)
        btn_buscar_erp = tk.Button(frame, text="SELECIONAR", command=lambda: self.selecionar_arquivo(self.entrada_erp), bg="#008CBA", fg="white", font=("Arial", 10, "bold"), bd=4, relief=tk.FLAT)
        btn_buscar_erp.grid(row=0, column=2, padx=5, pady=5)
        btn_buscar_erp.bind("<Enter>", lambda e: btn_buscar_erp.config(bg="#0a2040"))
        btn_buscar_erp.bind("<Leave>", lambda e: btn_buscar_erp.config(bg="#008CBA"))

        # Planilha Marketplace
        tk.Label(frame, text="PLANILHA MARKETPLACE:", fg="gray", font=("Arial", 10)).grid(row=1, column=0, sticky=tk.W, pady=5)
        self.entrada_marketplace = tk.Entry(frame, width=65, font=("Arial", 10))
        self.entrada_marketplace.grid(row=1, column=1, padx=3, pady=3)
        btn_buscar_marketplace = tk.Button(frame, text="SELECIONAR", command=lambda: self.selecionar_arquivo(self.entrada_marketplace), bg="#008CBA", fg="white", font=("Arial", 10, "bold"), bd=4, relief=tk.FLAT)
        btn_buscar_marketplace.grid(row=1, column=2, padx=5, pady=5)
        btn_buscar_marketplace.bind("<Enter>", lambda e: btn_buscar_marketplace.config(bg="#0a2040"))
        btn_buscar_marketplace.bind("<Leave>", lambda e: btn_buscar_marketplace.config(bg="#008CBA"))

        # Frame para marketplace identificado e imagem
        frame_marketplace = tk.Frame(self.root)
        frame_marketplace.pack(pady=5)
        
        # Label do Marketplace
        self.label_marketplace = tk.Label(frame_marketplace, text="Marketplace: Não identificado", fg="gray", font=("Arial", 10))
        self.label_marketplace.pack(side=tk.LEFT, padx=5)
        
        # Label para imagem do marketplace
        self.imagem_marketplace = tk.Label(frame_marketplace)
        self.imagem_marketplace.pack(side=tk.LEFT, padx=5)
        self.imagem_referencia = None
        
        # Frame para o botão de comparar
        frame_botao = tk.Frame(self.root)
        frame_botao.pack(pady=10)
        
        # Botão Comparar Prazos
        btn_comparar = tk.Button(
            frame_botao, 
            text="COMPARAR PRAZOS", 
            command=self.abrir_mapeamento_colunas, 
            bg="#008CBA", 
            fg="white", 
            font=("Arial", 10, "bold"),
            bd=4, 
            relief=tk.FLAT
        )
        btn_comparar.pack(side=tk.LEFT, padx=5)
        btn_comparar.bind("<Enter>", lambda e: btn_comparar.config(bg="#0a2040"))
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
            font=("Arial", 10, "bold"),
            bd=4, 
            relief=tk.FLAT
        )
        btn_limpar.pack(pady=5)
        btn_limpar.bind("<Enter>", lambda e: btn_limpar.config(bg="#0a2040"))
        btn_limpar.bind("<Leave>", lambda e: btn_limpar.config(bg="#008CBA"))

        # Botão Fechar
        btn_fechar = tk.Button(
            self.root, 
            text="FECHAR", 
            command=self.root.destroy, 
            bg="#f44336", 
            fg="white", 
            font=("Arial", 10, "bold"),
            bd=4, 
            relief=tk.FLAT
        )
        btn_fechar.pack(pady=10)
        btn_fechar.bind("<Enter>", lambda e: btn_fechar.config(bg="#e53935"))
        btn_fechar.bind("<Leave>", lambda e: btn_fechar.config(bg="#f44336"))

        # Rodapé
        frame_rodape = tk.Frame(self.root)
        frame_rodape.pack(side=tk.BOTTOM, fill=tk.X, pady=10)

        # Relógio no rodapé
        self.relogio = tk.Label(frame_rodape, font=("Arial", 10), fg="gray")
        self.relogio.pack(side=tk.RIGHT, padx=10)
        self.atualizar_relogio()

        # Dicionário de mapeamento de colunas por marketplace
        self.mapa_marketplaces = {
            "Wake": {
                "cod_barra": "EAN",
                "prazo": "Prazo Manuseio (Dias)",
                "chave_comparacao": "EAN",
                "prazo_erp": "DIAS P/ ENTREGA",
                "imagem": "wake.png"
            },
            "Tray": {
                "cod_barra": "EAN",
                "prazo": "Disponibilidade",
                "chave_comparacao": "EAN",
                "prazo_erp": "SITE_DISPONIBILIDADE",
                "imagem": "tray.png"
            },
            "Shoppe": {
                "cod_barra": "EAN_shoppe",
                "prazo": "Disponibilidade_shoppe",
                "chave_comparacao": "EAN",
                "prazo_erp": "SITE_DISPONIBILIDADE",
                "imagem": "shoppe.png"
            },
            "Mobly": {
                "cod_barra": "SellerSku",
                "prazo": "SupplierDeliveryTime",
                "chave_comparacao": "SellerSku",
                "prazo_erp": "SITE_DISPONIBILIDADE",
                "imagem": "mobly.png"
            },
            "MadeiraMadeira": {
                "cod_barra": "EAN",
                "prazo": "Prazo expedição",
                "chave_comparacao": "EAN",
                "prazo_erp": "SITE_DISPONIBILIDADE",
                "imagem": "madeiramadeira.png"
            },
            "WebContinental": {
                "cod_barra": "EAN",
                "prazo": "Crossdoc",
                "chave_comparacao": "EAN",
                "prazo_erp": "SITE_DISPONIBILIDADE",
                "imagem": "webcontinental.png"
            }
        }

    def carregar_imagens_marketplaces(self):
        """Carrega as imagens dos marketplaces da pasta IMG"""
        imagens = {}
        marketplaces = {
            "Wake": "wake.png",
            "Tray": "tray.png",
            "Shoppe": "shoppe.png",
            "Mobly": "mobly.png",
            "MadeiraMadeira": "madeiramadeira.png",
            "WebContinental": "webcontinental.png"
        }
        
        try:
            caminho_pasta_img = os.path.join(os.path.dirname(os.path.abspath(__file__)), "IMG")
            
            for nome, arquivo in marketplaces.items():
                caminho_imagem = os.path.join(caminho_pasta_img, arquivo)
                if os.path.exists(caminho_imagem):
                    try:
                        img = Image.open(caminho_imagem)
                        img = img.resize((50, 50), Image.LANCZOS)
                        photo_img = ImageTk.PhotoImage(img)
                        imagens[nome] = photo_img
                        if not hasattr(self, 'imagens_salvas'):
                            self.imagens_salvas = []
                        self.imagens_salvas.append(photo_img)
                    except:
                        pass
        except:
            pass

        return imagens

    def atualizar_imagem_marketplace(self, marketplace):
        """Atualiza a imagem do marketplace exibida na interface"""
        if marketplace and marketplace in self.imagens_marketplaces and self.imagens_marketplaces[marketplace] is not None:
            self.imagem_referencia = self.imagens_marketplaces[marketplace]
            self.imagem_marketplace.config(image=self.imagem_referencia)
        else:
            self.imagem_marketplace.config(image='')
            self.imagem_referencia = None

    def atualizar_relogio(self):
        agora = datetime.now()
        data_hora = agora.strftime("%d/%m/%Y %H:%M:%S")
        self.relogio.config(text=data_hora)
        self.root.after(1000, self.atualizar_relogio)

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
                    self.atualizar_imagem_marketplace(marketplace)
                    return

            if "EAN" in colunas and "Prazo de Entrega" in colunas:
                self.label_marketplace.config(text="Marketplace: MadeiraMadeira")
                self.atualizar_imagem_marketplace("MadeiraMadeira")
                return

            self.label_marketplace.config(text="Marketplace: Não identificado")
            self.atualizar_imagem_marketplace(None)
            self.log(f"[AVISO] Não foi possível identificar o marketplace. Verifique as colunas da planilha.", "aviso")
        except Exception as e:
            self.label_marketplace.config(text="Marketplace: Erro ao identificar")
            self.atualizar_imagem_marketplace(None)
            self.log(f"[ERRO] Erro ao identificar o marketplace: {e}", "erro")

    def ler_arquivo(self, caminho):
        try:
            if not os.path.exists(caminho):
                raise ValueError(f"O arquivo '{caminho}' não existe.")
            
            if caminho.endswith('.csv'):
                delimitador = self.detectar_delimitador(caminho)
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
        if pd.isna(texto):
            return 0
        numeros = re.findall(r'\d+', str(texto))
        return int(numeros[0]) if numeros else 0

    def log(self, mensagem, tipo="info"):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        mensagem_formatada = f"[{timestamp}] {mensagem}\n"

        self.log_area.config(state="normal")
        self.log_area.insert(tk.END, mensagem_formatada, tipo)
        self.log_area.config(state="disabled")
        self.log_area.yview(tk.END)
        self.root.update_idletasks()

    def limpar_log(self):
        self.log_area.config(state="normal")
        self.log_area.delete(1.0, tk.END)
        self.log_area.config(state="disabled")

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

            # Verificar colunas necessárias
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

            # Tratar a Tray
            if marketplace == "Tray":
                df_marketplace["DIAS P/ ENTREGA_MARKETPLACE"] = df_marketplace["DIAS P/ ENTREGA_MARKETPLACE"].apply(self.extrair_numeros)
                self.log(f"Valores de prazo no Tray: {df_marketplace['DIAS P/ ENTREGA_MARKETPLACE'].unique()}", "info")

            # Limpar dados
            df_marketplace["COD_COMPARACAO"] = df_marketplace["COD_COMPARACAO"].astype(str).str.replace(r"\.0$", "", regex=True)
            df_marketplace = df_marketplace[df_marketplace["COD_COMPARACAO"] != "nan"]

            # Converter códigos de barras
            coluna_chave_erp = "COD AUXILIAR" if marketplace == "Mobly" else "COD BARRA"
            df_erp[coluna_chave_erp] = df_erp[coluna_chave_erp].astype(str).str.strip()
            df_marketplace["COD_COMPARACAO"] = df_marketplace["COD_COMPARACAO"].astype(str).str.strip()

            # Realizar o merge
            df_comparacao = pd.merge(
                df_erp,
                df_marketplace, 
                left_on=coluna_chave_erp,
                right_on="COD_COMPARACAO",
                suffixes=("_ERP", "_MARKETPLACE")
                    )

            # Converter prazos para números
            df_comparacao["DIAS P/ ENTREGA_ERP"] = pd.to_numeric(df_comparacao["DIAS P/ ENTREGA_ERP"], errors="coerce").fillna(0)
            df_comparacao["DIAS P/ ENTREGA_MARKETPLACE"] = pd.to_numeric(df_comparacao["DIAS P/ ENTREGA_MARKETPLACE"], errors="coerce").fillna(0)

            self.log("Calculando diferenças de prazos...", "info")
            self.progresso["value"] = 70
            self.root.update_idletasks()

            df_comparacao["DIFERENCA_PRAZO"] = df_comparacao["DIAS P/ ENTREGA_MARKETPLACE"] - df_comparacao["DIAS P/ ENTREGA_ERP"]
###############################################################################################

            divergencias = df_comparacao[df_comparacao["DIFERENCA_PRAZO"] != 0]

            # Abrir diálogo para escolher local de salvamento
            self.root.attributes('-topmost', False)
            caminho_arquivo = filedialog.asksaveasfilename(
                parent=self.root,
                defaultextension=".xlsx",
                filetypes=[("Arquivos Excel", "*.xlsx")],
                title="Salvar Comparação de Prazos",
                initialfile="Comparacao_Prazos.xlsx"
            )
            self.root.attributes('-topmost', True)

            if not caminho_arquivo:  # Usuário cancelou
                self.log("Operação cancelada pelo usuário.", "aviso")
                self.progresso["value"] = 0
                return

            # Verificar se o diretório existe, criar se necessário
            pasta_destino = os.path.dirname(caminho_arquivo)
            if pasta_destino and not os.path.exists(pasta_destino):
                os.makedirs(pasta_destino)

            self.log("Salvando divergências...", "info")
            self.progresso["value"] = 90
            self.root.update_idletasks()

            # Salvar no local escolhido pelo usuário
            divergencias.to_excel(caminho_arquivo, index=False)

            self.progresso["value"] = 100
            self.root.update_idletasks()

            self.log("Produtos com divergência:", "aviso")
            for index, row in divergencias.iterrows():
                self.log_area.config(state="normal")
                self.log_area.insert(tk.END, f"{row[coluna_chave_erp]}\n", "divergencia")
                self.log_area.config(state="disabled")
                self.log_area.yview(tk.END)
                self.root.update_idletasks()

            total_itens = len(df_comparacao)
            total_divergencias = len(divergencias)
            self.log(f"Total de EAN/SKUS verificados: {total_itens}", "info")
            self.log(f"Total de EAN/SKUS com divergência: {total_divergencias}", "info")

            messagebox.showinfo("Concluído", 
                              f"Comparação de prazos finalizada com sucesso!\nArquivo salvo em:\n{caminho_arquivo}", 
                              parent=self.root)

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