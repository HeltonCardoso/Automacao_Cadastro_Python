import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from tkinter.ttk import Progressbar
import pandas as pd
import os
from bs4 import BeautifulSoup
import re
from datetime import datetime

class TelaExtracaoAtributos:
    def __init__(self, root):
        self.root = root
        self.root.title("EXTRAÇÃO ATRIBUTOS - V1.0")
        self.root.resizable(False, False)
        self.centralizar_janela(800, 510)
        self.aplicar_estilo_global()  # Aplica o estilo global

        # Adicionar ícone à janela
        try:
            self.root.iconbitmap("icone.ico")  # Substitua "icone.ico" pelo caminho do seu ícone
        except:
            pass  # Ignora se o ícone não for encontrado

        # Frame principal
        frame = tk.Frame(self.root)
        frame.pack(pady=20)

        # Label e Entry para selecionar a planilha
        ttk.Label(frame, text="Selecione a Planilha:").grid(row=0, column=0)
        self.entrada_arquivo = tk.Entry(frame, width=60, font=("Arial", 10))
        self.entrada_arquivo.grid(row=0, column=1)
        btn_buscar = tk.Button(
            frame, 
            text="SELECIONAR", 
            command=self.selecionar_arquivo, 
            bg="#008CBA", 
            fg="white", 
            font=("Arial", 10, "bold"), 
            bd=4, 
            relief=tk.FLAT
        )
        btn_buscar.grid(row=0, column=2)
        btn_buscar.bind("<Enter>", lambda e: btn_buscar.config(bg="#0a2040"))  # Efeito hover
        btn_buscar.bind("<Leave>", lambda e: btn_buscar.config(bg="#008CBA"))  # Efeito ao sair

        # Botão "Extrair Atributos"
        btn_extrair = tk.Button(
            self.root, 
            text="EXTRAIR ATRIBUTOS", 
            command=self.extrair_dados, 
            bg="#008CBA", 
            fg="white", 
            font=("Arial", 10, "bold"), 
            bd=4, 
            relief=tk.FLAT
        )
        btn_extrair.pack(pady=10)
        btn_extrair.bind("<Enter>", lambda e: btn_extrair.config(bg="#0a2040"))  # Efeito hover
        btn_extrair.bind("<Leave>", lambda e: btn_extrair.config(bg="#008CBA"))  # Efeito ao sair

        # Barra de progresso
        self.progresso = Progressbar(self.root, orient="horizontal", length=735, mode="determinate")
        self.progresso.pack(pady=5)

        # Label de status
        self.status_label = tk.Label(self.root, text="Pronto para iniciar...", fg="gray", font=("Arial", 10))
        self.status_label.pack(pady=5)

        # Área de logs
        self.log_area = scrolledtext.ScrolledText(
            self.root, 
            width=103, 
            height=15, 
            bg="white", 
            state="disabled", 
            font=("Courier", 8)
        )
        self.log_area.pack(pady=10)

        # Configuração de tags para cores
        self.log_area.tag_config("info", foreground="gray", font=("Courier", 11, "bold"))
        self.log_area.tag_config("erro", foreground="red", font=("Courier", 11, "bold"))
        self.log_area.tag_config("sucesso", foreground="green", font=("Courier", 11, "bold"))
        self.log_area.tag_config("aviso", foreground="orange", font=("Courier", 11, "bold"))

        # Botão "Fechar"
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
        btn_fechar.bind("<Enter>", lambda e: btn_fechar.config(bg="#e53935"))  # Efeito hover
        btn_fechar.bind("<Leave>", lambda e: btn_fechar.config(bg="#f44336"))  # Efeito ao sair

        # Rodapé
        frame_rodape = tk.Frame(root)
        frame_rodape.pack(side=tk.BOTTOM, fill=tk.X, pady=10)

        # Texto de versão
        texto_versao = tk.Label(frame_rodape, text="Versão 1.0 - Desenvolvido por Helton", fg="gray", font=("Arial", 8))
        texto_versao.pack(side=tk.LEFT, padx=10)

        # Relógio no rodapé
        self.relogio = tk.Label(frame_rodape, font=("Arial", 8), fg="gray")
        self.relogio.pack(side=tk.RIGHT, padx=10)
        self.atualizar_relogio()

    def aplicar_estilo_global(self):
        """Aplica o estilo global da interface."""
        style = ttk.Style(self.root)
        style.configure("TButton", font=("Arial", 10, "bold"), padding=10, relief="flat", background="#008CBA", foreground="white")
        style.map("TButton", background=[("active", "#0a2040")], foreground=[("active", "white")])
        style.configure("Fechar.TButton", background="#f44336", foreground="white")
        style.map("Fechar.TButton", background=[("active", "#e53935")], foreground=[("active", "white")])
        style.configure("TLabel", font=("Arial", 10), foreground="gray")
        style.configure("TEntry", font=("Arial", 10), relief="flat")

    def centralizar_janela(self, largura, altura):
        """Centraliza a janela na tela."""
        largura_tela = self.root.winfo_screenwidth()
        altura_tela = self.root.winfo_screenheight()
        pos_x = (largura_tela // 2) - (largura // 2)
        pos_y = (altura_tela // 2) - (altura // 2)
        self.root.geometry(f"{largura}x{altura}+{pos_x}+{pos_y}")

    def atualizar_relogio(self):
        """Atualiza o relógio no rodapé."""
        agora = datetime.now()
        data_hora = agora.strftime("%d/%m/%Y %H:%M:%S")
        self.relogio.config(text=data_hora)
        self.root.after(1000, self.atualizar_relogio)  # Atualiza a cada 1 segundo

    def selecionar_arquivo(self):
        """Abre o diálogo para selecionar um arquivo."""
        self.root.grab_release()
        caminho = filedialog.askopenfilename(
            parent=self.root,
            filetypes=[("Planilhas Excel", "*.xlsx")]
        )
        self.root.grab_set()
        self.root.lift()
        self.root.attributes('-topmost', True)
        self.root.focus_force()

        if caminho:
            self.entrada_arquivo.delete(0, tk.END)
            self.entrada_arquivo.insert(0, caminho)

    def log(self, mensagem, tipo="info"):
        """Adiciona uma mensagem à área de logs com timestamp."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        mensagem_formatada = f"[{timestamp}] {mensagem}\n"

        self.log_area.config(state="normal")
        self.log_area.insert(tk.END, mensagem_formatada, tipo)
        self.log_area.config(state="disabled")
        self.log_area.yview(tk.END)  # Rola para a última linha

    def extrair_atributos(self, descricao_html):
        """Extrai atributos de uma descrição HTML."""
        atributos = {
            "Largura": "", "Altura": "", "Profundidade": "", "Peso": "", "Cor": "", "Modelo": "", "Fabricante": "",
            "Volumes": "", "Material da Estrutura": "", "Material": "", "Peso Suportado": "", "Acabamento": "", "Possui Portas": "",
            "Quantidade de Portas": "", "Tipo de Porta": "", "Possui Prateleiras": "", "Quantidade de Prateleiras": "",
            "Conteúdo da Embalagem": "", "Quantidade de Gavetas": "", "Possui Gavetas": "", "Revestimento": "", "Quantidade de lugares": "","Possui Nicho": ""
        }

        if pd.notna(descricao_html):
            soup = BeautifulSoup(descricao_html, "html.parser")
            texto_limpo = soup.get_text()
        else:
            texto_limpo = ""

            # Regex para "Peso"
        regex_peso = re.compile(r"Peso[:\s]*(\d+[,\.]?\d*)\s*kg", re.IGNORECASE)
        match_peso = regex_peso.search(texto_limpo)
        if match_peso:
            # Extrai o número e formata como "X kg"
            peso = match_peso.group(1).replace(',', '.')
            atributos["Peso"] = f"{peso} kg"

        # Regex para "Peso Suportado" (flexível para variações de texto)
        regex_peso_suportado = re.compile(
            r"Peso\s*suportado\s*(?:distribuído)?[:\s]*(\d+[,\.]?\d*)\s*kg", 
            re.IGNORECASE
        )
        match_peso_suportado = regex_peso_suportado.search(texto_limpo)
        if match_peso_suportado:
            # Extrai o número e formata como "X kg"
            peso_suportado = match_peso_suportado.group(1).replace(',', '.')
            atributos["Peso Suportado"] = f"{peso_suportado} kg"

        # Compila as expressões regulares uma única vez
        regex_padroes = {
            "Largura": re.compile(r"Largura[:\s]*([\d,\.]+)\s*cm?", re.IGNORECASE),
            "Altura": re.compile(r"Altura[:\s]*([\d,\.]+)\s*cm?", re.IGNORECASE),
            "Profundidade": re.compile(r"Profundidade[:\s]*([\d,\.]+)\s*cm?", re.IGNORECASE),
            "Peso": re.compile(r"Peso[:\s]*([\d,\.]+)\s*kg?", re.IGNORECASE),
            "Volumes": re.compile(r"Volumes[:\s]*(\d+)", re.IGNORECASE),
            "Material da Estrutura": re.compile(r"Material da Estrutura[:\s]*([\w\s]+)", re.IGNORECASE),
            "Possui Portas": re.compile(r"Possui Portas[:\s]*(Sim|Não)", re.IGNORECASE),
            "Quantidade de Portas": re.compile(r"Quantidade de Portas[:\s]*(\d+)", re.IGNORECASE),
            "Tipo de Porta": re.compile(r"Tipo de Porta[:\s]*([\w\s]+)", re.IGNORECASE),
            "Possui Prateleiras": re.compile(r"Possui Prateleiras[:\s]*(Sim|Não)", re.IGNORECASE),
            "Quantidade de Prateleiras": re.compile(r"Quantidade de Prateleiras[:\s]*(\d+)", re.IGNORECASE),
            "Conteúdo da Embalagem": re.compile(r"Conteúdo da Embalagem[:\s]*([\w\s,]+)", re.IGNORECASE),
            "Quantidade de Gavetas": re.compile(r"Quantidade de Gavetas[:\s]*(\d+)", re.IGNORECASE),
            "Possui Gavetas": re.compile(r"Possui Gavetas[:\s]*(Sim|Não)", re.IGNORECASE),
            "Revestimento": re.compile(r"Revestimento[:\s]*([\w\s,]+)", re.IGNORECASE),
            "Quantidade de lugares": re.compile(r"Quantidade de lugares[:\s]*(\d+)", re.IGNORECASE),
            "Possui Nicho": re.compile(r"Possui Nicho[:\s]*(Sim|Não)", re.IGNORECASE),
        }

        # Captura medidas no formato (L x A x P)
        regex_medidas = re.compile(r"\b(\d+[,\.]?\d*)\s*(?:cm)?\s*x\s*(\d+[,\.]?\d*)\s*(?:cm)?\s*x\s*(\d+[,\.]?\d*)\s*(?:cm)?\b", re.IGNORECASE)

        # Busca padrões normais
        for chave, padrao in regex_padroes.items():
            match = padrao.search(texto_limpo)
            if match:
                valor = match.group(1).strip(" .")
                if chave in ["Largura", "Altura", "Profundidade"]:
                    atributos[chave] = valor + " cm"
                elif chave in ["Peso"]:
                    atributos[chave] = valor + " kg"
                else:
                    atributos[chave] = valor

        # Captura medidas gerais (L x A x P)
        matches_medidas = regex_medidas.findall(texto_limpo)

        larguras = []
        alturas = []
        profundidades = []

        for match in matches_medidas:
            largura = float(match[0].replace(",", "."))
            altura = float(match[1].replace(",", "."))
            profundidade = float(match[2].replace(",", "."))

            larguras.append(largura)
            alturas.append(altura)
            profundidades.append(profundidade)

        if larguras:
            atributos["Largura"] = f"{max(larguras):.1f} cm"
        if alturas:
            atributos["Altura"] = f"{max(alturas):.1f} cm"
        if profundidades:
            atributos["Profundidade"] = f"{max(profundidades):.1f} cm"

        # Captura "Peso Suportado" em diferentes formatos
       # regex_peso_suportado = re.compile(r"(?:Peso Suportado(?: Distribuído)?[:\s]*)?(\d+[,\.]?\d*)\s*kg", re.IGNORECASE)

      #  pesos_encontrados = regex_peso_suportado.findall(texto_limpo)

        #if pesos_encontrados:
          #  maior_peso = max([float(p.replace(",", ".")) for p in pesos_encontrados])
          #  atributos["Peso Suportado"] = f"{maior_peso:.1f} kg"

        # Busca a seção "Características do Produto"
        regex_caracteristicas = re.compile(r"(Características do Produto[:\-]?\s*)([\s\S]+?)(?:\n\n|\Z)", re.IGNORECASE)
        match_caracteristicas = regex_caracteristicas.search(texto_limpo)

        if match_caracteristicas:
            texto_caracteristicas = match_caracteristicas.group(2)
        else:
            texto_caracteristicas = texto_limpo  # Se não encontrar, usa todo o texto

        # Captura "Material" apenas na seção "Características do Produto"
        regex_material = re.compile(r"Material[:\s]*([\w\s]+)", re.IGNORECASE)
        match_material = regex_material.search(texto_caracteristicas)

        if match_material:
            atributos["Material"] = match_material.group(1).strip()

        # Captura "Acabamento" dentro das características
        regex_acabamento = re.compile(r"Acabamento[:\-]?\s*([\w\s\-,]+)", re.IGNORECASE)
        match_acabamento = regex_acabamento.search(texto_caracteristicas)

        if match_acabamento:
            atributos["Acabamento"] = match_acabamento.group(1).strip()

        return atributos

    def extrair_dados(self):
        """Extrai os dados da planilha selecionada."""
        caminho = self.entrada_arquivo.get()
        if not caminho:
            messagebox.showerror("Erro", "Selecione um arquivo primeiro!", parent=self.root)
            return

        if not os.path.exists(caminho):
            messagebox.showerror("Erro", "O arquivo selecionado não existe!", parent=self.root)
            return

        try:
            df = pd.read_excel(caminho)
        except PermissionError:
            messagebox.showerror("Erro", "Feche o arquivo e tente novamente!", parent=self.root)
            return
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ler o arquivo: {str(e)}", parent=self.root)
            return

        # Verifica se as colunas necessárias estão presentes
        colunas_necessarias = ["EAN", "NOMEE-COMMERCE", "DESCRICAOHTML", "MODMPZ", "COR"]
        for coluna in colunas_necessarias:
            if coluna not in df.columns:
                messagebox.showerror("Erro", f"A coluna '{coluna}' não foi encontrada no arquivo!", parent=self.root)
                return

        dados_extraidos = []
        colunas = [
            "EAN", "Nome", "Largura", "Altura", "Profundidade", "Peso", "Cor", "Modelo", "Fabricante", "Volumes",
            "Material da Estrutura", "Material", "Peso Suportado", "Acabamento", "Possui Portas", "Quantidade de Portas",
            "Tipo de Porta", "Possui Prateleiras", "Quantidade de Prateleiras", "Conteúdo da Embalagem",
            "Quantidade de Gavetas", "Possui Gavetas", "Revestimento", "Quantidade de lugares", "Possui Nicho"
        ]

        total_linhas = len(df)
        self.progresso["maximum"] = total_linhas
        self.progresso["value"] = 0

        for idx, row in df.iterrows():
            try:
                ean = str(row.get("EAN", "")).strip()
                nome = row.get("NOMEE-COMMERCE", "Desconhecido")
                descricao_html = row.get("DESCRICAOHTML", "")
                modelo = str(row.get("MODMPZ", "")).strip()
                cor = str(row.get("COR", "")).strip()
                fabricante = nome.split("-")[-1].strip() if "-" in nome else ""

                atributos = self.extrair_atributos(descricao_html)
                atributos["Cor"] = cor
                atributos["Modelo"] = modelo
                atributos["Fabricante"] = fabricante

                dados_extraidos.append([ean, nome] + list(atributos.values()))
                self.progresso["value"] = idx + 1
                self.status_label.config(text=f"Processando linha {idx + 1} de {total_linhas}...")
                self.log(f"- {nome}")
                self.root.update_idletasks()
            except Exception as e:
                self.log(f"Erro ao processar linha {idx + 1}: {str(e)}")

        df_saida = pd.DataFrame(dados_extraidos, columns=colunas)

        # Caminho relativo para a pasta ATRIBUTOS na raiz do projeto
        pasta_destino = os.path.join(os.getcwd(), "ATRIBUTOS")

        # Verificar se o diretório existe, caso contrário, criar
        if not os.path.exists(pasta_destino):
            os.makedirs(pasta_destino)

        arquivo_saida = os.path.join(pasta_destino, "Atributos_Extraidos.xlsx")

        try:
            df_saida.to_excel(arquivo_saida, index=False)
            self.log("Arquivo salvo com sucesso!")
            messagebox.showinfo("Sucesso", f"Atributos extraídos com sucesso!\nSalvo em:\n{arquivo_saida}", parent=self.root)
        except PermissionError:
            messagebox.showerror("Erro", "Feche o arquivo de saída e tente novamente!", parent=self.root)
            return

# Função principal
if __name__ == "__main__":
    root = tk.Tk()
    app = TelaExtracaoAtributos(root)
    root.mainloop()