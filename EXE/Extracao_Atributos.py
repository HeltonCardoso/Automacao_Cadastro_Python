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
        self.centralizar_janela(800, 450)


        # Adicionar ícone à janela
        try:
            self.root.iconbitmap("icone.ico")  # Substitua "icone.ico" pelo caminho do seu ícone
        except:
            pass  # Ignora se o ícone não for encontrado

        # Configuração de estilo
        self.style = ttk.Style()
        self.style.configure("TButton", padding=6, font=("Helvetica", 10))
        self.style.configure("TLabel", font=("Helvetica", 10))
        self.style.configure("TEntry", font=("Helvetica", 10))
        
        # Interface
        frame = tk.Frame(self.root)
        frame.pack(pady=20)

        ttk.Label(frame, text="Selecione a Planilha:").grid(row=0, column=0)
        self.entrada_arquivo = tk.Entry(frame, width=80)
        self.entrada_arquivo.grid(row=0, column=1)
        ttk.Button(frame, text="Buscar", command=self.selecionar_arquivo).grid(row=0, column=2)

        ttk.Button(self.root, text="Extrair Atributos", command=self.extrair_dados).pack(pady=10)

        # Barra de progresso
        self.progresso = Progressbar(self.root, orient="horizontal", length=720, mode="determinate")
        self.progresso.pack(pady=5)

        # Label de status
        self.status_label = tk.Label(self.root, text="Pronto para iniciar...", fg="blue")
        self.status_label.pack(pady=5)

        # Área de logs
        self.log_area = scrolledtext.ScrolledText(self.root, width=90, height=10, state="disabled")
        self.log_area.pack(pady=10)

        ttk.Button(self.root, text="FECHAR", command=self.root.destroy).pack(pady=10)


        # Rodapé
        frame_rodape = tk.Frame(root)
        frame_rodape.pack(side=tk.BOTTOM, fill=tk.X, pady=10)

        # Relógio no rodapé
        self.relogio = tk.Label(frame_rodape, font=("Arial", 10), fg="blue")
        self.relogio.pack(side=tk.RIGHT, padx=10)
        self.atualizar_relogio()

       
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

    def selecionar_arquivo(self):
        self.root.grab_release()
        # Abre o filedialog, associando-o à tela secundária
        caminho = filedialog.askopenfilename(
            parent=self.root,  # Associa o filedialog à tela secundária
            filetypes=[("Planilhas Excel", "*.xlsx")]
        )

        # Reativa o grab_set após fechar o filedialog
        self.root.grab_set()

        # Garante que a tela secundária fique no topo
        self.root.lift()
        self.root.attributes('-topmost', True)
        self.root.focus_force()

        if caminho:
            self.entrada_arquivo.delete(0, tk.END)
            self.entrada_arquivo.insert(0, caminho)

    def log(self, mensagem):
        """Adiciona uma mensagem à área de logs com timestamp."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        mensagem_formatada = f"[{timestamp}] {mensagem}\n"

        self.log_area.config(state="normal")
        self.log_area.insert(tk.END, mensagem_formatada)
        self.log_area.config(state="disabled")
        self.log_area.yview(tk.END)  # Rola para a última linha

    def extrair_atributos(self, descricao_html):
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
        regex_peso_suportado = re.compile(r"(?:Peso Suportado(?: Distribuído)?[:\s]*)?(\d+[,\.]?\d*)\s*kg", re.IGNORECASE)

        pesos_encontrados = regex_peso_suportado.findall(texto_limpo)

        if pesos_encontrados:
            maior_peso = max([float(p.replace(",", ".")) for p in pesos_encontrados])
            atributos["Peso Suportado"] = f"{maior_peso:.1f} kg"

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