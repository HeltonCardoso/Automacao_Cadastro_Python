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
            self.root.iconbitmap("IMG/icone.ico")  # Substitua "icone.ico" pelo caminho do seu ícone
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

    def extrair_pesos(self, texto_limpo):
        """
        Extrai 'Peso' e 'Peso Suportado' de um texto, incluindo casos com:
        - Peso simples: "Peso: 10 kg"
        - Peso Suportado Distribuído: "30 kg / 20 kg / 15 kg"
        - Formatos variados: "30Kg", "15 kg distribuídos"
        """
        def formatar_peso(valor):
            """Formata o valor para exibir como inteiro ou decimal"""
            if valor.is_integer():
                return f"{int(valor)} kg"
            else:
                return f"{valor:.1f} kg".replace(".", ",")

        pesos = {
            "Peso": "",
            "Peso Suportado": ""
        }

        # --- 1. Extrai PESO normal ---
        padrao_peso = re.compile(r"Peso[:\s]*([\d,\.]+)\s*kg", re.IGNORECASE)
        match_peso = padrao_peso.search(texto_limpo)
        if match_peso:
            valor = float(match_peso.group(1).replace(",", "."))
            pesos["Peso"] = formatar_peso(valor)

        # --- 2. Extrai PESO SUPORTADO (incluindo múltiplos valores) ---
        # Padrão para capturar todo o bloco após "Peso Suportado Distribuído:"
        padrao_bloco = re.compile(
            r"Peso\s*Suportado\s*Distribuído[:\s]*([^/\n]+(?:\/[^/\n]+)*)", 
            re.IGNORECASE
        )
        
        # Padrão para extrair valores individuais (ex: "30 kg", "15Kg distribuídos")
        padrao_valores = re.compile(r"([\d,\.]+)\s*kg", re.IGNORECASE)

        # Encontra todos os blocos de "Peso Suportado Distribuído"
        blocos = padrao_bloco.finditer(texto_limpo)
        valores_encontrados = []

        for bloco in blocos:
            # Extrai valores individuais de cada bloco (separados por "/")
            partes = bloco.group(1).split("/")
            for parte in partes:
                match_valor = padrao_valores.search(parte)
                if match_valor:
                    valor = float(match_valor.group(1).replace(",", "."))
                    valores_encontrados.append(valor)

        # Se encontrou valores, pega o maior
        if valores_encontrados:
            maior_valor = max(valores_encontrados)
            pesos["Peso Suportado"] = formatar_peso(maior_valor)
        else:
            # Fallback: procura por "Peso Suportado" simples (não distribuído)
            padrao_simples = re.compile(
                r"(?:Peso\s*Suportado|Suporta|Carga\s*Máxima)[:\s]*([\d,\.]+)\s*kg", 
                re.IGNORECASE
            )
            match_simples = padrao_simples.search(texto_limpo)
            if match_simples:
                valor = float(match_simples.group(1).replace(",", "."))
                pesos["Peso Suportado"] = formatar_peso(valor)

        return pesos

    def extrair_atributos(self, descricao_html):
        """Extrai atributos de uma descrição HTML."""
        atributos = {
            "Largura": "", "Altura": "", "Profundidade": "", "Peso": "", "Cor": "", 
            "Modelo": "", "Fabricante": "", "Volumes": "", "Material da Estrutura": "", 
            "Material": "", "Peso Suportado": "", "Acabamento": "", "Possui Portas": "",
            "Quantidade de Portas": "", "Tipo de Porta": "", "Possui Prateleiras": "", 
            "Quantidade de Prateleiras": "", "Conteúdo da Embalagem": "", 
            "Quantidade de Gavetas": "", "Possui Gavetas": "", "Revestimento": "", 
            "Quantidade de lugares": "", "Possui Nicho": "", "Quantidade de Assentos": "",
            "Tipo de Assento": "", "Sugestão de Lugares": "", "Tipo de Encosto": ""
        }

        if pd.notna(descricao_html):
            soup = BeautifulSoup(descricao_html, "html.parser")
            texto_limpo = soup.get_text()
        else:
            texto_limpo = ""

        # Função auxiliar para formatar medidas
        def formatar_medida(valor):
            if valor.is_integer():
                return f"{int(valor)} cm"
            else:
                return f"{valor:.1f} cm".replace(".", ",")

        # Regex para medidas no formato "L x A x P" (definido ANTES do uso)
        regex_medidas = re.compile(
            r"\b(\d+[,\.]?\d*)\s*(?:cm\s*)?x\s*(\d+[,\.]?\d*)\s*(?:cm\s*)?x\s*(\d+[,\.]?\d*)\s*cm\b",
            re.IGNORECASE
        )

        # --- 1. Extrai medidas explícitas ---
        for medida in ["Largura", "Altura", "Profundidade"]:
            padrao = rf"{medida}[:\s]*([\d,\.]+)\s*cm"
            match = re.search(padrao, texto_limpo, re.IGNORECASE)
            if match:
                valor = float(match.group(1).replace(",", "."))
                atributos[medida] = formatar_medida(valor)

        # --- 2. Fallback: medidas no formato "L x A x P" ---
        if not any(atributos.get(medida) for medida in ["Largura", "Altura", "Profundidade"]):
            matches_medidas = regex_medidas.findall(texto_limpo)
            if matches_medidas:
                # Pega os maiores valores de cada dimensão
                larguras = [float(m[0].replace(",", ".")) for m in matches_medidas]
                alturas = [float(m[1].replace(",", ".")) for m in matches_medidas]
                profundidades = [float(m[2].replace(",", ".")) for m in matches_medidas]
                
                atributos["Largura"] = formatar_medida(max(larguras))
                atributos["Altura"] = formatar_medida(max(alturas))
                atributos["Profundidade"] = formatar_medida(max(profundidades))

        # --- 3. Extrai PESOS (usando o método separado) ---
        pesos = self.extrair_pesos(texto_limpo)
        atributos.update(pesos)

        # --- 4. Extrai OUTROS ATRIBUTOS (Material, Cor, Volumes, etc.) ---
        regex_padroes = {
            "Cor": r"Cor[:\s]*([\w\s]+)",
            "Modelo": r"Modelo[:\s]*([\w\s]+)",
            "Fabricante": r"Fabricante[:\s]*([\w\s]+)",
            "Volumes": r"Volumes[:\s]*(\d+)",
            "Material da Estrutura": r"Material da Estrutura[:\s]*([\w\s]+)",
            "Possui Portas": r"Possui Portas[:\s]*(Sim|Não)",
            "Quantidade de Portas": r"Quantidade de Portas[:\s]*(\d+)",
            "Tipo de Porta": r"Tipo de Porta[:\s]*([\w\s]+)",
            "Possui Prateleiras": r"Possui Prateleiras[:\s]*(Sim|Não)",
            "Quantidade de Prateleiras": r"Quantidade de Prateleiras[:\s]*(\d+)",
            "Conteúdo da Embalagem": r"Conteúdo da Embalagem[:\s]*([\w\s,]+)",
            "Quantidade de Gavetas": r"Quantidade de Gavetas[:\s]*(\d+)",
            "Possui Gavetas": r"Possui Gavetas[:\s]*(Sim|Não)",
            "Quantidade de lugares": r"Quantidade de lugares[:\s]*(\d+)",
            "Sugestão de Lugares": r"Sugestão de Lugares[:\s]*(\d+)",
            "Quantidade de Assentos": r"Quantidade de Assentos[:\s]*(\d+)",
            "Tipo de Assento": r"Tipo de Assento[:\s]*([\w\s,]+)",
            "Possui Nicho": r"Possui Nicho[:\s]*(Sim|Não)",
            "Tipo de Encosto": r"Tipo de Encosto[:\s]*([\w\s,]+)"
        }

        for atributo, padrao in regex_padroes.items():
            match = re.search(padrao, texto_limpo, re.IGNORECASE)
            if match:
                atributos[atributo] = match.group(1).strip()

        # --- 5. Extrai ATRIBUTOS ESPECÍFICOS (Material, Acabamento, Revestimento) ---
        secao_caracteristicas = re.search(
            r"Características do Produto[:\-]?\s*([\s\S]+?)(?:\n\n|\Z)", 
            texto_limpo, 
            re.IGNORECASE
        )
        texto_caracteristicas = secao_caracteristicas.group(1) if secao_caracteristicas else texto_limpo

        # Material (prioridade na seção de características)
        match_material = re.search(r"Material[:\s]*([\w\s]+)", texto_caracteristicas, re.IGNORECASE)
        if match_material:
            atributos["Material"] = match_material.group(1).strip()

        # Acabamento e Revestimento
        for atributo, padrao in {
            "Acabamento": r"Acabamento[:\-]?\s*([\w\s\-,]+)",
            "Revestimento": r"Revestimento[:\s]*([\w\s,]+)"
        }.items():
            match = re.search(padrao, texto_caracteristicas, re.IGNORECASE)
            if match:
                atributos[atributo] = match.group(1).strip()

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
                    "Quantidade de Gavetas", "Possui Gavetas", "Revestimento", "Quantidade de lugares","Possui Nicho",
                    "Quantidade de Assentos","Tipo de Assento","Sugestão de Lugares","Tipo de Encosto"
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

                # Abrir diálogo para escolher local de salvamento
                self.root.attributes('-topmost', False)
                arquivo_saida = filedialog.asksaveasfilename(
                    parent=self.root,
                    defaultextension=".xlsx",
                    filetypes=[("Arquivos Excel", "*.xlsx")],
                    title="Salvar Atributos Extraídos",
                    initialfile="Atributos_Extraidos.xlsx"
                )
                self.root.attributes('-topmost', True)

                if not arquivo_saida:  # Usuário cancelou
                    self.log("Operação cancelada pelo usuário.", "aviso")
                    self.progresso["value"] = 0
                    return

                # Verificar se o diretório existe, criar se necessário
                pasta_destino = os.path.dirname(arquivo_saida)
                if pasta_destino and not os.path.exists(pasta_destino):
                    os.makedirs(pasta_destino)

                try:
                    df_saida.to_excel(arquivo_saida, index=False)
                    self.log("Arquivo salvo com sucesso!")
                    messagebox.showinfo("Sucesso", 
                                    f"Atributos extraídos com sucesso!\nSalvo em:\n{arquivo_saida}", 
                                    parent=self.root)
                except PermissionError:
                    messagebox.showerror("Erro", "Feche o arquivo de saída e tente novamente!", parent=self.root)
                    return
                except Exception as e:
                    messagebox.showerror("Erro", f"Erro ao salvar o arquivo: {str(e)}", parent=self.root)
                    return

# Função principal
if __name__ == "__main__":
    root = tk.Tk()
    app = TelaExtracaoAtributos(root)
    root.mainloop()