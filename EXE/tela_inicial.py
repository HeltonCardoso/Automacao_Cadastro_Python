import tkinter as tk
from tkinter import messagebox, scrolledtext
from tkinter.ttk import Style
from Extracao_Atributos import TelaExtracaoAtributos
from Cadastro_Produto import TelaCadastroProduto
from Comparacao_Prazos import TelaComparacaoPrazos
from datetime import datetime
import os
from pathlib import Path

class TelaPrincipal:
    def __init__(self, root):
        self.root = root

        self.janelas_abertas ={
            'cadastro'  : None,
            'extracao'  : None,
            'comparacao': None
            }
        self.root.title("CADASTRO MPOZENATO - V1.1")
        self.root.resizable(False, False)
        self.centralizar_janela(500, 350)  # Ajuste a altura para acomodar o relógio

        try:
            # Para Windows
            self.root.iconbitmap("IMG/icone.ico")  # Substitua pelo caminho do seu ícone .ico
        except Exception as e:
            print(f"Erro ao carregar o ícone .ico: {e}")


        # Menu superior
        menu_superior = tk.Menu(root)
        root.config(menu=menu_superior)

        # Menu Arquivo
        menu_arquivo = tk.Menu(menu_superior, tearoff=0)
        menu_superior.add_cascade(label="Arquivo", menu=menu_arquivo)
        menu_arquivo.add_command(label="Sair", command=root.quit)

        # Menu Ajuda
        menu_ajuda = tk.Menu(menu_superior, tearoff=0)
        menu_superior.add_cascade(label="Ajuda", menu=menu_ajuda)
        menu_ajuda.add_command(label="Sobre", command=self.mostrar_sobre)
        menu_ajuda.add_command(label="Documentação", command=self.mostrar_manual_html)

        # Frame principal
        frame_principal = tk.Frame(root)
        frame_principal.pack(pady=20)
        

        # Aplicar estilo moderno aos botões
        self.aplicar_estilo_botoes()

        # Botões
        btn_extrair = tk.Button(frame_principal, text="EXTRAIR ATRIBUTOS", command=self.abrir_extracao_atributos, width=25, bg="#008CBA", fg="white", font=("Arial", 10, "bold"), bd=0, relief=tk.FLAT)
        btn_extrair.pack(pady=10, ipadx=10, ipady=5)
        btn_extrair.bind("<Enter>", lambda e: btn_extrair.config(bg="#0a2040"))  # Efeito hover
        btn_extrair.bind("<Leave>", lambda e: btn_extrair.config(bg="#008CBA"))

        btn_cadastrar = tk.Button(frame_principal, text="PREENCHER PLANILHA ATHUS", command=self.abrir_cadastro_produto, width=25, bg="#008CBA", fg="white", font=("Arial", 10, "bold"), bd=0, relief=tk.FLAT)
        btn_cadastrar.pack(pady=10, ipadx=10, ipady=5)
        btn_cadastrar.bind("<Enter>", lambda e: btn_cadastrar.config(bg="#0a2040"))  # Efeito hover
        btn_cadastrar.bind("<Leave>", lambda e: btn_cadastrar.config(bg="#008CBA"))

        btn_comparar = tk.Button(frame_principal, text="VERIFICAR PRAZOS", command=self.abrir_comparacao_prazos, width=25, bg="#008CBA", fg="white", font=("Arial", 10, "bold"), bd=0, relief=tk.FLAT)
        btn_comparar.pack(pady=10, ipadx=10, ipady=5)
        btn_comparar.bind("<Enter>", lambda e: btn_comparar.config(bg="#0a2040"))  # Efeito hover
        btn_comparar.bind("<Leave>", lambda e: btn_comparar.config(bg="#008CBA"))

        btn_fechar = tk.Button(frame_principal, text="FECHAR", command=root.quit, width=25, bg="#f44336", fg="white", font=("Arial", 10, "bold"), bd=0, relief=tk.FLAT)
        btn_fechar.pack(pady=10, ipadx=10, ipady=5)
        btn_fechar.bind("<Enter>", lambda e: btn_fechar.config(bg="#e53935"))  # Efeito hover
        btn_fechar.bind("<Leave>", lambda e: btn_fechar.config(bg="#f44336"))


        # 4. Atalho F1
        self.root.bind("<F1>", lambda e: self.criar_menu_ajuda())


        # Rodapé
        frame_rodape = tk.Frame(root)
        frame_rodape.pack(side=tk.BOTTOM, fill=tk.X, pady=10)

        # Texto de versão
        texto_versao = tk.Label(frame_rodape, text="Versão 1.1 - Desenvolvido por Helton", fg="gray")
        texto_versao.pack(side=tk.LEFT, padx=10)

        # Relógio no rodapé
        self.relogio = tk.Label(frame_rodape, font=("Arial", 8), fg="gray")
        self.relogio.pack(side=tk.RIGHT, padx=10)
        self.atualizar_relogio()

    def centralizar_janela(self, largura, altura):
        largura_tela = self.root.winfo_screenwidth()
        altura_tela = self.root.winfo_screenheight()
        pos_x = (largura_tela // 2) - (largura // 2)
        pos_y = (altura_tela // 2) - (altura // 2)
        self.root.geometry(f"{largura}x{altura}+{pos_x}+{pos_y}")

    def aplicar_estilo_botoes(self):
        style = Style()
        style.configure("TButton", font=("Arial", 10, "bold"), padding=10, relief=tk.FLAT)
        style.map("TButton",
                  background=[("active", "#45a049")],  # Cor ao clicar
                  foreground=[("active", "blue")])

    def atualizar_relogio(self):
        agora = datetime.now()
        data_hora = agora.strftime("%d/%m/%Y %H:%M:%S")
        self.relogio.config(text=data_hora)
        self.root.after(1000, self.atualizar_relogio)  # Atualiza a cada 1 segundo

    def abrir_extracao_atributos(self):
        if self.janelas_abertas['extracao'] is None or not self.janelas_abertas['extracao'].winfo_exists():
            tela_extracao = tk.Toplevel(self.root)
            self.janelas_abertas['extracao'] = tela_extracao
            TelaExtracaoAtributos(tela_extracao)
            
            # Configura o fechamento para limpar a referência
            tela_extracao.protocol("WM_DELETE_WINDOW", lambda: self.fechar_janela('extracao'))
            
            tela_extracao.iconbitmap("IMG/icone.ico")
            tela_extracao.transient(self.root)
            tela_extracao.grab_set()
        else:
            self.janelas_abertas['extracao'].lift()

    def abrir_comparacao_prazos(self):
        if self.janelas_abertas['comparacao'] is None or not self.janelas_abertas['comparacao'].winfo_exists():
            tela_comparacao = tk.Toplevel(self.root)
            self.janelas_abertas['comparacao'] = tela_comparacao
            TelaComparacaoPrazos(tela_comparacao)
            
            # Configura o fechamento para limpar a referência
            tela_comparacao.protocol("WM_DELETE_WINDOW", lambda: self.fechar_janela('comparacao'))
            
            tela_comparacao.iconbitmap("IMG/icone.ico")
            tela_comparacao.transient(self.root)
            tela_comparacao.grab_set()
        else:
            self.janelas_abertas['comparacao'].lift()

    def abrir_cadastro_produto(self):
        if self.janelas_abertas['cadastro'] is None or not self.janelas_abertas['cadastro'].winfo_exists():
            tela_cadastro = tk.Toplevel(self.root)
            self.janelas_abertas['cadastro'] = tela_cadastro
            TelaCadastroProduto(tela_cadastro)
            
            # Configura o fechamento para limpar a referência
            tela_cadastro.protocol("WM_DELETE_WINDOW", lambda: self.fechar_janela('cadastro'))
            
            tela_cadastro.iconbitmap("IMG/icone.ico")
            tela_cadastro.transient(self.root)
            tela_cadastro.grab_set()
        else:
            self.janelas_abertas['cadastro'].lift()

    def fechar_janela(self, tipo):
        """Fecha a janela e limpa a referência"""
        if self.janelas_abertas[tipo] is not None:
            self.janelas_abertas[tipo].destroy()
        self.janelas_abertas[tipo] = None

    def mostrar_sobre(self):
        messagebox.showinfo("Sobre", "** AUTOMATIZAÇÃO PREENCHIMENTO PLANILHA ATHUS \n \n"
        " ** EXTRAÇÃO DE ATRIBUTOS PLANILHA ONLINE \n \n"
        " ** COMPARAÇÃO DE PRAZOS ONCLICK E MARKETPLACES.")

    def mostrar_leia_me(self):
        caminho = Path(__file__).parent / "DOCUMENTACAO" / "documentacao.txt"
        with open(caminho, 'r', encoding='utf-8') as f:
            conteudo = f.read()
        
        janela = tk.Toplevel(self.root)
        texto = scrolledtext.ScrolledText(janela, wrap=tk.WORD, width=80, height=25)
        texto.pack()
        texto.insert(tk.INSERT, conteudo)
        texto.config(state='disabled')
    
    def mostrar_manual_html(self):
        import webbrowser
        caminho = Path(__file__).parent / "DOCUMENTACAO" / "manual.html"
        webbrowser.open(caminho.as_uri())

    def mostrar_documentacao(self):
        # Cria uma nova janela para exibir a documentação
        janela_documentacao = tk.Toplevel(self.root)
        janela_documentacao.title("DOCUMENTAÇÃO")
        janela_documentacao.geometry("800x600")
        # Adiciona um widget ScrolledText para exibir o conteúdo
        texto_documentacao = scrolledtext.ScrolledText(janela_documentacao, wrap=tk.WORD, width=100, height=30)
        texto_documentacao.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        # Caminho relativo para a pasta DOCUMENTAÇÃO no diretório acima
        caminho_documentacao = Path(__file__).resolve().parent / "DOCUMENTACAO" / "documentacao.txt"
        # Abrir o arquivo de documentação
        with open(caminho_documentacao, "r", encoding="utf-8") as arquivo:
            conteudo = arquivo.read()
            texto_documentacao.insert(tk.INSERT, conteudo)

        # Desabilita a edição do texto
        texto_documentacao.configure(state="disabled")

if __name__ == "__main__":
    root = tk.Tk()
    app = TelaPrincipal(root)
    root.mainloop()