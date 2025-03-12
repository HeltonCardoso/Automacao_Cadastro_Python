import tkinter as tk
from tkinter import messagebox, scrolledtext
from tkinter.ttk import Style
from Extracao_Atributos import TelaExtracaoAtributos
from Cadastro_Produto import TelaCadastroProduto
from Comparacao_Prazos import TelaComparacaoPrazos
from datetime import datetime

class TelaPrincipal:
    def __init__(self, root):
        self.root = root
        self.root.title("CADASTRO DE PRODUTOS - V1.0")
        self.centralizar_janela(500, 350)  # Ajuste a altura para acomodar o relógio

        # Adicionar ícone à janela
        try:
            self.root.iconbitmap("icone.ico")  # Substitua "icone.ico" pelo caminho do seu ícone
            # Forçar o ícone na barra de tarefas (Windows)
            self.root.tk.call('wm', 'iconphoto', self.root._w, tk.PhotoImage(file="icone.png"))
        except Exception as e:
            print(f"Erro ao carregar o ícone: {e}")  # Mensagem de erro se o ícone não for encontrado

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
        menu_ajuda.add_command(label="Documentação", command=self.mostrar_documentacao)

        # Frame principal
        frame_principal = tk.Frame(root)
        frame_principal.pack(pady=20)
        

        # Aplicar estilo moderno aos botões
        self.aplicar_estilo_botoes()

        # Botões
        btn_extrair = tk.Button(frame_principal, text="EXTRAIR ATRIBUTOS", command=self.abrir_extracao_atributos, width=25, bg="#008CBA", fg="white", font=("Arial", 10, "bold"), bd=0, relief=tk.FLAT)
        btn_extrair.pack(pady=10, ipadx=10, ipady=5)
        btn_extrair.bind("<Enter>", lambda e: btn_extrair.config(bg="#45a049"))  # Efeito hover
        btn_extrair.bind("<Leave>", lambda e: btn_extrair.config(bg="#008CBA"))

        btn_cadastrar = tk.Button(frame_principal, text="PREENCHER PLANILHA ATHUS", command=self.abrir_cadastro_produto, width=25, bg="#008CBA", fg="white", font=("Arial", 10, "bold"), bd=0, relief=tk.FLAT)
        btn_cadastrar.pack(pady=10, ipadx=10, ipady=5)
        btn_cadastrar.bind("<Enter>", lambda e: btn_cadastrar.config(bg="#45a049"))  # Efeito hover
        btn_cadastrar.bind("<Leave>", lambda e: btn_cadastrar.config(bg="#008CBA"))

        btn_comparar = tk.Button(frame_principal, text="VERIFICAR PRAZOS", command=self.abrir_comparacao_prazos, width=25, bg="#008CBA", fg="white", font=("Arial", 10, "bold"), bd=0, relief=tk.FLAT)
        btn_comparar.pack(pady=10, ipadx=10, ipady=5)
        btn_comparar.bind("<Enter>", lambda e: btn_comparar.config(bg="#45a049"))  # Efeito hover
        btn_comparar.bind("<Leave>", lambda e: btn_comparar.config(bg="#008CBA"))

        btn_fechar = tk.Button(frame_principal, text="FECHAR", command=root.quit, width=25, bg="#f44336", fg="white", font=("Arial", 10, "bold"), bd=0, relief=tk.FLAT)
        btn_fechar.pack(pady=10, ipadx=10, ipady=5)
        btn_fechar.bind("<Enter>", lambda e: btn_fechar.config(bg="#e53935"))  # Efeito hover
        btn_fechar.bind("<Leave>", lambda e: btn_fechar.config(bg="#f44336"))

        # Rodapé
        frame_rodape = tk.Frame(root)
        frame_rodape.pack(side=tk.BOTTOM, fill=tk.X, pady=10)

        # Texto de versão
        texto_versao = tk.Label(frame_rodape, text="Versão 1.0 - Desenvolvido por Helton", fg="gray")
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
        tela_extracao = tk.Toplevel(self.root)
        TelaExtracaoAtributos(tela_extracao)
        tela_extracao.iconbitmap("icone.ico")  # Define o ícone para a janela filha
        tela_extracao.lift()
        tela_extracao.attributes('-topmost', True)
        tela_extracao.focus_force()

    def abrir_cadastro_produto(self):
        tela_cadastro = tk.Toplevel(self.root)
        TelaCadastroProduto(tela_cadastro)
        tela_cadastro.iconbitmap("icone.ico")  # Define o ícone para a janela filha
        tela_cadastro.lift()
        tela_cadastro.attributes('-topmost', True)
        tela_cadastro.focus_force()

    def abrir_comparacao_prazos(self):
        tela_comparacao = tk.Toplevel(self.root)
        TelaComparacaoPrazos(tela_comparacao)
        tela_comparacao.iconbitmap("icone.ico")  # Define o ícone para a janela filha
        tela_comparacao.lift()
        tela_comparacao.attributes('-topmost', True)
        tela_comparacao.focus_force()

    def mostrar_sobre(self):
        messagebox.showinfo("Sobre", "Sistema para automatizar o processo de Preenchimento da Planilha Athus e Extração dos principais atributos dos produtos.")

    def mostrar_documentacao(self):
        # Cria uma nova janela para exibir a documentação
        janela_documentacao = tk.Toplevel(self.root)
        janela_documentacao.title("Documentação do Sistema")
        janela_documentacao.geometry("800x600")

        # Adiciona um widget ScrolledText para exibir o conteúdo
        texto_documentacao = scrolledtext.ScrolledText(janela_documentacao, wrap=tk.WORD, width=100, height=30)
        texto_documentacao.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        # Insere o conteúdo da documentação
        with open("documentacao.txt", "r", encoding="utf-8") as arquivo:
            conteudo = arquivo.read()
            texto_documentacao.insert(tk.INSERT, conteudo)

        # Desabilita a edição do texto
        texto_documentacao.configure(state="disabled")

if __name__ == "__main__":
    root = tk.Tk()
    app = TelaPrincipal(root)
    root.mainloop()