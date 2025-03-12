
# Automação para Cadastro de Produtos

## Descrição

Este projeto é uma **automação personalizada** desenvolvida para suprir as necessidades de cadastro de produtos de uma empresa, aplicando regras de negócios específicas para o processo. A automação integra planilhas Excel e permite uma interação eficiente com os dados, facilitando tarefas como:

- Busca de atributos para marketplaces
- Cadastro de produtos em sistemas externos (Athus)
- Comparação de prazos de entrega entre o ERP e marketplaces

---

## Funcionalidades

- **Buscar Atributos**: Realiza uma varredura na planilha Excel de cadastro, no campo de descrição HTML, buscando os principais atributos do produto que serão usados nos marketplaces.
- **Cadastrar Produto**: Busca as informações dos produtos de uma planilha existente e preenche uma nova planilha específica para o sistema Athus, com todos os dados e descrições do produto.
- **Comparar Prazos**: Compara os prazos de entrega entre o ERP e os marketplaces, corrigindo discrepâncias de prazo nas planilhas fornecidas por ambos.

---

## Tecnologias Utilizadas

Este projeto foi desenvolvido utilizando as seguintes tecnologias:

- **Python**: Linguagem de programação principal para desenvolvimento da automação.
- **Pandas**: Biblioteca para manipulação e análise de dados, usada para manipulação das planilhas Excel.
- **Tkinter**: Biblioteca gráfica para criação da interface do usuário (GUI).
- **openpyxl**: Biblioteca para ler e escrever arquivos Excel (XLSX).
- **Matplotlib** (opcional): Para gerar gráficos comparativos de prazos de entrega, se necessário.

---

## Como Rodar o Projeto

### Pré-requisitos

- **Python 3.x** instalado.
- **pip** (gerenciador de pacotes Python).
- **Excel** (ou equivalente) para leitura e criação das planilhas.

### 1. Clonando o Repositório

Clone o repositório para sua máquina local:

```bash
git clone https://github.com/HeltonCardoso/Automacao_Cadastro_Python.git
cd Automacao_Cadastro_Python
```

### 2. Instalando Dependências

Instale as dependências do projeto (caso haja um arquivo `requirements.txt`):

```bash
pip install -r requirements.txt
```

Se não houver o `requirements.txt`, instale as dependências necessárias manualmente:

```bash
pip install pandas
pip install openpyxl
pip install tkinter
```

### 3. Executando o Projeto

Para rodar o projeto, execute o arquivo principal:

```bash
python tela_inicial.py
```

---

## Estrutura do Projeto

A estrutura do repositório é a seguinte:

```
Automacao_Cadastro_Python/
│
├── EX                          # Pasta dos arquivo.py 
│   ├── tela_inicial.py          # Tela Inicial do projeto com os botão das funcionalidades
│   ├── Cadastro_Produto.py      # Cadastro de produtos (Leitura de uma planilha preenchido e criação de uma nova padrão athus)
│   ├── Comparação_Prazos.py     # Comparação de prazos ERP com varios Marketplaces
│   └── Extracao_Atributos.py    # Busca de atributos em descrção HTML
│
├── requirements.txt            # Dependências do projeto (caso haja)
├── planilhas/                  # Pasta onde ficam as planilhas de entrada e saída
│   ├── Planilha_Online.xlsx    # Exemplo de planilha de cadastro de produtos
│   ├── Planilha_Athus.xlsx     # Exemplo de planilha de cadastro para o sistema Athus
│   └── Planilha_Onclick.xlsx   # Exemplo de planilha para comparação de prazos

└── README.md          # Este arquivo
```

---

## Como Contribuir

Se você deseja contribuir para o desenvolvimento deste projeto, siga os seguintes passos:

1. Faça um **fork** do repositório.
2. Crie uma **nova branch** para suas alterações:  
   `git checkout -b minha-alteracao`
3. Realize as alterações desejadas e faça o **commit** delas:  
   `git commit -am 'Descrição das alterações'`
4. Envie a branch para o repositório remoto:  
   `git push origin minha-alteracao`
5. Abra uma **pull request** explicando suas alterações.

---

## Licença

Este projeto está licenciado sob a **Licença MIT**. Veja o arquivo [LICENSE](LICENSE) para mais detalhes.

---

## Contato

Para dúvidas ou sugestões, entre em contato com o desenvolvedor:

- **E-mail**: heltoncardososuaves@gmail.com
- **GitHub**: [HeltonCardoso](https://github.com/HeltonCardoso)
