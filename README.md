# 🗂 ControleProcessos  
Automação de Transferência de Pastas (2025)  

Este projeto é um **software em Python** desenvolvido para realizar a **automação da movimentação interna de pastas** contendo registros de clientes dentro de um servidor, com base em um arquivo Excel de configuração.

---

## ✨ Funcionalidades  
✔️ Leitura de arquivo Excel contendo mapeamento de pastas/clientes  
✔️ Identificação e movimentação automática de pastas dentro do servidor  
✔️ Registro de logs de operações para auditoria  
✔️ Configuração simples para adaptar caminhos e regras de transferência  

---

## 🛠️ Tecnologias Utilizadas  
🐍 **Python 3.x**  
📂 Manipulação de arquivos e diretórios (os, shutil)  
📊 Leitura de planilha Excel (por exemplo, `pandas` ou `openpyxl`)  
📁 Funções de automação de sistema de arquivos  

---

## 🧩 Pré-requisitos  
- Python instalado (versão 3.x)  
- Biblioteca para leitura de Excel (ex: `pandas`, `openpyxl`)  
- Acesso aos diretórios/pastas do servidor conforme regras de movimentação  
- Arquivo Excel de mapeamento devidamente formatado  

---

## 🚀 Como Executar  

1. **Clone o repositório:**

   ```bash
   git clone https://github.com/bielwzz/ControleProcessos.git
   cd ControleProcessos
   
2. **Instale as dependências:**

pip install -r requirements.txt
Prepare o arquivo Excel de configuração:
Edite o arquivo (ex: mapeamento.xlsx) indicando colunas como “cliente”, “origem”, “destino”.

3. **Execute o software:**

python main.py --config mapeamento.xlsx
