import os
from PIL import Image
import customtkinter as ctk
from tkinter import ttk
import zipfile
import pandas as pd
import threading
import shutil
import stat
from pathlib import Path
from datetime import datetime

# Verifica se o ambiente suporta tkinter
try:

    import tkinter as tk
    from tkinter import filedialog, messagebox
    import customtkinter as ctk
except ModuleNotFoundError as e:
    raise ImportError("O módulo tkinter ou customtkinter não está instalado. Verifique seu ambiente.") from e

# Tema escuro e cores personalizadas
ctk.set_appearance_mode("light") 

# Cores do sistema
COR_PRIMARIA = "#01274b"
COR_SECUNDARIA = "#b34400"  
GRAY_BUTTON = "#4a4a4a"
TEXT_COLOR = "#f0f0f0"

# Variáveis globais
caminho_planilha = ""
caminho_base = ""
LOG_PATH = str(Path.home() / "Downloads" / "historico_transferencias.csv")

# Funções auxiliares
def zipar_pasta(origem, destino_zip):
    total = 0
    with zipfile.ZipFile(destino_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, _, files in os.walk(origem):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, origem)
                zipf.write(file_path, arcname)
                total += 1
    return total

def remover_pasta_forcado(caminho):
    def onerror(func, path, _):
        os.chmod(path, stat.S_IWRITE)
        func(path)
    shutil.rmtree(caminho, onerror=onerror)

def registrar_historico(pasta, status, origem, destino):
    df = pd.DataFrame([{
        "Data": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Pasta": pasta,
        "Status": status,
        "Origem": origem,
        "Destino": destino
    }])
    if Path(LOG_PATH).exists():
        df.to_csv(LOG_PATH, mode='a', header=False, index=False)
    else:
        df.to_csv(LOG_PATH, index=False)

def encurtar_caminho(caminho, partes=2):
    partes_caminho = caminho.replace("\\", "/").split("/")
    return "/".join(partes_caminho[-partes:])

# Funções principais

# Função para permitir escolha de processos antes de mover
def selecionar_processos(callback_movimentacao):
    if not caminho_planilha:
        messagebox.showwarning("Atenção", "Selecione a planilha primeiro.")
        return

    try:
        df = pd.read_excel(caminho_planilha)
        processos = df.iloc[:, 0].astype(str).tolist()

        if not processos:
            messagebox.showinfo("Aviso", "Nenhum processo encontrado na planilha.")
            return

        win = ctk.CTkToplevel(janela)
        win.title("Selecionar Processos")
        win.configure(fg_color="#2b2b2b")

        largura_janela = 400
        altura_janela = 540
        win.geometry(f"{largura_janela}x{altura_janela}")

        # Centraliza no centro da tela
        win.update_idletasks()
        largura_tela = win.winfo_screenwidth()
        altura_tela = win.winfo_screenheight()
        pos_x = int((largura_tela - largura_janela) / 2)
        pos_y = int((altura_tela - altura_janela) / 2)
        win.geometry(f"{largura_janela}x{altura_janela}+{pos_x}+{pos_y}")
        win.transient(janela)

        checkboxes = []
        frame_lista = ctk.CTkScrollableFrame(win, fg_color="#2c2c2c", height=400)
        frame_lista.pack(pady=10, padx=10, fill="both", expand=True)

        for processo in processos:
            var = tk.BooleanVar()
            chk = ctk.CTkCheckBox(
                frame_lista,
                text=processo,
                variable=var,
                text_color=TEXT_COLOR,
                fg_color="#444444",
                border_color="#aaaaaa",
                hover_color="#666666",
                checkmark_color=COR_SECUNDARIA
            )
            chk.pack(anchor="w", pady=2, padx=10)
            checkboxes.append((processo, var))

        # Variável e função para alternar todos
        todos_selecionados = tk.BooleanVar(value=False)

        def ao_alternar_todos():
            estado = todos_selecionados.get()
            for _, var in checkboxes:
                var.set(estado)
            checkbox_toggle.configure(text="Desmarcar todos" if estado else "Selecionar todos")

        def confirmar():
            selecionados = [proc for proc, var in checkboxes if var.get()]
            if not selecionados:
                messagebox.showwarning("Atenção", "Nenhum processo selecionado.")
                return
            win.destroy()
            callback_movimentacao(selecionados)

        # Frame inferior fixo com botão e checkbox
        botoes_frame = ctk.CTkFrame(win, fg_color="transparent")
        botoes_frame.pack(fill="x", pady=10, padx=10)

        # Botão Confirmar
        ctk.CTkButton(
            botoes_frame,
            text="Confirmar Seleção",
            command=confirmar,
            fg_color=COR_PRIMARIA,
            hover_color="#011e38",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color=TEXT_COLOR,
            width=150
        ).pack(side="right", padx=10)

        # CheckBox "Selecionar Todos" / "Desmarcar Todos"
        checkbox_toggle = ctk.CTkCheckBox(
            botoes_frame,
            text="Selecionar Todos",
            font=ctk.CTkFont(size=11, weight="bold"),
            variable=todos_selecionados,
            command=ao_alternar_todos,
            text_color=TEXT_COLOR,
            fg_color="#444444",
            border_color="#aaaaaa",
            hover_color="#666666",
            checkmark_color=COR_SECUNDARIA
        )
        checkbox_toggle.pack(side="left", padx=10)

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao carregar planilha:\n{e}")

# Função para selecionar os arquivos que deseja transferir
def selecionar_arquivo(tipo):
    global caminho_planilha, caminho_base

    if tipo == 'planilha':
        arquivo = filedialog.askopenfilename(filetypes=[("Planilhas Excel", "*.xlsx *.xls")])
        if arquivo:
            caminho_planilha = arquivo
            botao_arquivo.configure(text=encurtar_caminho(arquivo), fg_color="#2b2b2b")

    elif tipo == 'pasta':
        pasta = filedialog.askdirectory()
        if pasta:
            caminho_base = pasta
            botao_pasta.configure(text=encurtar_caminho(pasta), fg_color="#2b2b2b")

# Atualize mover_para_finalizado com contagem de arquivos
def mover_para_finalizado(selecionados):
    try:
        df = pd.read_excel(caminho_planilha)
        em_andamento_dir = os.path.join(caminho_base, 'EM ANDAMENTO')
        finalizado_dir = os.path.join(caminho_base, 'FINALIZADO')
        os.makedirs(finalizado_dir, exist_ok=True)

        total_arquivos = 0
        detalhes = []

        for _, row in df.iterrows():
            nome_pasta = str(row.iloc[0]).strip()
            status = str(row.iloc[1]).strip().lower()

            if nome_pasta not in selecionados:
                continue

            if status == 'finalizado':
                origem = os.path.join(em_andamento_dir, nome_pasta)
                destino_zip = os.path.join(finalizado_dir, f"{nome_pasta}.zip")

                if os.path.exists(origem):
                    qtde = zipar_pasta(origem, destino_zip)
                    total_arquivos += qtde
                    detalhes.append(f"{nome_pasta}: {qtde} de {qtde} arquivos")
                    remover_pasta_forcado(origem)
                    registrar_historico(nome_pasta, f"{status} (zipado)", origem, destino_zip)
                else:
                    detalhes.append(f"{nome_pasta}: pasta não encontrada")

        resumo = "\n".join(detalhes)
        resumo += f"\nTotal: {total_arquivos} de {total_arquivos} arquivos"
        messagebox.showinfo("Sucesso", f"Processos transferidos para FINALIZADO.\n\n{resumo}")

    except Exception as e:
        messagebox.showerror("Erro", str(e))

# Atualize mover_para_em_andamento com contagem de arquivos
def mover_para_em_andamento(selecionados):
    try:
        df = pd.read_excel(caminho_planilha)
        em_andamento_dir = os.path.join(caminho_base, 'EM ANDAMENTO')
        finalizado_dir = os.path.join(caminho_base, 'FINALIZADO')
        os.makedirs(em_andamento_dir, exist_ok=True)

        total_arquivos = 0
        detalhes = []

        for _, row in df.iterrows():
            nome_pasta = str(row.iloc[0]).strip()
            status = str(row.iloc[1]).strip().lower()

            if nome_pasta not in selecionados:
                continue

            if status in ['em execução', 'em elaboração de relatórios', 'em pré-teste']:
                origem_zip = os.path.join(finalizado_dir, f"{nome_pasta}.zip")
                destino = os.path.join(em_andamento_dir, nome_pasta)

                if not os.path.exists(origem_zip):
                    detalhes.append(f"{nome_pasta}: ZIP não encontrado")
                    continue
                if os.path.exists(destino):
                    detalhes.append(f"{nome_pasta}: já existe em EM ANDAMENTO")
                    continue

                os.makedirs(destino, exist_ok=True)
                with zipfile.ZipFile(origem_zip, 'r') as zip_ref:
                    zip_ref.extractall(destino)
                    qtde = len(zip_ref.namelist())
                    total_arquivos += qtde
                    detalhes.append(f"{nome_pasta}: {qtde} de {qtde} arquivos extraídos")
                os.remove(origem_zip)
                registrar_historico(nome_pasta, f"{status} (descompactado)", origem_zip, destino)

        resumo = "\n".join(detalhes)
        resumo += f"\nTotal: {total_arquivos} de {total_arquivos} arquivos"
        messagebox.showinfo("Sucesso", f"Processos transferidos para EM ANDAMENTO.\n\n{resumo}")

    except Exception as e:
        messagebox.showerror("Erro", str(e))

# Função para ver historico
def ver_historico():
    if not Path(LOG_PATH).exists():
        messagebox.showinfo("Histórico", "Nenhuma transferência registrada ainda.")
        return

    win = ctk.CTkToplevel(janela)
    win.title("Histórico de Transferências")

    win.configure(fg_color="#2b2b2b")

    largura_janela = 920
    altura_janela = 500
    win.geometry(f"{largura_janela}x{altura_janela}")

    # Centraliza no centro da tela
    win.update_idletasks()
    largura_tela = win.winfo_screenwidth()
    altura_tela = win.winfo_screenheight()
    pos_x = int((largura_tela - largura_janela) / 2)
    pos_y = int((altura_tela - altura_janela) / 2)
    win.geometry(f"{largura_janela}x{altura_janela}+{pos_x}+{pos_y}")
    win.transient(janela)

    win.transient(janela) # Faz com que a janela do histórico apareça na frente

    win.resizable(False, False) # Não permite que a janela seja redimensionada

    # Configurações de estilo para a Treeview (necessário para customtkinter)
    style = ttk.Style()
    style.theme_use("default") # Usa o tema padrão do ttk para customizar
    style.configure("Treeview",
                    background="#333333",
                    foreground=TEXT_COLOR,
                    fieldbackground="#333333",
                    bordercolor="#555555",
                    rowheight=25)
    style.map('Treeview',
              background=[('selected', COR_SECUNDARIA)]) 

    style.configure("Treeview.Heading",
                    font=("Roboto", 10, "bold"),
                    background="#555555",
                    foreground=TEXT_COLOR,
                    relief="flat")
    style.map("Treeview.Heading",
              background=[('active', '#666666')]) 

    entrada_filtro = ctk.CTkEntry(win, placeholder_text="Filtrar por pasta",
                                  fg_color="#333333", text_color=TEXT_COLOR,
                                  border_color=COR_SECUNDARIA)
    entrada_filtro.pack(pady=10, padx=10, fill="x")

    tree = ttk.Treeview(win, columns=["Data", "Pasta", "Status", "Origem", "Destino"], show="headings", style="Treeview")
    for col in ["Data", "Pasta", "Status", "Origem", "Destino"]:
        tree.heading(col, text=col)
        tree.column(col, anchor="center", width=150)
    tree.pack(padx=10, pady=5, fill="both", expand=True)

    def carregar_historico(filtro=""):
        for item in tree.get_children():
            tree.delete(item)
        df = pd.read_csv(LOG_PATH)
        if filtro:
            df = df[df["Pasta"].str.contains(filtro, case=False, na=False)]
        for _, row in df.iterrows():
            tree.insert("", "end", values=row.tolist())

    def limpar_log():
        if messagebox.askyesno("Confirmar", "Tem certeza que deseja limpar o histórico?"):
            Path(LOG_PATH).unlink(missing_ok=True)
            carregar_historico()
            messagebox.showinfo("Sucesso", "Histórico apagado.")

    botao_frame = ctk.CTkFrame(win, fg_color="transparent") # Frame transparente para os botões
    botao_frame.pack(pady=5)

    ctk.CTkButton(botao_frame, 
                  text="Filtrar", 
                  command=lambda: carregar_historico(entrada_filtro.get()),
                  fg_color=COR_PRIMARIA, 
                  hover_color="#011e38", 
                  font=ctk.CTkFont(size=12, weight="bold"),
                  text_color=TEXT_COLOR).pack(side="left", padx=10)
    
    ctk.CTkButton(botao_frame, 
                  text="Limpar Histórico", 
                  command=limpar_log,
                  fg_color="red", 
                  font=ctk.CTkFont(size=12, weight="bold"),
                  hover_color="#6d0505", 
                  text_color=TEXT_COLOR).pack(side="left", padx=10)

    carregar_historico()

# Interface principal
janela = ctk.CTk()
janela.title("Transferência de Processos ")

# Dimensões da janela
largura_janela = 600
altura_janela = 445

# Obtém dimensões da tela
largura_tela = janela.winfo_screenwidth()
altura_tela = janela.winfo_screenheight()

# Calcula posição x e y
pos_x = int((largura_tela - largura_janela) / 2)
pos_y = int((altura_tela - altura_janela) / 2)

# Define geometria centralizada
janela.geometry(f"{largura_janela}x{altura_janela}+{pos_x}+{pos_y}")

janela.resizable(False, False) # Impede o redimensionamento da janela principal

# Configurações de cores para a janela principal
janela.configure(fg_color="#2b2b2b")

titulo = ctk.CTkLabel(janela, text="Transferência de Processos",
                      font=ctk.CTkFont(family="Roboto", size=24, weight="bold"),
                      text_color=TEXT_COLOR)
titulo.pack(pady=20)

# Descrição acima do frame principal
descricao = ctk.CTkLabel(
    janela,
    text="Escolha o arquivo e a pasta base para iniciar a transferência.",
    font=ctk.CTkFont(size=15, weight="bold"),
    text_color="#cccccc",
    justify="center"
)
descricao.pack(padx=20)

# Frame principal
frame_principal = ctk.CTkFrame(janela, 
                               height=70, 
                                corner_radius=8,
                               fg_color="#3c3c3c")
frame_principal.pack(padx=25, pady=10, fill="x")
frame_principal.pack_propagate(False)


# Frame para selecionar arquivo Excel
frame_arquivo = ctk.CTkFrame(frame_principal, 
                             fg_color="#3c3c3c", 
                             width=200)
frame_arquivo.pack(side="left", padx=10, pady=5, fill="both", expand=True)

# Frame para selecionar a pasta
frame_pasta = ctk.CTkFrame(frame_principal, 
                           fg_color="#3c3c3c", 
                           width=200)
frame_pasta.pack(side="left", padx=10, pady=5, fill="both", expand=True)

# Botão para selecionar arquivo Excel
icone_arquivo = ctk.CTkImage(Image.open(r"C:\Projetos\icons\excel.png"), size=(22, 22))

botao_arquivo = ctk.CTkButton(
    frame_arquivo,
    text="Selecione o arquivo",
    image=icone_arquivo,
    compound="left",
    fg_color=GRAY_BUTTON,
    hover_color="#5a5a5a",
    text_color="white",
    font=ctk.CTkFont(size=12, weight="bold"),
    width=250,
    height=40,
    corner_radius=12,
    command=lambda: selecionar_arquivo("planilha")
)
botao_arquivo.pack(expand=True, anchor="center")

# Botão para selecionar a pasta
icone_pasta = ctk.CTkImage(Image.open(r"C:\Projetos\icons\pasta.png"), size=(30, 30))

botao_pasta = ctk.CTkButton(
    frame_pasta,
    text="Selecione a pasta",
    image=icone_pasta,
    compound="left",
    fg_color=GRAY_BUTTON,
    hover_color="#5a5a5a",
    text_color=TEXT_COLOR,
    corner_radius=12,
    height=40,
    width=250,          
    font=ctk.CTkFont(size=12, weight="bold"),
    command=lambda: selecionar_arquivo('pasta')
)
botao_pasta.pack(expand=True, anchor="center")

# Frame para ações
frame_acoes = ctk.CTkFrame(janela,corner_radius=12, fg_color="#3c3c3c")
frame_acoes.pack(padx=25, pady=15, fill="both", expand=True)

# Botão para mover processos para pasta FINALIZADO
ctk.CTkButton(
    frame_acoes,
    text="Mover processos para: FINALIZADO",
    command=lambda: selecionar_processos(mover_para_finalizado),
    fg_color=COR_PRIMARIA, 
    corner_radius=8,
    hover_color="#011e38",
    text_color=TEXT_COLOR,
    font=ctk.CTkFont(size=14, weight="bold"),
    height=40
).pack(pady=12, padx=20, fill="x")

# Botão para mover processos para pasta EM ANDAMENTO
ctk.CTkButton(
    frame_acoes,
    text="Mover processos para: EM ANDAMENTO",
    command=lambda: selecionar_processos(mover_para_em_andamento),
    fg_color=COR_SECUNDARIA,
    corner_radius=8,
    hover_color="#cc5500",
    font=ctk.CTkFont(size=14, weight="bold"),
    height=40 
).pack(padx=20, fill="x")

# Botão para vizualizar o histórico de transferências
icone_historico = ctk.CTkImage(Image.open(r"C:\Projetos\icons\historico.png"), size=(30, 30))

ctk.CTkButton(
    janela,
    text="Histórico de Transferências",
    command=ver_historico,
    image=icone_historico,
    compound="left",
    fg_color=GRAY_BUTTON,
    hover_color="#5a5a5a",
    text_color=TEXT_COLOR,
    font=ctk.CTkFont(size=12, weight="bold"),
    height=38
).pack(pady=12, padx=20)

# Rodapé
ctk.CTkLabel(janela, text="v1.0",
             font=ctk.CTkFont(size=12), text_color="#aaaaaa").pack(side="right", padx=15, pady=10)

janela.mainloop()