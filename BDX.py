import os
import shutil
from datetime import datetime
import tkinter as tk
from tkinter import messagebox
import sys

 # Pegando data atual
data_atual = datetime.now()
nome_pasta = data_atual.strftime("%d-%m-%Y_%H%M")

# Caminhos
pasta_origem = "Docs"

# Cria a janela principal
janela = tk.Tk()
janela.title("BDX 1.2")  # Título da janela
janela.geometry("600x400")  # Largura x Altura

# Adiciona um botão
def buscar_xml_por_chave():
    arquivo_lista = "chave.txt"

    if os.path.exists(arquivo_lista) and os.path.exists(pasta_origem):
        pass
    else:
        messagebox.showerror(message='Verifique se a pasta ‘Docs’ ou o arquivo ‘chave.txt’ existem na pasta onde está o executável.')
        sys.exit()
    pasta_destino = f"{nome_pasta} BDX CHAVE"
    campo_query.config(state='normal') 
    campo_query.delete("1.0",tk.END)

    # Lê a lista de chaves (sem extensão, sem -nfe)
    chaves = []
    with open(arquivo_lista, "r", encoding="utf-8") as f:
          for linha in f:
              linha = linha.strip()
              linha = linha.lower()
              linha = linha + '-nfe.xml'
              chaves.append(linha.lower())  
    encontrados = set()
    # Percorre a pasta e subpastas
    for raiz, dirs, arquivos in os.walk(pasta_origem):
        for nome_arquivo in arquivos:
            nome_lower = nome_arquivo.lower()
            # Verifica se alguma chave aparece dentro do nome do arquivo
            for chave in chaves:
                 if chave == nome_lower:
                    # Cria a pasta de destino se não existir
                    os.makedirs(pasta_destino, exist_ok=True)
                    caminho_origem = os.path.join(raiz, nome_arquivo)
                    caminho_destino = os.path.join(pasta_destino, nome_arquivo)
                    shutil.copy2(caminho_origem, caminho_destino)
                    encontrados.add(chave)
                    campo_query.insert(tk.END, f"✅ Copiado: {nome_arquivo}\n")
                    break  # evita copiar o mesmo arquivo mais de uma vez

    # Mostra os que não foram encontrados
    nao_encontrados = [c for c in chaves if c not in encontrados]
    if nao_encontrados:
        campo_query.insert(tk.END,"\n⚠️ Arquivos não encontrados:\n")
        for c in nao_encontrados:
            campo_query.insert(tk.END,f"{c}\n")
    else:
        campo_query.insert(tk.END,"\n🟢 Todos os arquivos foram encontrados e copiados!")

    campo_query.config(state='disabled')     
        
def buscar_xml_por_coo():
    chaves=[]
    arquivo_lista = "coo.txt"

    if os.path.exists(arquivo_lista) and os.path.exists(pasta_origem):
        pass
    else:
        messagebox.showerror(message='Verifique se a pasta ‘Docs’ ou o arquivo coo.txt’ existem na pasta onde está o executável.')
        sys.exit()
    campo_query.config(state='normal') 
    campo_query.delete("1.0",tk.END)
    pasta_destino = f"{nome_pasta} BDX COO"
    campo_query.delete("1.0",tk.END)
    with open(arquivo_lista, "r", encoding="utf-8") as f:
          for linha in f:
              linha = linha.strip()
              linha = linha.lower()
              chaves.append(linha.lower())  
    encontrados = set()
    for raiz, dirs, arquivos in os.walk(pasta_origem):
        for nome_arquivo in arquivos:
            if nome_arquivo[-8:] == '-nfe.xml':
                nome_lower = nome_arquivo.lower()
                for chave in chaves:
                    if str(chave).lstrip("0") == nome_lower[25:34].lstrip("0"):
                        # Cria a pasta de destino se não existir
                        os.makedirs(pasta_destino, exist_ok=True)
                        caminho_origem = os.path.join(raiz, nome_arquivo)
                        caminho_destino = os.path.join(pasta_destino, nome_arquivo)
                        shutil.copy2(caminho_origem, caminho_destino)
                        encontrados.add(chave)
                        campo_query.insert(tk.END, f"✅ Copiado: {nome_arquivo}\n")
                        break  # evita copiar o mesmo arquivo mais de uma vez
                    else:
                        pass
            else:
                pass  
    # Mostra os que não foram encontrados
    nao_encontrados = [c for c in chaves if c not in encontrados]
    if nao_encontrados:
        campo_query.insert(tk.END,"\n⚠️ Arquivos não encontrados:\n")
        for c in nao_encontrados:
            campo_query.insert(tk.END,f"{c}\n")
    else:
        campo_query.insert(tk.END,"\n🟢 Todos os arquivos foram encontrados e copiados!")
    campo_query.config(state='disabled')                   
           
botao = tk.Button(janela, text="Buscar xml por chave", command=buscar_xml_por_chave, padx=20, pady=20,fg='white',bg='green')
botao.pack(pady=10)
botao2 = tk.Button(janela, text="Buscar xml por coo", command=buscar_xml_por_coo, padx=20, pady=20,fg='white',bg='green')
botao2.pack(pady=10)

campo_query = tk.Text(janela, width=80, height=15,bg="lightblue",fg="white",state='disabled')
campo_query.pack(pady=5)

# Inicia o loop principal
janela.mainloop()
