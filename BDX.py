import os
import shutil
from datetime import datetime
import tkinter as tk
from tkinter import messagebox
import requests
from OpenSSL import crypto
import xml.etree.ElementTree as ET
import urllib3
import os
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
import threading
import time

data_atual = datetime.now()
nome_pasta = data_atual.strftime("%d-%m-%Y_%H%M")
pasta_origem = 'docs'


def load_pfx(pfx_path, password):
    """Carrega certificado PFX e retorna caminhos do PEM."""
    with open(pfx_path, "rb") as f:
            pfx = f.read()
    p12 = crypto.load_pkcs12(pfx, password.encode())
    private_key = crypto.dump_privatekey(crypto.FILETYPE_PEM, p12.get_privatekey())
    certificate = crypto.dump_certificate(crypto.FILETYPE_PEM, p12.get_certificate())
    open("cert.pem", "wb").write(certificate)
    open("key.pem", "wb").write(private_key)
    return "cert.pem", "key.pem"


def limpar(xml):
    """Remove quebras de linha e BOM do XML."""
    return "".join(l.strip() for l in xml.lstrip("\ufeff").splitlines() if l.strip())


def xml_consulta(chave):
    """Gera XML de consulta de NFe."""
    return f"""
    <consSitNFe xmlns="http://www.portalfiscal.inf.br/nfe" versao="4.00">
        <tpAmb>1</tpAmb>
        <xServ>CONSULTAR</xServ>
        <chNFe>{chave}</chNFe>
    </consSitNFe>
    """.strip()


def montar_soap(xml_nfe):
    """Envolve o XML de consulta em SOAP."""
    return f"""
    <soap12:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                     xmlns:xsd="http://www.w3.org/2001/XMLSchema"
                     xmlns:soap12="http://www.w3.org/2003/05/soap-envelope">
        <soap12:Body>
            <nfeDadosMsg xmlns="http://www.portalfiscal.inf.br/nfe/wsdl/NFeConsultaProtocolo4">
                {xml_nfe}
            </nfeDadosMsg>
        </soap12:Body>
    </soap12:Envelope>
    """.strip()


def extrair_prot(xml_retorno):
    """Extrai cStat, xMotivo, nProt e tpEmis do XML retornado da SEFAZ."""
    ns = {
        "soap": "http://www.w3.org/2003/05/soap-envelope",
        "nfe": "http://www.portalfiscal.inf.br/nfe",
        "ws": "http://www.portalfiscal.inf.br/nfe/wsdl/NFeConsultaProtocolo4"
    }

    root = ET.fromstring(xml_retorno)

    prot = root.find(
        ".//ws:nfeResultMsg/nfe:retConsSitNFe/nfe:protNFe/nfe:infProt",
        ns
    )

    if prot is None:
        return None

    tpEmis = prot.find("nfe:tpEmis", ns)
    return {
        "cStat": prot.find("nfe:cStat", ns).text if prot.find("nfe:cStat", ns) is not None else None,
        "xMotivo": prot.find("nfe:xMotivo", ns).text if prot.find("nfe:xMotivo", ns) is not None else None,
        "nProt": prot.find("nfe:nProt", ns).text if prot.find("nfe:nProt", ns) is not None else None,
        "tpEmis": tpEmis.text if tpEmis is not None else None
    }



def janela():
    global campo_query, janela_principal
    janela_principal = tk.Tk()
    janela_principal.title("BDX 1.4")  # Título da janela
    janela_principal.geometry("600x600")  # Largura x Altura
    
    botao = tk.Button(janela_principal, text="Buscar xml por chave", command=buscar_xml_por_chave, padx=20, pady=20,fg='white',bg='green')
    botao.pack(pady=10)
    botao2 = tk.Button(janela_principal, text="Buscar xml por coo", command=buscar_xml_por_coo, padx=20, pady=20,fg='white',bg='green')
    botao2.pack(pady=10)
    botao3 = tk.Button(janela_principal, text="Validação de xml", command=janela_nova, padx=20, pady=20,fg='white',bg='blue')
    botao3.pack(pady=10)

    campo_query = tk.Text(janela_principal, width=120, height=100,bg="lightblue",fg="white",state='disabled')
    campo_query.pack(pady=5)
    janela_principal.mainloop()


def buscar_xml_por_chave():
    arquivo_lista = "chave.txt"

    if os.path.exists(arquivo_lista) and os.path.exists(pasta_origem):
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
    else:
        messagebox.showerror(message='Verifique se a pasta ‘Docs’ ou o arquivo ‘chave.txt’ existem na pasta onde está o executável.')
    

def buscar_xml_por_coo():
    chaves=[]
    arquivo_lista = "coo.txt"

    if os.path.exists(arquivo_lista) and os.path.exists(pasta_origem):
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
        
    else:
        messagebox.showerror(message='Verifique se a pasta ‘Docs’ ou o arquivo coo.txt’ existem na pasta onde está o executável.')


def validar_xml(pasta,certificado,senha):
    if os.path.exists(pasta) and os.path.exists(certificado):
        certificado = str(certificado).strip()
        senha = str(senha).strip()
        CERT_FILE, KEY_FILE = load_pfx(certificado, senha)
             
        for raiz, dirs, arquivos in os.walk(pasta):
            for nome_arquivo in arquivos:
                if nome_arquivo.endswith('-nfe.xml'):
                    chave = nome_arquivo[25:34].lstrip("0")
                    chave_modificada = nome_arquivo.removesuffix('-nfe.xml')
                    xml_nfe = limpar(xml_consulta(chave_modificada))
                    soap_xml = montar_soap(xml_nfe)
                    url = "https://nfce.fazenda.sp.gov.br/ws/NFeConsultaProtocolo4.asmx"
                    headers = {"Content-Type": "application/soap+xml; charset=utf-8"}

                    response = requests.post(
                        url, data=soap_xml.encode("utf-8"), headers=headers, cert=(CERT_FILE, KEY_FILE), verify=False
                    )

                    ret = extrair_prot(response.text)

                    campo_query1.after(0, lambda: campo_query1.insert(tk.END, "\n" + "*"*30 + "\n"))

                    if ret:  # XML encontrado via SEFAZ
                        campo_query1.after(0, lambda: campo_query1.insert(tk.END,'🟢 XML ENCONTRADO! 🟢\n'))
                        campo_query1.after(0, lambda: campo_query1.insert(tk.END,f"COO: {chave}\n"))
                        campo_query1.after(0, lambda: campo_query1.insert(tk.END,f"CHAVE: {nome_arquivo}\n"))
                        campo_query1.after(0, lambda: campo_query1.insert(tk.END, f"cStat: {ret.get('cStat')}\n"))
                        campo_query1.after(0, lambda: campo_query1.insert(tk.END, f"Motivo: {ret.get('xMotivo')}\n"))
                        campo_query1.after(0, lambda: campo_query1.insert(tk.END, f"Protocolo: {ret.get('nProt')}\n"))
                        tipo = ret.get("tpEmis")
                    else:
                        arquivo_xml = os.path.join(raiz, nome_arquivo)
                        tree = ET.parse(arquivo_xml)
                        root = tree.getroot()

                        ns = {"nfe": "http://www.portalfiscal.inf.br/nfe"}
                        ide = root.find(".//nfe:ide", ns)
                        tipo = ide.findtext("nfe:tpEmis", default=None, namespaces=ns) if ide is not None else None

                        os.remove(arquivo_xml)
                        campo_query1.after(0, lambda: campo_query1.insert(tk.END,'⚠️ XML NAO ENCONTRADO! ⚠️\n'))
                        campo_query1.after(0, lambda: campo_query1.insert(tk.END, f"COO: {chave}\n"))
                        campo_query1.after(0, lambda: campo_query1.insert(tk.END, f"CHAVE: {nome_arquivo}\n"))

                    campo_query1.after(0, lambda tipo=tipo: campo_query1.insert(tk.END, f"Tipo de emissão (tpEmis): {tipo}\n" + "*"*30 + "\n"))
                    campo_query1.after(0, campo_query1.update_idletasks)
                    campo_query1.see(tk.END)
                    time.sleep(0.5)
                
    else:  
        janela2.after(
            0,
            lambda: messagebox.showerror("Erro", "Verifique suas configurações novamente!")
        )

       
def janela_nova():
    global janela2
    global campo_query1
    janela_principal.withdraw()
    janela2 = tk.Toplevel()
    janela2.title("BDX 1.4")
    janela2.geometry("1000x800")

    label_pasta = tk.Label(janela2, text="Caminho da pasta dos XMLs:")
    label_pasta.pack(pady=5)
    entry_pasta = tk.Entry(janela2, width=50)
    entry_pasta.pack(pady=5) 

    label_pasta1 = tk.Label(janela2, text="Caminho do certificado:")
    label_pasta1.pack(pady=5)
    entry_pasta1 = tk.Entry(janela2, width=50)
    entry_pasta1.pack(pady=5) 

    label_pasta2 = tk.Label(janela2, text="Senha do certificado:")
    label_pasta2.pack(pady=5)
    entry_pasta2 = tk.Entry(janela2, width=50)
    entry_pasta2.pack(pady=5) 

    botao5 = tk.Button(janela2, text="Validação de xml", command=lambda: validar_xml_thread(entry_pasta.get(),entry_pasta1.get(),entry_pasta2.get()), padx=20, pady=20,fg='white',bg='green')
    botao5.pack(pady=5)    
    botao6 = tk.Button(janela2, text="Voltar", command= voltar, padx=20, pady=20,fg='white',bg='blue')
    botao6.pack(pady=5)   

    campo_query1 = tk.Text(janela2, width=130, height=100,bg="lightblue",fg="white")
    campo_query1.pack(pady=5,side='bottom')


def validar_xml_thread(pasta,certificado,senha):
    threading.Thread(target=validar_xml,args=(pasta,certificado,senha)).start()


def voltar():
    janela2.withdraw()
    janela_principal.deiconify()


janela()
