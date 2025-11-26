import os
import base64
import time
import pandas as pd
import locale
import win32com.client as win32
import tkinter as tk
from tkinter import messagebox

# Configurações Globais
CAMINHO_BASE = r"C:\Caminho\Para\Seus\Arquivos"
CAMINHO_ASSINATURA = r"C:\Caminho\Para\assinatura.png"
LOCALE_PT_BR = 'pt_BR.UTF-8'

def obter_imagem_base64(caminho_imagem):
    """Lê uma imagem e retorna sua string em base64."""
    try:
        with open(caminho_imagem, "rb") as image_file:
            return base64.b64encode(image_file.read()).decode('utf-8')
    except Exception as e:
        print(f"Erro ao ler imagem de assinatura: {e}")
        return ""

def gerar_corpo_padrao(departamento, assinatura_b64):
    """Gera o HTML padrão para os emails de departamento."""
    return f"""
    <div style='font-family: Calibri; font-size: 11pt;'> 
        <p style='margin-top: 20px; margin-bottom: 10px;'>
        Bom dia, <br><br>
        Segue em anexo a Planilha da <b>{departamento}</b>, contendo todas ordens de produção que faltam consumo de componente.<br>
        Sendo assim, por favor, efetuem o consumo, para evitar erros de valorização e saldo no nosso estoque.<br>
        <br><br>
        Atenciosamente – Best regards<br>
        </p>
        <p>
        <b>Seu Nome</b><br>
        Seu Cargo<br>
        Seu Departamento<br>
        <br>
        <img src='data:image/png;base64,{assinatura_b64}'/>
        </p>
        <p>
        Sua Empresa<br>
        Seu Endereço<br>
        <a href='mailto:seuemail@empresa.com'>seuemail@empresa.com</a> · <a href='www.site.com'>www.site.com</a>
        </p>
    </div>
    """

def enviar_email(destinatarios, copia, assunto, corpo_html, anexo=None):
    """Função genérica para envio de e-mail via Outlook."""
    try:
        outlook = win32.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        mail.To = destinatarios
        mail.CC = copia
        mail.Subject = assunto
        mail.HTMLBody = corpo_html
        
        if anexo and os.path.exists(anexo):
            mail.Attachments.Add(anexo)
        elif anexo:
            print(f"Aviso: Anexo não encontrado em {anexo}")

        mail.Send()
        print(f"E-mail '{assunto}' enviado com sucesso!")
        return True
    except Exception as e:
        print(f"Erro ao enviar e-mail '{assunto}': {e}")
        return False

def processar_relatorios_departamentos(assinatura_b64):
    """Processa e envia os e-mails padrão para BT, CS e TA."""
    
    # Configuração de cada departamento
    config_deptos = [
        {
            "nome": "MU-BT",
            "arquivo": "BT.xlsx",
            "to": "gestor1@empresa.com; equipe_bt@empresa.com",
            "cc": "supervisor@empresa.com"
        },
        {
            "nome": "MU-CS",
            "arquivo": "CS.xlsx",
            "to": "gestor2@empresa.com; equipe_cs@empresa.com",
            "cc": "supervisor@empresa.com"
        },
        {
            "nome": "MU-TA",
            "arquivo": "TA.xlsx",
            "to": "gestor3@empresa.com; equipe_ta@empresa.com",
            "cc": "supervisor@empresa.com"
        }
    ]

    for depto in config_deptos:
        caminho_anexo = os.path.join(CAMINHO_BASE, depto["arquivo"])
        assunto = f"Desvio de Consumo Ordens - {depto['nome']}"
        corpo = gerar_corpo_padrao(depto['nome'], assinatura_b64)
        
        enviar_email(depto["to"], depto["cc"], assunto, corpo, caminho_anexo)
        time.sleep(2) # Pausa para evitar travamento do Outlook

def processar_relatorio_resumo(assinatura_b64):
    """Processa o Excel, gera tabela HTML e envia o e-mail gerencial."""
    try:
        locale.setlocale(locale.LC_ALL, LOCALE_PT_BR)
    except:
        locale.setlocale(locale.LC_ALL, '') # Fallback para o padrão do sistema

    caminho_excel_base = os.path.join(CAMINHO_BASE, "Basei.xlsx")
    
    if not os.path.exists(caminho_excel_base):
        print(f"Arquivo base não encontrado: {caminho_excel_base}")
        return

    try:
        # Automação Excel para ler células específicas (mantendo lógica original)
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(caminho_excel_base)
        ws = wb.Worksheets("Geral")

        dados = []
        # Lendo intervalo B2:C8
        for row in range(2, 9): 
            rotulo = ws.Cells(row, 2).Value
            valor = ws.Cells(row, 3).Value
            if rotulo: # Evita linhas vazias
                dados.append([str(rotulo), valor])
        
        wb.Close(False)
        excel.Quit()

        # Processamento com Pandas
        df = pd.DataFrame(dados, columns=["Planta/Segmento", "Soma de Desvio Consumo"])
        
        # Limpeza e Formatação
        df["Planta/Segmento"] = df["Planta/Segmento"].str.replace(".", "", regex=False).str.replace("0$", "", regex=True)
        
        # Cálculo do total
        total_val = pd.to_numeric(df["Soma de Desvio Consumo"], errors='coerce').fillna(0).sum()
        
        # Formatação monetária
        def formatar_moeda(v):
            return locale.currency(float(v), grouping=True) if isinstance(v, (int, float)) else v

        df["Soma de Desvio Consumo"] = df["Soma de Desvio Consumo"].apply(formatar_moeda)
        
        # Adiciona Total
        df.loc[len(df)] = ["0050", formatar_moeda(total_val)]
        
        # Remove a primeira linha se necessário (mantendo lógica original)
        if len(df) > 1:
            df = df.iloc[1:].copy()

        # Geração da Tabela HTML
        html_rows = ""
        for index, row in df.iterrows():
            segmento = row['Planta/Segmento']
            valor = row['Soma de Desvio Consumo']
            
            bgcolor = "#D9E1F2" # Padrão (Cinza claro/Azul claro)
            estilo_fonte = ""

            if segmento in ["701", "702", "0050"]:
                segmento = f"<b>{segmento}</b>"
                valor = f"<b>{valor}</b>"
                bgcolor = "#B4C6E7" # Destaque
            elif index % 2 == 0:
                bgcolor = "#B4C6E7"

            html_rows += f"""
            <tr style="{estilo_fonte}" bgcolor="{bgcolor}">
                <td>{segmento}</td>
                <td>{valor}</td>
            </tr>"""

        html_table = f"""
        <table border="1" cellpadding="5" style="border-collapse: collapse; font-family: Calibri; font-size: 11pt;">
          <tr>
            <th bgcolor="#203864" style="color: white;">Planta/Segmento</th> 
            <th bgcolor="#203864" style="color: white;">Soma de Desvio Consumo</th>
          </tr>
          {html_rows}
        </table>
        """

        # Corpo do Email Resumo
        corpo_resumo = f"""
        <div style='font-family: Calibri; font-size: 11pt;'> 
            <p>Bom dia Gestor,</p>
            <p>Este é nosso potencial de redução de estoque de MP e SFG, caso a produção corrija o consumo.</p>
            {html_table}
            <br>
            <p>Atenciosamente,</p>
            <img src='data:image/png;base64,{assinatura_b64}'/>
        </div>
        """

        enviar_email(
            destinatarios="diretor@empresa.com",
            copia="gerente@empresa.com",
            assunto="Potencial Redução Estoque - Desvio de Consumo",
            corpo_html=corpo_resumo
        )

    except Exception as e:
        print(f"Erro no processamento do resumo: {e}")

def main():
    print("Iniciando automação...")
    assinatura = obter_imagem_base64(CAMINHO_ASSINATURA)
    
    if not assinatura:
        print("Aviso: Assinatura não carregada. Enviando sem imagem.")

    # 1. Envia emails operacionais
    processar_relatorios_departamentos(assinatura)
    
    # 2. Envia email gerencial
    processar_relatorio_resumo(assinatura)

    # Feedback final
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo("Concluído", "Processo de envio de e-mails finalizado!")
    root.destroy()

if __name__ == "__main__":
    main()
