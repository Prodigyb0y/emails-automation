# Automação de Relatórios de Produção via Outlook

Este script Python automatiza o envio de relatórios diários de desvio de consumo para diferentes departamentos de produção e gera um resumo gerencial formatado em HTML com dados extraídos do Excel.

## 🚀 Funcionalidades

* **Envio em Lote:** Envia e-mails personalizados com anexos específicos para departamentos (ex: BT, CS, TA).
* **Integração com Excel:** Lê dados de uma planilha mestre usando `win32com` e `pandas`.
* **Relatório HTML:** Gera uma tabela HTML estilizada no corpo do e-mail com base nos dados processados.
* **Assinatura com Imagem:** Incorpora a assinatura do usuário diretamente no corpo do e-mail (base64) para evitar que apareça como anexo bloqueado.
* **Feedback Visual:** Utiliza `tkinter` para exibir um popup ao finalizar o processo.

## 🛠️ Pré-requisitos

* Windows OS (devido à dependência do Outlook/Win32).
* Microsoft Outlook instalado e configurado.
* Microsoft Excel instalado.
* Python 3.x.

### Bibliotecas Necessárias

Instale as dependências utilizando o pip:

```bash
pip install pandas pywin32
