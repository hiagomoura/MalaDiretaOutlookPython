import pandas as pd
import win32com.client as win32
import os

#Lendo a base de dados em CSV:

bdEmail = pd.read_csv("bd.csv",sep=";",encoding = "ANSI", engine='python')
numeroLinhasBdEmail = bdEmail["Nome"].count()
print(f" Número total de linhas importadas do CSV: {numeroLinhasBdEmail}")

#Percorrendo as linhas do banco de dados de emails (dataframe do Pandas) salvando os dados de cada destinatário em variáveis,
# E enviando emails para cada uma delas:

for i in range(numeroLinhasBdEmail):
    
    #Variáveis para os parâmetros do email:

    nomeCompletoDestinatario =  bdEmail["Nome"].loc[i]
    primeiroNomeDestinatario = nomeCompletoDestinatario[0: nomeCompletoDestinatario.index(' ')]
    emailDestinatario = bdEmail["Email"].loc[i]
    assunto = "Parabéns Educador Certificado Microsoft do Século XXI"
    
    #Criando integração com o Cliente de emai Outlook:
    outlook = win32.Dispatch('outlook.application')

    #Criando um email:
    email = outlook.CreateItem(0)

    #Configurar informações do seu e-email:

    email.To = emailDestinatario
    email.Subject = assunto

    # Área de anexo

    #anexo = os.path.abspath(f"anexos/{nomeCompletoDestinatario}.pdf")
    #email.Attachments.Add(anexo)
    

    #Escrevendo o corpo do email usando HTML:
       
    email.HTMLBody = f"""

    <p>Olá {primeiroNomeDestinatario},</p>

    Parabéns, você é um educador do século XXI, certificado internacionalmente pela Microsoft, usa as tecnologias como aliadas do processo educativo, incentiva o trabalho colaborativo e compartilha suas experiências com outros educadores e está sempre aberto a aprender!<br>
    Você valoriza as experiências pessoais dos alunos, unindo o conteúdo trabalhado em sala de aula a contextos autênticos, propiciando o aprendizado além dos muros da escola.

    <p>Eu Hiago, junto a Microsoft reconheço e agradeço por todo o trabalho prestado na educação desse país.</p>

    <p><i>Educação não transforma o mundo. Educação muda as pessoas. Pessoas transformam o mundo.</i></p>

    <b><i>Paulo Freire</i></b>

    <p><b>Seguem em anexo o certificado e o selo Educador Microsoft para usar em apresentações pessoais, assinaturas de email e redes sociais.</b></p>

    """

    email.Send()
    print(f"Email Nº {i} para {nomeCompletoDestinatario} {emailDestinatario} Enviado com sucesso! [OK]")