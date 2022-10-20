##IMPORTAR ARQUIVOS E BIBLIOTECAS

import pandas as pd
import pathlib
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

emails = pd.read_excel(r'Bases de Dados\Emails.xlsx')
lojas = pd.read_csv(r'Bases de Dados\Lojas.csv', encoding='latin1', sep=';')
vendas = pd.read_excel(r'Bases de Dados\Vendas.xlsx')
#mesclando as tabelas lojas e vendas
lojas_df = lojas.merge(vendas, on='ID Loja')


#DEFINIR CRIAR UMA TABELA PARA CADA LOJA E DEFINIR O DIA DO INDICADOR

#foi usado o for para percorrer a tabela lojas e usando um dicionario vazio armazenar as irformações nele e usando também a logica do .loc
dicionario_lojas ={}
for loja in lojas['Loja']:
    dicionario_lojas[loja] = lojas_df.loc[lojas_df['Loja']==loja, :]

#indicador serve para pegar a data mais recente, então se a planilha for atualizada, ele sempre vai pegar o dia mais recente.
dia_indicador = lojas_df['Data'].max()
print(dia_indicador)


#SALVAR A PLANILHA NA PASTA DE BACKUP

#identificar se a pasta já existe.
caminho_backup = pathlib.Path(r'Backup Arquivos Lojas')
arquivos_pasta_backup = caminho_backup.iterdir()

lista_nomes_backup = [arquivo.name for arquivo in arquivos_pasta_backup]
#criando a pasta caso ela não exista.
for loja in dicionario_lojas:
    if loja not in lista_nomes_backup:
        nova_pasta = (caminho_backup / loja).mkdir()
    #salvando dentro da pasta
    local_arquivo = caminho_backup / loja / '{}_{}_{}.xlsx'.format(dia_indicador.month, dia_indicador.day, loja)
    dicionario_lojas[loja].to_excel(local_arquivo)


#metas
meta_faturamento_ano = 1650000
meta_faturamento_dia = 1000
meta_diversidade_ano = 120
meta_diversidade_dia = 4
meta_ticketmedio_ano = 500
meta_ticketmedio_dia = 500

# indicador
for loja in dicionario_lojas:

    vendas_loja = dicionario_lojas[loja]
    vendas_loja_dia = vendas_loja.loc[vendas_loja['Data'] == dia_indicador, :]

    # faturamento loja
    faturamento_ano = vendas_loja['Valor Final'].sum()
    faturamento_dia = vendas_loja_dia['Valor Final'].sum()

    # diversidade de produtos
    qntd_produtos_ano = len(vendas_loja['Produto'].unique())
    qntd_produtos_dia = len(vendas_loja_dia['Produto'].unique())

    # ticket medio
    ticket_medio_ano = vendas_loja.groupby('Código Venda').sum()
    ticket_medio_ano = ticket_medio_ano['Valor Final'].mean()
    ticket_medio_dia = vendas_loja_dia.groupby('Código Venda').sum()
    ticket_medio_dia = ticket_medio_dia['Valor Final'].mean()

    nome = emails.loc[emails['Loja'] == loja, 'Gerente'].values[0]

    # INDICADORES DOS CENARIOS
    if faturamento_dia >= meta_faturamento_dia:
        res_fat_dia = 'Positivo'
    else:
        res_fat_dia = 'Negativo'
    if faturamento_ano >= meta_faturamento_ano:
        res_fat_ano = 'Positivo'
    else:
        res_fat_ano = 'Negativo'
    if qntd_produtos_dia >= meta_diversidade_dia:
        res_qntd_dia = 'Positivo'
    else:
        res_qntd_dia = 'Negativo'
    if qntd_produtos_ano >= meta_diversidade_ano:
        res_qntd_ano = 'Positivo'
    else:
        res_qntd_ano = 'Negativo'
    if ticket_medio_dia >= meta_ticketmedio_dia:
        res_ticket_dia = 'Positivo'
    else:
        res_ticket_dia = 'Negativo'
    if ticket_medio_ano >= meta_ticketmedio_ano:
        res_ticket_ano = 'Positivo'
    else:
        res_ticket_ano = 'Negativo'

    # 1- STARTAR O SERVIDOR SMTP
    host = 'smtp.gmail.com'
    port = '587'
    login = 'maykon.rubens@gmail.com'
    senha = 'mwcxukhhwfmbkdjl'

    # Dando start no servidor
    server = smtplib.SMTP(host,port)
    server.ehlo()
    server.starttls()
    server.login(login, senha)

    # 2- CONSTRUIR O EMAIL TIPO MIME
    corpo = f'''<p>Bom dia, {nome}</p>
    <p> O resultado de ontem <b>({dia_indicador.day}/{dia_indicador.month})</b> da <b>Loja {loja}</b> foi:</p>

    <table>
      <tr>
        <th>Indicador</th>
        <th>Valor Dia</th>
        <th>Meta Dia</th>
        <th>Cenário Dia</th>
      </tr>
      <tr>
        <td>Faturamento</td>
        <td style="text-align: center">R${faturamento_dia:.2f}</td>
        <td style="text-align: center">R${meta_faturamento_dia:.2f}</td>
        <td style="text-align: center">{res_fat_dia}</td>
      </tr>
      <tr>
        <td>Diversidade de Produtos</td>
        <td style="text-align: center">{qntd_produtos_dia}</td>
        <td style="text-align: center">{meta_diversidade_dia}</td>
        <td style="text-align: center">{res_qntd_dia}</td>
      </tr>
      <tr>
        <td>Ticket Médio</td>
        <td style="text-align: center">R${ticket_medio_dia:.2f}</td>
        <td style="text-align: center">R${meta_ticketmedio_dia:.2f}</td>
        <td style="text-align: center">{res_ticket_dia}</td>
      </tr>
    </table>
    <br>
    <table>
      <tr>
        <th>Indicador</th>
        <th>Valor Ano</th>
        <th>Meta Ano</th>
        <th>Cenário Ano</th>
      </tr>
      <tr>
        <td>Faturamento</td>
        <td style="text-align: center">R${faturamento_ano:.2f}</td>
        <td style="text-align: center">R${meta_faturamento_ano:.2f}</td>
        <td style="text-align: center">{res_fat_ano}</td>
      </tr>
      <tr>
        <td>Diversidade de Produtos</td>
        <td style="text-align: center">{qntd_produtos_ano}</td>
        <td style="text-align: center">{meta_diversidade_ano}</td>
        <td style="text-align: center">{res_qntd_ano}</td>
      </tr>
      <tr>
        <td>Ticket Médio</td>
        <td style="text-align: center">R${ticket_medio_ano:.2f}</td>
        <td style="text-align: center">R${meta_ticketmedio_ano:.2f}</td>
        <td style="text-align: center">{res_ticket_ano}</td>
      </tr>
    </table>

    <p>Segue em anexo a planilha com todos os dados para mais detalhes.</p>

    <p>Qaulquer dúvida estou à disposição.</p>
    <p>Att., Maykon</p>
    '''
    email_msg = MIMEMultipart()
    email_msg['From'] = login
    email_msg['To'] = emails.loc[emails['Loja'] == loja, 'E-mail'].values[0]
    email_msg['Subject'] = f'OnePage Dia {dia_indicador.day}/{dia_indicador.month} - Loja {loja}'
    email_msg.attach(MIMEText(corpo, 'html'))  # serve para anexar o corpo no email.

    # 3- ENVIANDO ANEXO / Abrindo o arquivo em modo leitura e binary
    cam_arquivo = pathlib.Path.cwd() / caminho_backup / loja / f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx'
    attachment = open(cam_arquivo, 'rb')  # rb = READ BINARY / LEITURA BINARIA
    # Lemos o arquivo no modo binario e jogamos codificado em base 64 (que é oque o email precisa)
    att = MIMEBase('application', 'octet-stream')
    att.set_payload(attachment.read())
    encoders.encode_base64(att)
    # Adicionamos o cabeçalho no tipo anexo de email
    att.add_header('Content-Disposition', f'attachment; filename={dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx')
    # fechamos o arquivo
    attachment.close()
    # colocamos o anexo no corpo do email
    email_msg.attach(att)

    # 4- ENVIAR O EMAIL TIPO MIME NO SERVIDOR SMTP
    server.sendmail(email_msg['From'], email_msg['To'], email_msg.as_string())

    server.quit()
    print(f'E-mail da Loja {loja} Enviado.')


#CRIANDO RANKING PARA DIRETORIA

faturamento_lojas = lojas_df.groupby('Loja')[['Loja', 'Valor Final']].sum()
faturamento_lojas_ano = faturamento_lojas.sort_values(by='Valor Final', ascending=False)

nome_arquivo = f'{dia_indicador.month}_{dia_indicador.day}_Ranking Anual.xlsx'
faturamento_lojas_ano.to_excel(caminho_backup / f'{nome_arquivo}')

vendas_dia = lojas_df.loc[lojas_df['Data']==dia_indicador, :]
faturamento_lojas_dia = vendas_dia.groupby('Loja')[['Loja', 'Valor Final']].sum().sort_values(by='Valor Final', ascending=False)

nome_arquivo = f'{dia_indicador.month}_{dia_indicador.day}_Ranking Dia.xlsx'
faturamento_lojas_dia.to_excel(caminho_backup / f'{nome_arquivo}')


#ENVIANDO E-MAIL PARA DIRETORIA

#1- STARTAR O SERVIDOR SMTP
host = 'smtp.gmail.com'
port = '587'
login = 'maykon.rubens@gmail.com'
senha = 'mwcxukhhwfmbkdjl'

#Dando start no servidor
server = smtplib.SMTP(host,port)
server.ehlo()
server.starttls()
server.login(login,senha)

#2- CONSTRUIR O EMAIL TIPO MIME
corpo = f'''
Prezados, bom dia

Melhor Loja do Dia em Faturamento: Loja {faturamento_lojas_dia.index[0]} com faturamento R${faturamento_lojas_dia.iloc[0, 0]:.2f}
Pior Loja do Dia em Faturamento: Loja {faturamento_lojas_dia.index[-1]} com faturamento R${faturamento_lojas_dia.iloc[-1, 0]:.2f}

Melhor Loja do Ano em Faturamento: Loja {faturamento_lojas_ano.index[0]} com faturamento R${faturamento_lojas_ano.iloc[0, 0]:.2f}
Pior Loja do Ano em Faturamento: Loja {faturamento_lojas_ano.index[-1]} com faturamento R${faturamento_lojas_ano.iloc[-1, 0]:.2f}

Segue em anexo os rankings do ano e do dia de todas as lojas.

Qualquer dúvida estou à disposição.

Att., 
Maykon
'''
email_msg = MIMEMultipart()
email_msg['From'] = login
email_msg['To'] = emails.loc[emails['Loja']=='Diretoria', 'E-mail'].values[0]
email_msg['Subject'] = f'Ranking Dia {dia_indicador.day}/{dia_indicador.month}'
email_msg.attach(MIMEText(corpo,'plain'))#serve para anexar o corpo no email.

#3- ENVIANDO ANEXO / Abrindo o arquivo em modo leitura e binary
cam_arquivo = pathlib.Path.cwd() / caminho_backup / f'{dia_indicador.month}_{dia_indicador.day}_Ranking Dia.xlsx'
attachment = open(cam_arquivo,'rb')#rb = READ BINARY / LEITURA BINARIA
#Lemos o arquivo no modo binario e jogamos codificado em base 64 (que é oque o email precisa)
att = MIMEBase('application', 'octet-stream')
att.set_payload(attachment.read())
encoders.encode_base64(att)
#Adicionamos o cabeçalho no tipo anexo de email
att.add_header('Content-Disposition', f'attachment; filename={dia_indicador.month}_{dia_indicador.day}_Ranking Dia.xlsx')
#fechamos o arquivo
attachment.close()
#colocamos o anexo no corpo do email
email_msg.attach(att)

#3- ENVIANDO ANEXO / Abrindo o arquivo em modo leitura e binary
cam_arquivo = pathlib.Path.cwd() / caminho_backup / f'{dia_indicador.month}_{dia_indicador.day}_Ranking Anual.xlsx'
attachment = open(cam_arquivo,'rb')#rb = READ BINARY / LEITURA BINARIA
#Lemos o arquivo no modo binario e jogamos codificado em base 64 (que é oque o email precisa)
att = MIMEBase('application', 'octet-stream')
att.set_payload(attachment.read())
encoders.encode_base64(att)
#Adicionamos o cabeçalho no tipo anexo de email
att.add_header('Content-Disposition', f'attachment; filename={dia_indicador.month}_{dia_indicador.day}_Ranking Anual.xlsx')
#fechamos o arquivo
attachment.close()
#colocamos o anexo no corpo do email
email_msg.attach(att)

#4- ENVIAR O EMAIL TIPO MIME NO SERVIDOR SMTP
server.sendmail(email_msg['From'],email_msg['To'],email_msg.as_string())

server.quit()
print(f'E-mail da Diretoria Enviado.')