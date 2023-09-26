from email.message import EmailMessage
import smtplib
import openpyxl

excel_dataframe = openpyxl.load_workbook('pruebaPy.xlsx')
df = excel_dataframe.active

ocho=[]
doce=[]
doceO=[]
 
for row in df.iter_rows(min_row=2, values_only=True):
    if row[3]==8:
        ocho.append(row)
    if row[3]==12:
        doce.append(row)
    if row[3]==120:
        doceO.append(row)

for row in ocho:
    destinatario = row[1]
    nombre = row[0]
    asunto = row[2]

    if destinatario != '':
        remitente = "laboratoriosmcdla@swissmedical.com.ar"

        password = "septiembre2023"

        mensaje = 'Hola ' + nombre 

        email = EmailMessage()
        email["from"] = remitente
        email["to"] = destinatario
        email["subject"] = asunto

        email.set_content(mensaje)
                
        smtp = smtplib.SMTP("smtp-mail.outlook.com", port = 587 )

        smtp.ehlo()

        smtp.starttls()

        smtp.login(remitente,password)

        smtp.sendmail(remitente, destinatario, email.as_string())

        smtp.quit()


    else :
        print 
        "Error"

for row in doce:
    destinatario = row[1]
    nombre = row[0]
    asunto = row[2]

    if destinatario != '':
        remitente = "laboratoriosmcdla@swissmedical.com.ar"

        password = "septiembre2023"

        mensaje = 'Hola ' + nombre 

        email = EmailMessage()
        email["from"] = remitente
        email["to"] = destinatario
        email["subject"] = asunto

        email.set_content(mensaje)
                
        smtp = smtplib.SMTP("smtp-mail.outlook.com", port = 587 )

        smtp.ehlo()

        smtp.starttls()

        smtp.login(remitente,password)

        smtp.sendmail(remitente, destinatario, email.as_string())

        smtp.quit()


    else :
        print 
        "Error"

for row in doceO:
    destinatario = row[1]
    nombre = row[0]
    asunto = row[2]

    if destinatario != '':
        remitente = "laboratoriosmcdla@swissmedical.com.ar"

        password = "septiembre2023"

        mensaje = 'Hola ' + nombre 

        email = EmailMessage()
        email["from"] = remitente
        email["to"] = destinatario
        email["subject"] = asunto

        email.set_content(mensaje)

        with open("PSA.pdf", "rb") as f:
            email.add_attachment(
                f.read(),
                filename="PSA.pdf",
                maintype="application",
                subtype="pdf"
            )

                
        smtp = smtplib.SMTP("smtp-mail.outlook.com", port = 587 )

        smtp.ehlo()

        smtp.starttls()

        smtp.login(remitente,password)

        smtp.sendmail(remitente, destinatario, email.as_string())

        smtp.quit()


    else :
        print 
        "Error"


    
