from email.message import EmailMessage
import smtplib
import openpyxl

excel_dataframe = openpyxl.load_workbook('pruebaPy.xlsx')
df = excel_dataframe.active

def onlyName(name):
    nombre= ''
    for i in name:
        nombre = nombre + i
        if i == ' ':
            return nombre

ocho=[]
ochoO = []
doce=[]
doceO=[]
psa=[]
 
for row in df.iter_rows(min_row=2, values_only=True):
    if row[3]==8:
        ocho.append(row)
    if row[3]==12:
        doce.append(row)
    if row[3]==120:
        doceO.append(row)

for row in ocho:
    destinatario = row[1]
    nombre = onlyName(row[0])
    fecha = str(row[4])
    asunto = 'Indicaciones turno laboratorio '+ fecha
    firma = 'https://imgur.com/BXGZMDT'

    if destinatario != '':
        remitente = "laboratoriosmcdla@swissmedical.com.ar"

        password = "septiembre2023"

        mensaje = 'Buenos días, ' + nombre + ', \n \n Le confirmamos el turno de laboratorio para el día ' + fecha + '. \n De acuerdo a la orden adjunta al momento de tomar el turno, la indicación para la correcta realización del estudio es:\n\n -  8 horas de ayuno (puede tomar agua)\n\n Saludos cordiales. \n\n'+ firma + ' Mauro Sanzberro\nRecepcionista\nSMCDLA Laboratorio\nGuatemala 5455, C.A.B.A.\nTel.: 4778-4650'

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


    
