import win32com.client

def outlookOpenConnection():
    return win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

def outlookOpenConnectionToEmail():
    return win32com.client.Dispatch("Outlook.Application")  
    
#Cierra el subproceso de outlook, destruye el objeto tipo Outlook
def outlookCloseConnection(objOutlook):
    objOutlook.Quit
    objOutlook = Nothing

#Retorna la capeta indicada por la constante SFolder, Nota: Por defecto InboxFolder tiene como valor la constante entera: 6
def outlookSetFolder(objOutlook, sFolder):  
	return objOutlook.GetDefaultFolder(sFolder)

#Parametro sEmailUnreadStatusToFind: "True" para emails no leidos, filtra por propiedad del Email (sEmailPropertyToVerify) Valor: sEMailPropertyValue - sMAxAttemptToFindEmail, numero de intentos para filtrar Email #Retorna BOOLEANO: TRUE o FALSE
#Ej: OutlookCheckUnreadEmails("SUBJECT","Hallazgos de Automatizacion Reportes",3,"True")
def outlookCheckUnreadEmails(sEmailPropertyToVerify, sEmailPropertyValue, sMaxAttemptToFindEmail, sEmailUnreadStatusToFind):
	
	for j in range(0,sMaxAttemptToFindEmail):
		oFilteredEmails = outlookFilterEmail("Unread", sEmailUnreadStatusToFind)
		sFilteredEmailCount = oFilteredEmails.Count
		sEmailActualPropertyValue= ""
		result=False
		for sFilteredEmail in oFilteredEmails:
			sEmailActualPropertyValue=getTheEmailUsingProperty(sFilteredEmail,sEmailPropertyToVerify)
			if sEmailPropertyValue in sEmailActualPropertyValue: return True
			else: result = False
	return result

#Filtra email por propiedad actual del mail y valor de la propiedad actual, doc: https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.mailitem_properties.aspx
#Ej: Set EMail=OutlookFilterEmail("Unread", "True") / Set Email=OutlookFilterEmail("Subject", "Hallazgos de Automatizacion Reportes")
def outlookFilterEmail(sEmailProperty, sEmailPropertyValue):
	
	objOutlook = outlookOpenConnection()
	objFolder=outlookSetFolder(objOutlook,6)
	objAllMails = objFolder.Items
	objFilterEmail = objAllMails.Restrict("[" + sEmailProperty + "] = " + sEmailPropertyValue )
	return objFilterEmail

#Filtra emails por por alguna de sus propiedades, retorna la propiedad del elemento exacta, Filtra la propiedad por valores inexactos y la propiedad exacta
def outlookFilterEmailsByUnexactlyProperty(sEmailPropertyToVerify, sEmailPropertyValue, sMaxAttemptToFindEmail, sEmailUnreadStatusToFind):

	for j in range(0,sMaxAttemptToFindEmail):
		oFilteredEmails = outlookFilterEmail("Unread", sEmailUnreadStatusToFind)
		for sFilteredEmail in oFilteredEmails:
			sEmailActualPropertyValue=getTheEmailUsingProperty(sFilteredEmail,sEmailPropertyToVerify)
			if sEmailPropertyValue in sEmailActualPropertyValue: return sEmailActualPropertyValue
	return 'Not found'

def outlookGetEmailProperty(sEmailPropertyToVerify, sEmailPropertyValue, sMaxAttemptToFindEmail, sEmailUnreadStatusToFind):        
	sActualResult = ""

	def getEmailAddress():
		
		try:
			return sFilteredEmail.Sender.GetExchangeUser().PrimarySmtpAddress
		except:
			return getattr(sFilteredEmail, 'SenderEmailAddress', '<UNKNOWN>')

	for j in range(0,sMaxAttemptToFindEmail):
		oFilteredEmails = outlookFilterEmail("Unread", sEmailUnreadStatusToFind)
		emails=[]
		emailsToRead=[]
		sEmailActualPropertyValue= None
		
		for sFilteredEmail in oFilteredEmails:
			sEmailActualPropertyValue=getTheEmailUsingProperty(sFilteredEmail,sEmailPropertyToVerify)

			if sEmailPropertyValue in str(sEmailActualPropertyValue):
				EmailProperties={
				"SUBJECT": getattr(sFilteredEmail, 'Subject', '<UNKNOWN>'),
				"SENDER NAME": getattr(sFilteredEmail, 'SenderName', '<UNKNOWN>'),
				"SENDER EMAIL ADDRESS": getEmailAddress(), #getattr(sFilteredEmail, 'SenderEmailAddress', '<UNKNOWN>'),
				"SENDER EMAIL TYPE": getattr(sFilteredEmail, 'SenderEmailType', '<UNKNOWN>'),
				#MailItem.Sender.GetExchnageUser().ProimarySmtpAddress
				"TO": getattr(sFilteredEmail, 'To', '<UNKNOWN>'),
				"CC": getattr(sFilteredEmail, 'CC', '<UNKNOWN>'),
				"BCC": getattr(sFilteredEmail, 'BCC', '<UNKNOWN>'),
				"BODY": getattr(sFilteredEmail, 'BODY', '<UNKNOWN>'),
				"CREATION TIME": getattr(sFilteredEmail, 'CreationTime', '<UNKNOWN>'),
				"IMPORTANCE": getattr(sFilteredEmail, 'Importance', '<UNKNOWN>'),
				"RECEIVED TIME": getattr(sFilteredEmail, 'ReceivedTime', '<UNKNOWN>'),
				}
				emails.append({key: str(value).encode('utf-8') for key, value in EmailProperties.items()})
				emailsToRead.append(sFilteredEmail)

	for emailToRead in emailsToRead:
		#TODO: descomentar para setear correos en readed
		setattr(emailToRead,'Unread','False')
		pass
	return emails


def outlookSendEmails(objOutlook, mailTo, mailSubject, mailBody, mailHTMLBody):
    mail = objOutlook.CreateItem(0)
    mail.To = mailTo
    mail.Subject = mailSubject
    mail.Body = mailBody
    mail.HTMLBody = mailHTMLBody# this field is optional
    
    # attachment  = "Path to the attachment"
    # mail.Attachments.Add(attachment)

    mail.Send()

def getTheEmailUsingProperty(sFilteredEmail,sEmailPropertyToVerify):
		if sEmailPropertyToVerify=="SUBJECT": sEmailActualPropertyValue=getattr(sFilteredEmail, 'Subject', '<UNKNOWN>')
		elif sEmailPropertyToVerify=="SENDER NAME": sEmailActualPropertyValue = getattr(sFilteredEmail, 'SenderName', '<UNKNOWN>')
		elif sEmailPropertyToVerify=="SENDER EMAIL ADDRESS": sEmailActualPropertyValue = getattr(sFilteredEmail, 'SenderEmailAddress', '<UNKNOWN>')
		elif sEmailPropertyToVerify=="SENDER EMAIL TYPE": sEmailActualPropertyValue = getattr(sFilteredEmail, 'SenderEmailType', '<UNKNOWN>')
		elif sEmailPropertyToVerify=="TO": sEmailActualPropertyValue = getattr(sFilteredEmail, 'To', '<UNKNOWN>')
		elif sEmailPropertyToVerify=="CC": sEmailActualPropertyValue = getattr(sFilteredEmail, 'CC', '<UNKNOWN>')
		elif sEmailPropertyToVerify=="BCC": sEmailActualPropertyValue = getattr(sFilteredEmail, 'BCC', '<UNKNOWN>')
		elif sEmailPropertyToVerify=="BODY": sEmailActualPropertyValue = getattr(sFilteredEmail, 'BODY', '<UNKNOWN>')
		elif sEmailPropertyToVerify=="CREATION TIME": sEmailActualPropertyValue = getattr(sFilteredEmail, 'CreationTime', '<UNKNOWN>')
		elif sEmailPropertyToVerify=="IMPORTANCE": sEmailActualPropertyValue = getattr(sFilteredEmail, 'Importance', '<UNKNOWN>')
		elif sEmailPropertyToVerify=="RECEIVED TIME": sEmailActualPropertyValue = getattr(sFilteredEmail, 'ReceivedTime', '<UNKNOWN>')
		else: sEmailActualPropertyValue=None
		if sEmailActualPropertyValue: return sEmailActualPropertyValue			