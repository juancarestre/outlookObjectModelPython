#coding=utf-8
import win32com.client
#******************************************** FUNCTION ****************************************************************
#Crea el objeto tipo Outlook
#Lo retorna como OutlookOpenConnection
def outlookOpenConnection():
    return win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    # return win32com.client.Dispatch("Outlook.Application")    

def outlookOpenConnectionToEmail():
    return win32com.client.Dispatch("Outlook.Application")  
    
#**********************************************************************************************************************
#******************************************** FUNCTION ****************************************************************
#Cierra el subproceso de outlook, destruye el objeto tipo Outlook
def outlookCloseConnection(objOutlook):
    objOutlook.Quit
    objOutlook = Nothing
#**********************************************************************************************************************
#******************************************** FUNCTION ****************************************************************
#Selecciona carpeta de trabajo dentro del Outlook, las carpetas se identifican por una Constante, 
#se ingresan como parametro SFolder, dentro de la funcion se encuentran
#Especificadas las demas cosntantes de Folders disponibles
#Nota: Por defecto InboxFolder tiene como valor la constante entera: 6
#Retorna la capeta indicada por la constante SFolder
def outlookSetFolder(objOutlook, sFolder):  
    #	Name	Value	Description
		#olFolderCalendar	9	The Calendar folder.
		#olFolderConflicts	19	The Conflicts folder (subfolder of the Sync Issues folder). Only available for an Exchange account.
		#olFolderContacts	10	The Contacts folder.
		#olFolderDeletedItems	3	The Deleted Items folder.
		#olFolderDrafts	16	The Drafts folder.
		#olFolderInbox	6	The Inbox folder.
		#olFolderJournal	11	The Journal folder.
		#olFolderJunk	23	The Junk E-Mail folder.
		#olFolderLocalFailures	21	The Local Failures folder (subfolder of the Sync Issues folder). Only available for an Exchange account.
		#olFolderManagedEmail	29	The top-level folder in the Managed Folders group. For more information on Managed Folders, see the Help in Microsoft Outlook. Only available for an Exchange account.
		#olFolderNotes	12	The Notes folder.
		#olFolderOutbox	4	The Outbox folder.
		#olFolderSentMail	5	The Sent Mail folder.
		#olFolderServerFailures	22	The Server Failures folder (subfolder of the Sync Issues folder). Only available for an Exchange account.
		#olFolderSuggestedContacts	30	The Suggested Contacts folder.
		#olFolderSyncIssues	20	The Sync Issues folder. Only available for an Exchange account.
		#olFolderTasks	13	The Tasks folder.
		#olFolderToDo	28	The To Do folder.
		#olPublicFoldersAllPublicFolders	18	The All Public Folders folder in the Exchange Public Folders store. Only available for an Exchange account.
		#olFolderRssFeeds	25	The RSS Feeds folder.
	return objOutlook.GetDefaultFolder(sFolder)
#**********************************************************************************************************************
#******************************************** FUNCTION ****************************************************************
#Filtra Emails por Leidos o no Leidos (Parametro sEmailUnreadStatusToFind: "True" para emails no leidos
#Subfiltra por propiedad del Email (Parametro sEmailPropertyToVerify), las propiedades para filtrar son los emails son:
#SUBJECT, SENDER NAME, SENDER EMAIL ADDRESS, SENDER EMAIL TYPE, TO, CC, BCC, BODY, CREATION TIME, IMPORTANCE, RECEIVED TIME
#Valor de la propiedad a filtrar Parametro: sEMailPropertyValue
#Parametro sMAxAttemptToFindEmail, numero de intentos para filtrar Email, ingresa como numero entero
#Retorna BOOLEANO: TRUE o FALSE, en caso de que el email haya sido encontrado
#Ej: OutlookCheckUnreadEmails("SUBJECT","Hallazgos de Automatizacion Reportes",3,"True")
def outlookCheckUnreadEmails(sEmailPropertyToVerify, sEmailPropertyValue, sMaxAttemptToFindEmail, sEmailUnreadStatusToFind):
	
	sActualResult = ""
	
	for j in range(0,sMaxAttemptToFindEmail):
            oFilteredEmails = outlookFilterEmail("Unread", sEmailUnreadStatusToFind)
            sFilteredEmailCount = oFilteredEmails.Count
            sEmailActualPropertyValue= ""
            result=False
            for sFilteredEmail in oFilteredEmails:
            
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
                else: print("Not Found")
            
                if sEmailPropertyValue in sEmailActualPropertyValue:
                    sActualResult="True"
                    sEmailActualPropertyValue=""
                    return True
                else: 
                    result = False
        return result

		
#**********************************************************************************************************************
#******************************************** FUNCTION ****************************************************************
#Filtra email por propiedad actual del mail y valor de la propiedad actual
#Ej: Set EMail=OutlookFilterEmail("Unread", "True")
#Ej: Set Email=OutlookFilterEmail("Subject", "Hallazgos de Automatizacion Reportes")
#Propiedades de los emails en outlook: https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.mailitem_properties.aspx
#El valor de la propiedad debe ser exacto para poder filtrar el mensaje
def outlookFilterEmail(sEmailProperty, sEmailPropertyValue):
	
	objOutlook = outlookOpenConnection()
	objFolder=outlookSetFolder(objOutlook,6)
	#Find all items in the Folder
	objAllMails = objFolder.Items
	objFilterEmail = objAllMails.Restrict("[" + sEmailProperty + "] = " + sEmailPropertyValue )
    # objFilterEmail.GetNext()
	return objFilterEmail
#**********************************************************************************************************************


#******************************************** FUNCTION ****************************************************************
#Filtra emails por por alguna de sus propiedades
#Retorna la propiedad del elemento exacta
#Filtra la propiedad por valores inexactos y la propiedad exacta
def outlookFilterEmailsByUnexactlyProperty(sEmailPropertyToVerify, sEmailPropertyValue, sMaxAttemptToFindEmail, sEmailUnreadStatusToFind):
	
	sActualResult = ""
	
	for j in range(0,sMaxAttemptToFindEmail):
            oFilteredEmails = outlookFilterEmail("Unread", sEmailUnreadStatusToFind)
            # setattr(oFilteredEmails,'Unread','<UNKNOWN>')
            # oFilteredEmails.save()
            sFilteredEmailCount = oFilteredEmails.Count
            sEmailActualPropertyValue= ""
            result=False
            for sFilteredEmail in oFilteredEmails:
                # setattr(sFilteredEmail,'Unread','False')
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
                else: a=0
            
                if sEmailPropertyValue in sEmailActualPropertyValue:
                    sActualResult="True"
                    # print (( str (sEmailActualPropertyValue) ))
                    return (sEmailActualPropertyValue)
                else: 
                    result = 'Not found'
        return result


def outlookSendEmails(objOutlook, mailTo, mailSubject, mailBody, mailHTMLBody):
    mail = objOutlook.CreateItem(0)
    mail.To = mailTo
    mail.Subject = mailSubject
    mail.Body = mailBody
    mail.HTMLBody = mailHTMLBody# this field is optional
    
    #In case you want to attach a file to the email
    # attachment  = "Path to the attachment"
    # mail.Attachments.Add(attachment)

    mail.Send()