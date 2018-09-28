' adPasswordReminder.vbs sends alerts about ad account password expiration.
' Copyright (C) 2018  Ramón Román Castro <ramonromancastro@gmail.com>

' This program is free software; you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation; either version 2 of the License, or
' (at your option) any later version.

' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.

' You should have received a copy of the GNU General Public License along
' with this program; if not, write to the Free Software Foundation, Inc.,
' 51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.

On Error Resume Next

'-----------------------------------------------------------------------------------------
' CONFIGURACION (Estas son las opciones personalizables del script)
'-----------------------------------------------------------------------------------------

Const FIRST_ADVICE	= 15
Const SECOND_ADVICE	= 10
Const NEXT_ADVICE	= 5

Const DOMAIN_NAME	= "DC=domain,DC=local"
Const DOMAIN_SERVER = "server.domain.local"
Const DOMAIN_DN		= "CN=Users,DC=domain,DC=local"

Const MAIL_SERVER	= "mail.domain.local"
Const MAIL_PORT		= 25
Const MAIL_USERNAME	= "example"
Const MAIL_PASSWORD	= "P@$$w0rd"
Const MAIL_FROM		= "example@domain.local"
Const MAIL_SUBJECT	= "Notificación de expiración de contraseña"
Const MAIL_TIMEOUT	= 60
Const MAIL_SSL		= False

Const EMAIL_TEMPLATE= ".\adPasswordReminder.tpl.html"

'-----------------------------------------------------------------------------------------
' CONSTANTES
'-----------------------------------------------------------------------------------------

Const ADS_UF_DONT_EXPIRE_PASSWD = &h10000
Const E_ADS_PROPERTY_NOT_FOUND  = &h8000500D
Const ONE_HUNDRED_NANOSECOND    = .000000100
Const SECONDS_IN_DAY            = 86400

Const cdoSendUsingPort	= 2
Const cdoBasic			= 1

Const TAG_USERNAME	= "<!--USERNAME-->"
Const TAG_NAME		= "<!--NAME-->"	
Const TAG_DAYS_LEFT	= "<!--DAYS_LEFT-->"
Const TAG_DATETIME	= "<!--DATETIME-->"

'-----------------------------------------------------------------------------------------
' FUNCIONES
'-----------------------------------------------------------------------------------------

Function loadTemplate
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objTextFile = objFSO.OpenTextFile(EMAIL_TEMPLATE, 1)
	strText = objTextFile.ReadAll
	objTextFile.Close
	loadTemplate = strText
End Function

Sub sendMailCDOSYS(sFrom, sTO, sSubject, sMailBody)
	On Error Resume Next
	
	Dim objCDOConf, objCDOSYS

	Set objCDOSYS = WScript.CreateObject("CDO.Message")
	Set objCDOConf = WScript.CreateObject("CDO.Configuration")

	With objCDOConf
		.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = MAIL_SERVER
		.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = MAIL_PORT
		.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPort
		.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
		.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = MAIL_USERNAME
		.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = MAIL_PASSWORD
		.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = MAIL_TIMEOUT
		.Fields("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = MAIL_SSL
		.Fields.Update 
	End With

	Set objCDOSYS.Configuration = objCDOConf

	With objCDOSYS     
		.From = sFrom
		.To = sTo
		.Subject = sSubject
		.TextBody = sMailBody
		.HTMLBody = sMailBody
		.Send
	End With
	
	Set objCDOSYS = Nothing
	
	On Error GoTo 0
End Sub

Function printf(S,Args)
    Dim I
    printf = S
    For I = LBound(Args) To UBound(Args)
        If InStr(printf, "%s") <> 0 Then
           printf = Replace(printf, "%s", Args(I), 1, 1)
        End If
    Next
End Function

Sub sendAdvice(mail)
	On Error Resume Next
	email = loadTemplate

	email = Replace(email,TAG_USERNAME,oADRecordSet.Fields("samAccountName").Value)
	email = Replace(email,TAG_NAME,oADRecordSet.Fields("displayName").Value)
	email = Replace(email,TAG_DAYS_LEFT,Int(dblMaxPwdDays - intTimeInterval))
	email = Replace(email,TAG_DATETIME,Date&" "&Time)

	sendMailCDOSYS	MAIL_FROM,_
					mail,_
					MAIL_SUBJECT,_
					email
	On Error GoTo 0
End Sub

'-----------------------------------------------------------------------------------------
' INICIALIZANDO...
'-----------------------------------------------------------------------------------------

Set oADConnection	= CreateObject("ADODB.Connection")
Set oADCommand		= CreateObject("ADODB.Command")
Set objDomain		= GetObject("LDAP://" & DOMAIN_NAME)
Set objMaxPwdAge	= objDomain.Get("maxPwdAge")

'-----------------------------------------------------------------------------------------
' CODIGO PRINCIPAL
'-----------------------------------------------------------------------------------------

If objMaxPwdAge.LowPart = 0 Then
	WScript.Quit
Else
	dblMaxPwdNano = Abs(objMaxPwdAge.HighPart * 2^32 + objMaxPwdAge.LowPart)
	dblMaxPwdSecs = dblMaxPwdNano * ONE_HUNDRED_NANOSECOND
	dblMaxPwdDays = Int(dblMaxPwdSecs / SECONDS_IN_DAY)
End If

oADConnection.Provider = "ADsDSOObject"
oADConnection.Open "Active Directory Provider"
Set oADCommand.ActiveConnection = oADConnection

oADCommand.CommandText = "<LDAP://" & DOMAIN_SERVER & "/" & DOMAIN_DN & ">;(&(objectCategory=User)(!(userAccountControl:1.2.840.113556.1.4.803:=2)));displayName,cn,mail,company,physicalDeliveryOfficeName,department,samAccountName,userAccountControl,ADsPath;subtree"
oADCommand.Properties("Sort On") = "cn"
Set oADRecordSet = oADCommand.Execute

Do While NOT oADRecordSet.Eof

	Set objUser = GetObject(oADRecordSet.Fields("AdsPath").Value)
	intUserAccountControl = objUser.Get("userAccountControl")
	
	If (NOT intUserAccountControl AND ADS_UF_DONT_EXPIRE_PASSWD) AND (Len(oADRecordSet.Fields("mail").Value) > 0) Then
		dtmValue = objUser.PasswordLastChanged
		If Err.Number = E_ADS_PROPERTY_NOT_FOUND Then
			Err.Clear
		Else
			intTimeInterval = Int(Now - dtmValue)
			If (dblMaxPwdDays - intTimeInterval = FIRST_ADVICE) Then
				sendAdvice oADRecordSet.Fields("mail").Value
			Else
				if (dblMaxPwdDays - intTimeInterval = SECOND_ADVICE) Then
					sendAdvice oADRecordSet.Fields("mail").Value
				Else
					if (dblMaxPwdDays - intTimeInterval <= NEXT_ADVICE) AND (dblMaxPwdDays - intTimeInterval >= 0) Then
						sendAdvice oADRecordSet.Fields("mail").Value
					End If
				End If
			End If
		End If
	End If
	oADRecordSet.MoveNext
Loop

On Error GoTo 0