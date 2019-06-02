Option Explicit
On Error Resume Next

Const ImgBurnDirectory = "C:\Program Files (x86)\ImgBurn\"
Const GoogleDriveCopyFileDirectory = ".\"

Const defSendEmail = False

Const mailPlain = 0
Const mailHTML = 1 

Dim company
Dim backupSourceInfo
Dim backupDest
Dim sendEmailArg
Dim emailTo
Dim emailFrom
Dim emailCC
Dim emailReplyTo

If Not IsCScript() Then
  Wscript.Echo WScript.ScriptName & " must be run with CScript."
  Wscript.Quit 1
End If

WScript.Quit ExecuteBackupProcess()

'***************************************************************************************

Function ExecuteBackupProcess()
  On Error Resume Next
  Dim ErrMsg
  ExecuteBackupProcess = 0
  ErrMsg = ArgumentsOK()
  If ErrMsg <> "" then
    Call ShowUsage(ErrMsg)
  Else
    Call PerformBackup()
    if Err.Number = 0 Then
      Call ProcessSuccess()
    end if
    if Err.Number <> 0 Then
      ExecuteBackupProcess = 1
      ProcessError(Err)
      If Err.Number <> 0 Then
        WScript.Echo "Error in Processing Error Routine: " & Err.Number & " " & Err.Description
      End if
    end if
  End If
End Function


'***************************************************************************************

Sub ShowUsage(ByVal ErrMsg)
  If ErrMsg <> "" Then
    WScript.Echo "ERROR: " & ErrMsg
  End if
  WScript.Echo "USAGE: cscript " & WScript.ScriptName & vbNewLine & _
               "(/BackupID: /BackupSource: /BackupDest: [/SendEmail:true|FALSE]) | " & vbNewLine & _
               "(/UseConfigFile:)"
end Sub

'***************************************************************************************

Function ArgumentsOK()
  If WScript.Arguments.Named.Exists("UseConfigFile") Then
    ArgumentsOK = GetArgumentsFromConfigFile()
  Else
    ArgumentsOK = GetArgumentsFromCommandLine()
  End If
End Function

'***************************************************************************************

Function GetArgumentsFromCommandLine()
  Dim validOptions
  Dim arg

  validOptions = Array("BackupID", "BackupSource", "BackupDest", "SendEmail", _
					   "EmailTo", "EmailFrom", "EmailCC", "EmailReplyTo")
  arg = NamedArgsOK(validOptions)
  If arg <> "" Then
    GetArgumentsFromCommandLine = "Illegal Command Line Argument: " & arg
    Exit Function
  End If

  company          = Trim(WScript.Arguments.Named("BackupID"))
  backupSourceInfo = Trim(WScript.Arguments.Named("BackupSource"))
  backupDest       = Trim(WScript.Arguments.Named("BackupDest"))
  sendEmailArg     = Trim(WScript.Arguments.Named("SendEmail"))
  emailTo          = Trim(WScript.Arguments.Named("EmailTo"))
  emailFrom        = Trim(WScript.Arguments.Named("EmailFrom"))
  emailCC          = Trim(WScript.Arguments.Named("EmailCC"))
  emailReplyTo     = Trim(WScript.Arguments.Named("EmailReplyTo"))

  GetArgumentsFromCommandLine = VerifyParameters()
End Function

'***************************************************************************************

Function GetArgumentsFromConfigFile()
  Dim validOptions
  Dim xmlDoc
  Dim arg

  validOptions = Array("UseConfigFile")
  arg = NamedArgsOK(validOptions)
  If arg <> "" Then
    GetArgumentsFromConfigFile = "Illegal Command Line Argument: " & arg
    Exit Function
  End If

  arg = WScript.Arguments.Named("UseConfigFile")
  If Not FileExists(arg) Then
    GetArgumentsFromConfigFile = "Config File Not Found: " & arg
    Exit Function
  End If
  
  Set xmlDoc = WScript.CreateObject("Microsoft.XMLDOM")
  xmlDoc.async = False
  xmlDoc.load(arg)
  if xmlDoc.parseError.errorCode <> 0 Then
    GetArgumentsFromConfigFile = "XML Parse Error " & xmlDoc.parseError.reason & " at Line " & xmlDoc.parseError.line
  Else
    company          = getXMLValue(xmlDoc, "BackupID")
    backupSourceInfo = getXMLValue(xmlDoc, "BackupSource/Location")
    backupDest       = getXMLValue(xmlDoc, "BackupDest")
    sendEmailArg     = getXMLValue(xmlDoc, "SendEmail")
    emailTo          = getXMLValue(xmlDoc, "Email/To")
    emailFrom        = getXMLValue(xmlDoc, "Email/From")
    emailCC          = getXMLValue(xmlDoc, "Email/CC")
    emailReplyTo     = getXMLValue(xmlDoc, "Email/ReplyTo")

    GetArgumentsFromConfigFile = VerifyParameters()
  End if
  Set xmlDoc = Nothing
End Function

'***************************************************************************************

Function VerifyParameters()
  Dim i
  If TypeName(BackupSourceInfo) = "String" then
    backupSourceInfo = Split(BackupSourceInfo, ";")
    For i = 0 To UBound(backupSourceInfo)
      backupSourceInfo(i) = Split(backupSourceInfo(i), "=")
      if UBound(backupSourceInfo(i)) <> 1 Then
        VerifyParameters = "Illegally Formed Backup Source Specification"
        Exit Function
      End If
      backupSourceInfo(i)(0) = Trim(backupSourceInfo(i)(0))
      backupSourceInfo(i)(1) = Trim(backupSourceInfo(i)(1))
      If backupSourceInfo(i)(0) = "" Or backupSourceInfo(i)(1) = "" Then
        VerifyParameters = "Illegally Formed Backup Source Specification"
        Exit Function
      End if
    Next
  ElseIf TypeName(BackupSourceInfo(0)) = "String" Then
    backupSourceInfo = Array(backupSourceInfo)
  End If

  If Not IDOK(company) Then
    VerifyParameters = "BackupID Must Be A Maximum of 5 Alphanumeric Characters" & " " & company
    Exit Function
  End If

  If BackupDest = "" Then
    VerifyParameters = "BackupDest Not Specified"
    Exit Function
  End If
  If Not BackupDestExists(BackupDest) Then
    VerifyParameters = "Illegal Backup Destination Specified"
    Exit Function
  End If
  If Right(BackupDest, 1) <> "\" Then
    BackupDest = BackupDest & "\"
  End If
  
  sendEmailArg = UCase(sendEmailArg)
  If sendEmailArg = "" Then
    sendEmailArg = defSendEmail
  ElseIf sendEmailArg = "FALSE" Then
    sendEmailArg = False
  ElseIf sendEmailArg = "TRUE" Then
    sendEmailArg = True
  Else
    VerifyParameters = "SendEmail parameter must be TRUE or FALSE"
    Exit Function
  End If

  If sendEmailArg Then
    If Not EmailOK(emailTo) Then
      VerifyParameters = "Email TO must be legal email: " & emailTo
      Exit Function
    End If
    If Not EmailOK(emailFrom) Then
      VerifyParameters = "Email FROM must be legal email: " & emailFrom
      Exit Function
    End If
    If emailCC <> "" Then
      If Not EmailOK(emailCC) Then
        VerifyParameters = "Email CC must be legal email: " & emailCC
        Exit Function
      End If
    End if
    If emailReplyTo <> "" Then
      If Not EmailOK(emailReplyTo) Then
        VerifyParameters = "Email ReplyTo must be legal email: " & emailReplyTo
        Exit Function
      End If
    End if
  End if
End Function

'***************************************************************************************

Sub ShowParameters()
  Dim i
  WScript.Echo "Backup ID:  " & company
  WScript.Echo "Dest:       " & backupDest
  WScript.Echo "Send Email: " & sendEmailArg
  WScript.Echo "Email To:   " & emailTo
  WScript.Echo "Email From: " & emailFrom
  WScript.Echo "Email CC:   " & emailCC
  WScript.Echo "Email Reply:" & emailReplyTo
  WScript.Echo "Source:     "

  For i = 0 To UBound(backupSourceInfo)
    WScript.Echo backupSourceInfo(i)(0) & " " & backupSourceInfo(i)(1)
  Next
End Sub

'***************************************************************************************

Function getXMLValue(ByRef xmlDoc, ByVal tag)
  Dim NodeList, arr(), i
  If tag = "" Then
    Set NodeList = xmlDoc.getElementsByTagName("*")
    If NodeList.Length = 0 Then
      getXMLValue = Trim(xmlDoc.text)  '.childNodes(0).nodeValue
    Else
      ReDim arr(NodeList.Length - 1)
      For i = 0 To UBound(arr)
        arr(i) = getXMLValue(NodeList(i), "")
      Next
      getXMLValue = arr
    End if
  Else
    Set NodeList = xmlDoc.getElementsByTagName(tag)
    If NodeList.Length = 0 Then
      getXMLValue = ""
    ElseIf NodeList.Length = 1 Then
      getXMLValue = getXMLValue(NodeList(0), "")
    Else
      ReDim arr(NodeList.Length - 1)
      For i = 0 To UBound(arr)
        arr(i) = getXMLValue(NodeList(i), "")
      Next
      getXMLValue = arr
    End If
  End if
  Set NodeList = Nothing
End Function

'***************************************************************************************

Function NamedArgsOK(ByRef validOptions)
  Dim arg, i, j
  Dim rx, matches

  Set rx = New regexp
  rx.Pattern = "/([a-z]+):?"
  rx.Global = True
  rx.IgnoreCase = True
  For i = 0 To WScript.Arguments.Count - 1
    arg = WScript.Arguments(i)
    Set matches = rx.execute(arg)
    If matches.count = 1 Then
       arg = UCase(matches(0).submatches(0))
       For j = 0 To UBound(validOptions)
         If arg = UCase(validOptions(j)) Then
           Exit For
         End If
       Next
       If j > UBound(validOptions) Then
         NamedArgsOK = arg
         Exit Function
       End if
    End If
    Set matches = Nothing
  Next
  Set rx = Nothing
End Function

'***************************************************************************************

Sub ProcessSuccess()
  Dim Msg
  Dim WshShell
  Msg = "SUCCESS in " & company & " Backup!!!"

  WScript.Echo Msg

  Set WshShell = WScript.CreateObject("WScript.Shell")
  WshShell.LogEvent 4, Msg
  Set WshShell = Nothing

  Call SendEmail(emailFrom, emailTo, emailCC, emailReplyTo, Msg, Msg)
End Sub


'***************************************************************************************

Sub ProcessError(ByRef Err)
  Dim ErrMsg
  Dim WshShell
  ErrMsg = "ERROR in " & company & " Backup: " & Err.Number & " " & Err.Description & vbNewLine

  WScript.Echo ErrMsg

  Err.Clear

  Set WshShell = WScript.CreateObject("WScript.Shell")
  WshShell.LogEvent 1, ErrMsg
  Set WshShell = Nothing

  Call SendEmail(emailFrom, emailTo, emailCC, emailReplyTo, "ERROR in " & company & " Backup", ErrMsg)
End Sub

'***************************************************************************************

Sub PerformBackup()
  On Error Resume Next
  Dim Y: Y = Year(now())
  Dim M: M = Right("0" & Month(now()), 2)
  Dim D: D = Right("0" & Day(now()), 2)
  Dim WshShell
  Dim command
  Dim i
  Dim errNum
  Dim myBackupFile

  myBackupFile = backupDest & company & "-" & Y & M & D & ".iso"

  command = ""
  command = command & """" & ImgBurnDirectory & "ImgBurn.exe"" /MODE BUILD /BUILDOUTPUTMODE IMAGEFILE "
  command = command & "/SRC """
  For i = 0 To UBound(backupSourceInfo)
    command = command & backupSourceInfo(i)(1)
    If i < UBound(backupSourceInfo) Then
      command = command & "|"
    End if
  Next
  command = command & """"
  command = command & " /DEST """ & myBackupFile & """ /FILESYSTEM ""ISO9660 + UDF"" /UDFREVISION ""1.02"""
  command = command & " /VOLUMELABEL """ & myBackupFile & """"
  command = command & " /PRESERVEFULLPATHNAMES NO /VERIFY YES /START /CLOSE /NOIMAGEDETAILS /OVERWRITE YES"
  
  Set WshShell = WScript.CreateObject("WScript.Shell")
  errNum = WshShell.Run(command, 0, true)
  Set WshShell = Nothing
  
  On Error GoTo 0
  If Err.Number <> 0 Then
    Err.Raise vbObjectError + 1, , "Error " & Err.Number & " While Trying To Launch Process with command:" & vbNewLine & command & vbNewLine
  elseIf errNum <> 0 Then
    Err.Raise vbObjectError + 1, , "Error " & errNum & " Returned From Launching ImgBurn with command:" & vbNewLine & command & vbNewLine
  End If
End Sub

'***************************************************************************************

Sub SendEmail(ByVal strfrom, ByVal toRecip, ByVal strcc, ByVal replyTo, ByVal strsubject, ByVal strbody)
  Call SendGMail(toRecip, strfrom, strcc, "", replyTo, strsubject, strbody, mailPlain)
End Sub
   
'***************************************************************************************

Sub SendGmail(ByVal strTo, ByVal strFrom, ByVal strCC, ByVal strBCC, ByVal strReplyTo, ByVal Subject, ByVal msgBody, ByVal mailType)
  On Error Resume Next
  Dim SMTPServer, SMTPusername, SMTPpassword
  Dim sch, cdoConfig, cdoMessage
  SMTPserver   = "smtp.gmail.com"
  SMTPusername = "nearchives@gmail.com"
  SMTPpassword = "Hello$World"
  sch = "http://schemas.microsoft.com/cdo/configuration/"
  Set cdoConfig = WScript.CreateObject("CDO.Configuration")
  With cdoConfig.Fields
      .Item(sch & "smtpauthenticate") = 1
      .Item(sch & "smtpusessl") = True
      .Item(sch & "smtpserver") = SMTPserver
      .Item(sch & "sendusername") = SMTPusername
      .Item(sch & "sendpassword") = SMTPpassword
      .Item(sch & "smtpserverport") = 465 '587
      .Item(sch & "sendusing") = 2
      .Item(sch & "connectiontimeout") = 100
      .update
  End With
  Set cdoMessage = WScript.CreateObject("CDO.Message")
  Set cdoMessage.Configuration = cdoConfig
  cdoMessage.From = strFrom
  cdoMessage.To = strTo
  cdoMessage.Cc = strCC
  cdoMessage.Bcc = strBCC
  cdoMessage.ReplyTo = strReplyTo
  cdoMessage.Subject = Subject
  if mailtype = mailHTML then
      cdoMessage.HTMLBody = msgBody
  Else
      cdoMessage.TextBody = msgBody
  End If
  cdoMessage.Send
  Set cdoMessage = Nothing
  Set cdoConfig = Nothing
End Sub

'***************************************************************************************

Function iif(ByVal cond, ByVal tpart, ByVal fpart)
  If cond Then
    iif = tpart
  Else
    iif = fpart
  End if
End Function

'***************************************************************************************

Function FileExists(ByVal File)
  Dim filesys
  Set filesys = WScript.CreateObject("Scripting.FileSystemObject")
  FileExists = filesys.FileExists(File)
  Set filesys = Nothing
End Function

'***************************************************************************************

Function IDOK(ByVal ID)
  Dim rx
  Set rx = New RegExp
  rx.Pattern = "^[A-Z0-9]{1,5}$"
  rx.Global = True
  rx.IgnoreCase = True
  IDOK = rx.Test(ID)
  Set rx = Nothing
End Function

'***************************************************************************************

Function BackupDestExists(ByVal folder)
  On Error Resume Next
  Dim filesys
  Set filesys = WScript.CreateObject("Scripting.FileSystemObject")
  BackupDestExists = filesys.FolderExists(folder)
  If Err.Number <> 0 Then
    BackupDestExists = False
  End if
  Set filesys = Nothing
End Function

'***************************************************************************************

Function EmailOK(ByVal Email)
  Dim rx
  Set rx = New RegExp
  rx.Pattern = "^([0-9a-zA-Z]([-.\w]*[0-9a-zA-Z])*@(([0-9a-zA-Z])+([-\w]*[0-9a-zA-Z])*\.)+[a-zA-Z]{2,9})$"
  rx.Global = True
  rx.IgnoreCase = True
  EmailOK = rx.Test(Email)
  Set rx = Nothing
End Function

'***************************************************************************************

Function IsCScript()
  Select Case Ucase(StrReverse(Left(StrReverse(WScript.FullName), Instr(StrReverse(WScript.FullName),"\") - 1)))
    Case "CSCRIPT.EXE"
      IsCscript = True
    Case Else
      IsCscript = False
  End Select
End Function
