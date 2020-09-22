VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSendEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''Requires: Microsoft Winsock control to be placed on a form
'
''Examples:
' Dim eml As New clsSendEmail
'
' '1. Send message without attachment to one person
' Call eml.SendEmail(Form1.Winsock1, "mail.server.com", "me@server.com", "you@another.com", "subject", "body")
'
' '2. Send message with attachment to one person
' Call eml.SendEmail(Form1.Winsock1, "mail.server.com", "me@server.com", "you@another.com", "subject", "body","c:\file.dat")
'
' '3. Send message without attachment to many people
' Call eml.SendEmail(Form1.Winsock1, "mail.server.com", "me@server.com", "you@another.com,her@another.com,him@company.com", "subject", "body")

Private m_intStage As Integer
Private WithEvents Winsock1 As Winsock
Attribute Winsock1.VB_VarHelpID = -1

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
    On Error Resume Next
    Winsock1.GetData strData, vbString
    m_intStage = m_intStage + 1
End Sub

Private Sub Winsock1_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox "Email Sending Error: " & Description, vbOKOnly
    m_intStage = 0
    If Winsock1.State <> sckClosed Then
        Winsock1.Close
    End If
End Sub

Public Sub SendEmail(Wnsk As Winsock, strMailServer As String, strFrom As String, strTo As String, strSubject As String, strBody As String, Optional strAttachment As String)
Dim strTemp As String
Dim intOldStage As Integer
Dim lngStart As Long
Dim lngTmr As Long
Dim strTo2 As String
    If Trim(strTo) = "" Then Exit Sub 'quit if no recipients
    Set Winsock1 = Wnsk
    While 1 = 1
        Select Case m_intStage
        Case 0:
            If Winsock1.State <> sckClosed Then
                Winsock1.Close
            End If
            Winsock1.LocalPort = 0
            Winsock1.Protocol = sckTCPProtocol
            Winsock1.Connect strMailServer, "25"
        Case 1:
            strTemp = Winsock1.LocalIP
            Winsock1.SendData "HELO " & strTemp & vbCrLf
        Case 2:
            Winsock1.SendData "MAIL FROM: <" & Trim(strFrom) & ">" & vbCrLf
        Case 3:
            If InStr(strTo, ",") > 0 Then 'multiple recipients
                m_intStage = m_intStage - 1 'set to repeat send for each recipient
                Winsock1.SendData "RCPT TO: <" & Trim(Left$(strTo, InStr(strTo, ",") - 1)) & ">" & vbCrLf
                strTo2 = strTo2 & Trim(Left$(strTo, InStr(strTo, ",")))
                strTo = Trim(Mid$(strTo, InStr(strTo, ",") + 1))
            Else
                Winsock1.SendData "RCPT TO: <" & Trim(strTo) & ">" & vbCrLf
                strTo2 = strTo2 & Trim(strTo)
                strTo = strTo2 'restore 'to' list
            End If
        Case 4:
            Winsock1.SendData "DATA" & vbCrLf
        Case 5:
            strTemp = ""
            strTemp = strTemp & "From: " & Chr$(34) & strFrom & Chr$(34) & " <" & strFrom & ">" & vbNewLine
            'strTemp = strTemp & "Reply-To: " & strFrom & vbNewLine
            strTemp = strTemp & "To: " & strTo & vbNewLine
            strTemp = strTemp & "Subject: " & strSubject & vbNewLine
            strTemp = strTemp & "X-Mailer: VB Program - " & App.EXEName & vbNewLine
            
            If Not IsMissing(strAttachment) Then
                If strAttachment <> "" Then
                    strTemp = strTemp & "Mime-Version: 1.0" & vbNewLine
                    strTemp = strTemp & "Content-Type: multipart/mixed; boundary=" & Chr(34) & "NextMimePart" & Chr(34) & vbNewLine & vbNewLine
                    strTemp = strTemp & "This is a multi-part message in MIME format." & vbNewLine
                    strTemp = strTemp & "--NextMimePart" & vbNewLine
                End If
            End If
            
            strTemp = strTemp & vbCrLf & strBody
            Winsock1.SendData strTemp
            
            If Not IsMissing(strAttachment) Then
                If strAttachment <> "" Then
                    Call SendMimeAttachement(strAttachment)
                End If
            End If
            
            Winsock1.SendData vbCrLf & "." & vbCrLf
        Case 6:
            Winsock1.SendData "QUIT" & vbCrLf
        Case Else
            m_intStage = 0
            Set Winsock1 = Nothing
            Exit Sub 'done sending
        End Select
        
        intOldStage = m_intStage 'wait for data arrival
        lngStart = Timer
        Do While intOldStage = m_intStage And Winsock1.State <> sckClosed
            lngTmr = Timer - lngStart
            DoEvents
            
            If lngTmr > 50 Then 'Wait 50 seconds for response
                MsgBox "SMTP service error, timed out while waiting for response", 64, "Error!"
                m_intStage = 0
                If Winsock1.State <> sckClosed Then
                    Winsock1.Close
                End If
                Exit Do
            End If
        Loop
        
        If m_intStage = 0 Then 'AND/OR Winsock1.State = sckClosed
            Set Winsock1 = Nothing
            Exit Sub 'winsock error or time occurred
        End If
    Wend
End Sub

Private Sub SendMimeAttachement(strPath As String)
Dim l As Long, i As Long, FileIn As Long
Dim strTemp As String
Dim b As Integer
Dim Base64Tab As Variant
Dim bin(3) As Byte
Dim s As Variant
    
    'Base64Tab holds the encode tab
    Base64Tab = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "+", "/")
    
    FileIn = FreeFile
    Open strPath For Binary As FileIn
    
    'Preparing the Mime Header
    strTemp = vbCrLf & "--NextMimePart" & vbNewLine
    strTemp = strTemp & "Content-Type: application/octet-stream; name=" & Chr(34) & GetFileName(strPath) & Chr(34) & vbNewLine
    strTemp = strTemp & "Content-Transfer-Encoding: base64" & vbNewLine
    strTemp = strTemp & "Content-Disposition: attachment; filename=" & Chr(34) & GetFileName(strPath) & Chr(34) & vbNewLine
    Winsock1.SendData strTemp & vbCrLf
    
    l = LOF(FileIn) - (LOF(FileIn) Mod 3)
    
    For i = 1 To l Step 3
        'Read three bytes
        Get FileIn, , bin(0)
        Get FileIn, , bin(1)
        Get FileIn, , bin(2)
        
        'Always wait until there're more then 64 characters
        If Len(s) > 64 Then
            Do
            Loop Until Winsock1.State = 7
            DoEvents
            s = s & vbCrLf
            Winsock1.SendData s
            s = ""
        End If
        
        'Calc Base64-encoded char
        b = (bin(0) \ 4) And &H3F 'right shift 2 bits (&H3F=111111b)
        
        'the character s holds the encoded chars
        s = s & Base64Tab(b)
        
        b = ((bin(0) And &H3) * 16) Or ((bin(1) \ 16) And &HF)
        s = s & Base64Tab(b)
        b = ((bin(1) And &HF) * 4) Or ((bin(2) \ 64) And &H3)
        s = s & Base64Tab(b)
        b = bin(2) And &H3F
        s = s & Base64Tab(b)
    Next i
    
    If Not (LOF(FileIn) Mod 3 = 0) Then 'check if there is something left
        For i = 1 To (LOF(FileIn) Mod 3) 'Reads the number of bytes left
            Get FileIn, , bin(i - 1)
        Next i
        If (LOF(FileIn) Mod 3) = 2 Then 'If there are only 2 chars left
            b = (bin(0) \ 4) And &H3F 'right shift 2 bits (&H3F=111111b)
            s = s & Base64Tab(b)
            b = ((bin(0) And &H3) * 16) Or ((bin(1) \ 16) And &HF)
            s = s & Base64Tab(b)
            b = ((bin(1) And &HF) * 4) Or ((bin(2) \ 64) And &H3)
            s = s & Base64Tab(b)
            s = s & "="
        Else 'If there is only one char left
            b = (bin(0) \ 4) And &H3F 'right shift 2 bits (&H3F=111111b)
            s = s & Base64Tab(b)
            b = ((bin(1) And &H3) * 16) Or ((bin(1) \ 16) And &HF)
            s = s & Base64Tab(b)
            s = s & "=="
        End If
    End If
    
    If s <> "" Then 'Send the characters left
        s = s & vbCrLf
        Winsock1.SendData s
    End If
    
    'Send the last part of the MIME Body
    Winsock1.SendData vbCrLf & "--NextMimePart--" & vbCrLf
    Close FileIn
End Sub

Private Function GetFileName(strFilePath As String) As String
Dim intStart As Integer
    intStart = 0
    While InStr(intStart + 1, strFilePath, "\") > 0
        intStart = InStr(intStart + 1, strFilePath, "\")
    Wend
    GetFileName = Mid$(strFilePath, intStart + 1)
End Function

