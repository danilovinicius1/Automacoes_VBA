Public WrkB                As Workbook
Public WrkS                As Worksheet

Public IntervaloMailing    As Range
Public Celula              As Range

Public AppOutk As Outlook.Application
Public MailOutk As Outlook.MailItem
Dim email As String

'Dim Account As String


Declare Sub Sleep Lib "kernel32" _
(ByVal dwMilliseconds As Long)


Public Sub MandarEmail()

Set WrkB = ThisWorkbook
Set WrkS = WrkB.Sheets("Enviar_Email")

Set IntervaloMailing = WrkS.Range("A6:A100000")


With WrkS
    .Select
        For Each Celula In IntervaloMailing
            Call CriaEmail
            Sleep (1000)
            Next
        
End With

End Sub

Sub CriaEmail()
On Error GoTo Erro
Set AppOutk = New Outlook.Application
Set MailOutk = AppOutk.CreateItem(olMailItem)
If WrkS.Cells(Celula.Row, 2) = 0 Then
MsgBox "ENVIO DOS E-MAIL'S CONCLUÍDO !", vbInformation, "Envio Concluído"
End
End If
email = WrkS.Cells(2, 2)
With MailOutk
    .Display
    .SentOnBehalfOfName = email
    .To = WrkS.Cells(Celula.Row, 2).Value
    .CC = WrkS.Cells(Celula.Row, 3).Value
    .BCC = WrkS.Cells(Celula.Row, 4).Value
    .Subject = WrkS.Cells(Celula.Row, 5).Value
    'Body = WrkS.Cells(Celula.Row, 6).Value
    '.HTMLBody = WrkS.Cells(Celula.Row, 6).Value & .HTMLBody
    .Attachments.Add WrkS.Cells(Celula.Row, 7).Value
    .Attachments.Add WrkS.Cells(Celula.Row, 8).Value
    .Attachments.Add WrkS.Cells(Celula.Row, 9).Value
    .Attachments.Add WrkS.Cells(Celula.Row, 10).Value
    .Attachments.Add WrkS.Cells(Celula.Row, 11).Value
    .Attachments.Add WrkS.Cells(Celula.Row, 12).Value
Erro:
.Importance = olImportanceHigh
.Send
End With

Set MailOutk = Nothing
Set AppOutk = Nothing

End Sub
