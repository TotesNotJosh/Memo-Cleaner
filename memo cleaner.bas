'Put this code into your module
Sub clean_email_body()
    Dim outlook As Object
    Dim email As Object
    Dim body_text As String
    Set outlook = CreateObject("Outlook.Application")
    Set email = outlook.ActiveExplorer.Selection.Item(1)
    If TypeName(email) = "MailItem" Then
        body_text = email.Body
        bodyText = remove_white_space(body_text)
        email.Body = body_text
        MsgBox "Email Cleaned"
    End If
End Sub

Function remove_white_space(text As String)
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "[" & ChrW(&HC) & ChrW(&H20) & ChrW(&H2000) & ChrW(&H2001) & ChrW(&H2002) & ChrW(&H2003) & _
    ChrW(&H2004) & ChrW(&H2005) & ChrW(&H2006) & ChrW(&H2007) & ChrW(&H2008) & ChrW(&H2009) & ChrW(&H200A) & _
    ChrW(&H200B) & ChrW(&H200C) & ChrW(&H200D) & ChrW(&H202F) & ChrW(&H205F) & ChrW(&H3000) & "]"
    regex.Global = True
    text = regex.Replace(text, " ")
    regex.Pattern = "[ ]{2,}"
    text = regex.Replace(text, " ")
    remove_white_space = Trim(text)
End Function
