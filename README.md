<div align="center">

## Outlook Automation w/ body formattting


</div>

### Description

Use VB to automate sending an email behind the scenes with Outlook. Allows text formatting of the body message (IE, bold, font size, color, italics, etc.)
 
### More Info
 
Assume you are using Microsoft Outlook for email


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jason Allen](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jason-allen.md)
**Level**          |Beginner
**User Rating**    |3.7 (26 globes from 7 users)
**Compatibility**  |VB 5\.0, VB 6\.0, VBA MS Access, VBA MS Excel
**Category**       |[Microsoft Office Apps/VBA](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/microsoft-office-apps-vba__1-42.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jason-allen-outlook-automation-w-body-formattting__1-31075/archive/master.zip)





### Source Code

```
'This code shows how to use VB to automate creating an Outlook Email. The main point to this is
  'to show how you can use automation and still format your email body with bolding, font colors, sizes, etc.
  'You must have the Outlook reference library checked for this code to use. Goto
  'References and select the Microsoft Outlook Object Library.
  'strings that store your variables for the email
  Dim mySubject As String
  Dim myBody As String
  Dim myText As String
  Dim myEmail As String
  'this is where you can enter your body message. You can add on later by appending with the & symbol
  myBody = myBody & "You can use any HTML tags to format your body message as you want."
  'your subject goes here
  mySubject = "This is my subject"
  'declare a new instance of Outlook
  Dim myOutlook As New Outlook.Application
  'set the variable as a new outlook app
  Set myOutlook = New Outlook.Application
  'set a variable as a new outlook mail item
  Set myMessage = myOutlook.CreateItem(olMailItem)
  'this copies your body variable into the 'body' of the email.
  'note the HTML declaration, as this allows you to use bold, italic, etc. tags in your body with Automation.
  'This is the reason I created this code sample, so you can see how to format an email body any way you want
  'using automation (you could use font tags and set font colors and everything this way.)
  myMessage.HTMLBody = myBody
  'sends your email
  myMessage.send
  'puts your subject variable into the 'subject' field
  myMessage.Subject = mySubject
  'puts your email address variable into the 'to' field
  myMessage.To = myEmail
  'Empties the Variables from memory
  Set myMessage = Nothing
  Set myOutlook = Nothing
  'confirmation dialog box
  myResult = MsgBox("Your email has been sent.", vbOKOnly, "Email Sent")
```

