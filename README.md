<div align="center">

## Create Email in default client / Open URL


</div>

### Description

Simple calls to open up the default web browser to a given URL or to create an email message in the default email client.
 
### More Info
 
' To open a URL - Specify the URL

' OpenInternet Me, "http://www.search.com", Normal

'

' Specify Email Address and Subject

' OpenInternet Me, _

'  "mailto:anyone@domain.com?SUBJECT=Hello World", Normal

' I am NOT the original author of this code, but posting

' it because I haven't found this type of implementation

' here yet.

' Does not send the Email message - up to you to do that


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[John 'Hemo' Hiemenz](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/john-hemo-hiemenz.md)
**Level**          |Beginner
**User Rating**    |3.7 (11 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/john-hemo-hiemenz-create-email-in-default-client-open-url__1-5685/archive/master.zip)

### API Declarations

```
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
  (ByVal hwnd As Long, ByVal lpOperation As String, _
   ByVal lpFile As String, ByVal lpParameters As String, _
   ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Enum T_WindowStyle
Maximized = 3
Normal = 1
ShowOnly = 5
End Enum
```


### Source Code

```
Public Sub OpenInternet(Parent As Form, URL As String, _
            WindowStyle As T_WindowStyle)
ShellExecute Parent.hwnd, "Open", URL, "", "", WindowStyle
End Sub
```

