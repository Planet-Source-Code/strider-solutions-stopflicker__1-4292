<div align="center">

## StopFlicker


</div>

### Description

Avoid the Flickering

Use this routine to stop a control (like a list or treeview) from flickering when it is getting it's data.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Strider Solutions](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/strider-solutions.md)
**Level**          |Unknown
**User Rating**    |4.2 (169 globes from 40 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/strider-solutions-stopflicker__1-4292/archive/master.zip)





### Source Code

```
'Get more great source code from
' http://www.stridersolutions.com/products/cs/
Option Explicit
#If Win16 Then
  Private Declare Function LockWindowUpdate Lib "User" (ByVal hWndLock As Integer) As Integer
#Else
  Private Declare Function LockWindowUpdate Lib "User32" (ByVal hWndLock As Long) As Long
#End If
Private Sub StopFlicker(ByVal lHWnd as Long)
  Dim lRet As Long
  'Object will not flicker - just be blank
  lRet = LockWindowUpdate(lHWnd)
 End Sub
Private Sub Release()
  Dim lRet As Long
  lRet = LockWindowUpdate(0)
End Sub
```

