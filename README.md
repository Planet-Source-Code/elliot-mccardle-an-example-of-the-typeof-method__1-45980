<div align="center">

## \_An Example of the TypeOf method


</div>

### Description

I had several forms with several controls on each. I wanted to clear each control, but I wanted to make it as painless as possible. This code takes a form and iterates through each control clearing the values. Currently, it works with text boxes, combo boxes, data pickers, and masked edit controls. all you have to do is pass a form name to the procedure. It uses the for each...next loop, along with the typeof method. I hope this helps.
 
### More Info
 
for name


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Elliot McCardle](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/elliot-mccardle.md)
**Level**          |Beginner
**User Rating**    |3.4 (17 globes from 5 users)
**Compatibility**  |VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/elliot-mccardle-an-example-of-the-typeof-method__1-45980/archive/master.zip)





### Source Code

```
Public Sub ClearForm(frm As Form) 'Pass a form name to it
 Dim sMask As String
 For Each Control In frm.Controls
  If TypeOf Control Is TextBox Or TypeOf Control Is ComboBox Then
   Control.Text = "" 'Clear text
  End If
  If TypeOf Control Is MaskEdBox Then
   With Control
    sMask = .Mask 'Save the existing mask
    .Mask = "" 'Clear mask
    .Text = "" 'Clear text
    .Mask = sMask 'Reset mask
   End With
  End If
  If TypeOf Control Is DTPicker Then
   Control.Date = Date 'Set to current date
  End If
 Next Control
End Sub
```

