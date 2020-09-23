<div align="center">

## Combo Auto\-Search


</div>

### Description

Automatically search through a combo-box's list as the user types into the text portion of the control. The code exists for allowing entry of text that does not

already appear in the list, as well as handling the Backspace and Delete keys in a logical manner.
 
### More Info
 
This code, as it stands, assumes that the form contains a Combo Box control, named Combo1, with the Style property set to 0 - Dropdown Combo or

1 - Simple Combo.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[O'Neill](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/o-neill.md)
**Level**          |Unknown
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/o-neill-combo-auto-search__1-1267/archive/master.zip)





### Source Code

```
Option Explicit
' This code demonstartes an auto-search combo box.
' As the user types into the combo, the list is searched, and if a
' partial match is made, then the remaining text is entered into the
' Text portion of the combo, and selected so that any further
' typing will automatically overwrite the Auto-search results.
'
' The IgnoreTextChange flag is used internally to tell the
' Combo1_Changed event not to perform the Auto-search.
Dim IgnoreTextChange As Boolean
Private Sub Combo1_Change()
  Dim i%
  Dim NewText$
  ' Check to see if a serch is required.
  If Not IgnoreTextChange And Combo1.ListCount > 0 Then
    ' Loop through the list searching for a partial match of
    ' the entered text.
    For i = 0 To Combo1.ListCount - 1
      NewText = Combo1.List(i)
      If InStr(1, NewText, Combo1.Text, 1) = 1 Then
        If Len(Combo1.Text) <> Len(NewText) Then
          ' Partial match found
          ' Avoid recursively entering this event
          IgnoreTextChange = True
          i = Len(Combo1.Text)
          ' Attach the full text from the list to what has
          ' already been entered. This technique preserves
          ' the case entered by the user.
          Combo1.Text = Combo1.Text & Mid$(NewText, i + 1)
          ' Select the text that is auto-entered
          Combo1.SelStart = i
          Combo1.SelLength = Len(Mid$(NewText, i + 1))
          Exit For
        End If
      End If
    Next
  Else
    ' The IgnoreTwextChange Flag is only effective for one
    ' Changed event.
    IgnoreTextChange = False
  End If
End Sub
Private Sub Combo1_GotFocus()
  ' Select existing text on entry to the combo box
  Combo1.SelStart = 0
  Combo1.SelLength = Len(Combo1.Text)
End Sub
Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
  ' If a user presses the "Delete" key, then the selected text
  ' is removed.
  If KeyCode = vbKeyDelete And Combo1.SelText <> "" Then
    ' Make sure that the text is not automatically re-entered
    ' as soon as it is deleted
    IgnoreTextChange = True
    Combo1.SelText = ""
    KeyCode = 0
  End If
End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
  ' If a user presses the "Backspace" key, then the selected text
  ' is removed. Autosearch is not re-performed, as that would only
  ' put it straight back again.
  If KeyAscii = 8 Then
    IgnoreTextChange = True
    If Len(Combo1.SelText) Then
      Combo1.SelText = ""
      KeyAscii = 0
    End If
  End If
End Sub
```

