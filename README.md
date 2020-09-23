<div align="center">

## Copy a listbox to clipboard


</div>

### Description

Have you ever tried to the contents a listbox to the clipboard? Annoying, huh? Well this code can help! This will copy all the contents in a listbox and put ", " between each one
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Joey Schoblaska](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/joey-schoblaska.md)
**Level**          |Unknown
**User Rating**    |4.0 (16 globes from 4 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB Script
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/joey-schoblaska-copy-a-listbox-to-clipboard__1-2998/archive/master.zip)





### Source Code

```
'Add a textbox, a listbox, and a command button
'Put this code in the command button
Clipboard.SetText ""
Text1.Text = ""
For X = 0 To List1.ListCount - 1
Text1.Text = Text1.Text & List1.List(X) & ", "
Next X
Clipboard.SetText Text1
```

