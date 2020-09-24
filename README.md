<div align="center">

## Place a Progress Bar in a Status Bar


</div>

### Description

Places a Progress Bar in a Status Bar, with Windows Common Controls 6, without any third party OCXs or API.
 
### More Info
 
The progress bar must have it's Appearance property set to 0-ccFlat and its BorderStyle set to 0-ccNone, it may be located anywhere on the form.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[xsteinberger](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/xsteinberger.md)
**Level**          |Beginner
**User Rating**    |4.0 (24 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/xsteinberger-place-a-progress-bar-in-a-status-bar__1-6015/archive/master.zip)





### Source Code

```
Private Sub Form_Paint()
 Dim WidthOfBorder As Single
 ScaleMode = vbTwips
 WidthOfBorder = (Width - ScaleWidth) / 4
 'assuming the progress bar is named ProgressBar1 and the status bar named StatusBar1, and placing the progress bar in panel 2
 'moving the progressbar to the statusbar and adjusting size
 ProgressBar1.Move StatusBar1.Panels(2).Left + 30, _
 StatusBar1.Top + WidthOfBorder + 20, _
 StatusBar1.Panels(2).Width - 50, _
 StatusBar1.Height - WidthOfBorder - 30
 'the values are hardcoded to allow the border to display to make the progressbar appear 3d and look smart. the progressbar may be hidden and replaced with text normally using the .panels().text property of the statusbar, as the progressbar is not actually in the statusbar, merely hovering above.
End Sub
```

