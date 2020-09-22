<div align="center">

## A\+\+ Transparent Forms \*\*MUST SEE\*\*


</div>

### Description

Makes forms transparent like glass- coooooool effect ... WinAmp3 uses the SAME effect under win2000 and XP!!!

here on psc are some codes, but this is 10000 times better - it autoupdates, is very fast and you can adjust the effect from 0 to 254

please vote for me :-))
 
### More Info
 
works only under win2000 - XP


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Daniel Pramel](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/daniel-pramel.md)
**Level**          |Advanced
**User Rating**    |4.1 (145 globes from 35 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Graphics](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/graphics__1-46.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/daniel-pramel-a-transparent-forms-must-see__1-28630/archive/master.zip)

### API Declarations

```
Option Explicit
Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Declare Function UpdateLayeredWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
Public Type POINTAPI
 x As Long
 y As Long
End Type
Public Type SIZE
 cx As Long
 cy As Long
End Type
Public Type BLENDFUNCTION
 BlendOp As Byte
 BlendFlags As Byte
 SourceConstantAlpha As Byte
 AlphaFormat As Byte
End Type
Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000
Public Const ULW_COLORKEY = &H1
Public Const ULW_ALPHA = &H2
Public Const ULW_OPAQUE = &H4
Public Const AC_SRC_OVER = &H0
Public Const AC_SRC_ALPHA = &H1
Public Const AC_SRC_NO_PREMULT_ALPHA = &H1
Public Const AC_SRC_NO_ALPHA = &H2
Public Const AC_DST_NO_PREMULT_ALPHA = &H10
Public Const AC_DST_NO_ALPHA = &H20
Public Const LWA_COLORKEY = &H1
Public Const LWA_ALPHA = &H2
```


### Source Code

```
Public Sub Mache_Transparent(hWnd As Long, Rate As Byte)
'### funktioniert nur unter Win2000 - XP !!!
'### macht das Fenster, dessen hWnd übergeben wurde, transparent
'### Rate: 254 = normal 0 = ganz transparent
'### 190 ist z.B. ein guter Wert
Dim WinInfo As Long
 WinInfo = GetWindowLong(hWnd, GWL_EXSTYLE)
 WinInfo = WinInfo Or WS_EX_LAYERED
 SetWindowLong hWnd, GWL_EXSTYLE, WinInfo
 SetLayeredWindowAttributes hWnd, 0, Rate, LWA_ALPHA
End Sub
```

