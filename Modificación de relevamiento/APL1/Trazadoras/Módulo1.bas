Attribute VB_Name = "Módulo1"
Option Explicit
Sub Pintar()
Attribute Pintar.VB_ProcData.VB_Invoke_Func = "f\n14"
'
' Pintar Macro
'
' Acceso directo: CTRL+f
'
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 13082801
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
