Attribute VB_Name = "bisecc"
Function fx(m)
fx = 2 * m ^ 3 + Log(m) - Cos(m) / Exp(m) + Sin(m)
End Function

Sub Limpiar()

'limpia celdas

Range("I10:P100").Select
Selection.ClearContents

End Sub

Sub BiseccionOK()
Sheets(5).Select

Call Limpiar

Xini = Cells(3, 2)
Xfin = Cells(4, 2)

Tolerancia = Cells(6, 2)


If fx(Xini) * fx(Xfin) < 0 Then
 i = 0
 ErrorAbs = 100
 
 While (ErrorAbs > Tolerancia And i < 100)
 
    Xm = (Xini + Xfin) / 2
    
    Cells(10 + i, 9) = i
    Cells(10 + i, 10) = Xini
    Cells(10 + i, 11) = Xfin
    Cells(10 + i, 12) = Xm
    Cells(10 + i, 13) = fx(Xini)
    Cells(10 + i, 14) = fx(Xfin)
    Cells(10 + i, 15) = fx(Xm)
    
        If i > 0 Then
         ErrorAbs = Abs(Xm - Xm_old)
         Cells(10 + i, 16) = ErrorAbs
        End If
            Xm_old = Xm

    If fx(Xini) * fx(Xm) < 0 Then
      Xfin = Xm
    Else
 
    If fx(Xini) * fx(Xm) > 0 Then
      Xini = Xm
    Else
      Err = Tolerancia
    End If
    
    End If

 i = i + 1

Wend

Else
MsgBox "No hay solucion en ese intervalo"
End If

Cells(7, 2) = Xm
End Sub
