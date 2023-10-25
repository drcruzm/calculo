Attribute VB_Name = "bisecc"
Function fx(m)
fx = 2 * m ^ 3 + Log(m) - Cos(m) / Exp(m) + Sin(m)
End Function

Sub Limpiar()

'limpia celdas

Range("I10:P100").Select
Selection.ClearContents

End Sub

Sub Biseccion()
Call Limpiar

Xini = Cells(3, 2).Value
Xfin = Cells(4, 2).Value

Tolerancia = Cells(6, 2).Value


If fx(Xini) * fx(Xfin) < 0 Then
 i = 0
 ErrorAbs = 100
 
 While (ErrorAbs > Tolerancia And i < 100)
    Cells(8, 16).Value = Err
 
    Xm = (Xini + Xfin) / 2
    
    Cells(10 + i, 9).Value = i
    Cells(10 + i, 10).Value = Xini
    Cells(10 + i, 11).Value = Xfin
    Cells(10 + i, 12).Value = Xm
    Cells(10 + i, 13).Value = fx(Xini)
    Cells(10 + i, 14).Value = fx(Xfin)
    Cells(10 + i, 15).Value = fx(Xm)
    
        If i > 0 Then
         ErrorAbs = Abs(Xm - Xm_old)
         Cells(10 + i, 16).Value = ErrorAbs
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
Cells(7, 2).Value = Xm
 i = i + 1

Wend

Else
MsgBox "No hay solucion en ese intervalo"

End If
End Sub
