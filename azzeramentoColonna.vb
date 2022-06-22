Sub azzera_righe()
    Dim rigaIniziale As Integer
    Dim rigaFinale As Integer
    Dim colonnaAzzeramento As String
    Dim colonnaVariabile As String
    
    'Indicare intervallo di righe nel quale effettuare azzeramento
    rigaIniziale = InputBox(Prompt:="Riga iniziale (as integer)")
    rigaFinale = InputBox(Prompt:="Riga Finale (as integer)")
    
    'Inidicare la colonna da azzerare
    colonnaAzzeramento = InputBox(Prompt:="Colonna di azzeramento (as char, e.g. A)")
    
    'Inidicare la colonna della variabile su cui azzerare
    colonnaVariabile = InputBox(Prompt:="Colonna delle variabili (as char, e.g. A)")

    For i = rigaIniziale To rigaFinale
        ' Cells(riga,colonna)
        Cells(i, colonnaAzzeramento).GoalSeek goal:=0, ChangingCell:=Cells(i, colonnaVariabile)
    Next i
End Sub

