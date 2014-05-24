Private Sub cmd1_Click()
    ' Dichiarazione delle variabili
    Dim lunghezza As Byte
    Dim cella_I As Variant
    Dim cella_J As Byte
    Dim colonna_0 As Byte
    Dim riga_0 As Byte
    Dim lambda As Byte
    Dim NumeroFattori As Byte
    Dim ultimaRigaPiena As Byte
    Dim ultimaColonnaPiena As Byte
    
    ' Intestazione del triangolo(Anni di sviluppo e di accadimento)
    lunghezza = Application.WorksheetFunction.Count(Range(Ref_I))
    NumeroFattori = lunghezza
    Range(Cells(ActiveCell.Row, ActiveCell.Column), Cells(ActiveCell.Row, ActiveCell.Column + lunghezza - 1)) = Range(Ref_J).Value
    With _
    Range(Cells(ActiveCell.Row - 1, ActiveCell.Column), Cells(ActiveCell.Row - 1, ActiveCell.Column + lunghezza - 1))
    .Merge
    If optIncrementali Then
        .Value = "Pagamenti " & optIncrementali.Caption
        .HorizontalAlignment = xlCenter
        
    Else
        .Value = "Pagamenti " & optCumulati.Caption
        .HorizontalAlignment = xlCenter
    End If
    End With
    ActiveCell.Offset(1, -1).Select
    Range(Cells(ActiveCell.Row, ActiveCell.Column), Cells(ActiveCell.Row + lunghezza - 1, ActiveCell.Column)) = Range(Ref_I).Value
    ActiveCell.Offset(0, 1).Select
    
    
    If optIncrementali Then
    ' Calcolo dei pagamenti cumulati
    colonna_0 = ActiveCell.Column
    riga_0 = ActiveCell.Row
    For Each cella_I In Range(Ref_I)
        For cella_J = 0 To lunghezza - 1
            ActiveCell.Value = _
            Application.WorksheetFunction.Sum(Range(Cells(cella_I.Row, cella_I.Column + 1), Cells(cella_I.Row, cella_I.Column + 1 + cella_J)))
            ActiveCell.Offset(0, 1).Select
        Next cella_J
        ActiveCell.Offset(1, -lunghezza).Select
        lunghezza = lunghezza - 1
    Next cella_I
    
    ' Stima dei fattori di sviluppo
    ActiveCell.Offset(1, 1).Select
    For cella_J = colonna_0 + 1 To colonna_0 + NumeroFattori - 1
        ActiveCell.Value = _
        Application.WorksheetFunction.Sum(Range(Cells(riga_0, cella_J), Cells(riga_0 + NumeroFattori - 1 - cella_J + colonna_0, cella_J))) / _
        Application.WorksheetFunction.Sum(Range(Cells(riga_0, cella_J - 1), Cells(riga_0 + NumeroFattori - 1 - cella_J + colonna_0, cella_J - 1)))
        ActiveCell.Offset(0, 1).Select
    Next cella_J
    ActiveCell.Offset(0, -NumeroFattori).Value = "Fattori di sviluppo"
    
    ' Stima dei pagamenti cumulati futuri
    For cella_I = riga_0 + 1 To riga_0 + NumeroFattori - 1
            ' Per ogni riga trovo la prima cella vuota
            Range(Cells(cella_I, colonna_0 - 1), Cells(cella_I, colonna_0 + NumeroFattori - 1)).End(xlToRight).Select
            ultimaRigaPiena = ActiveCell.Row
            ultimaColonnaPiena = ActiveCell.Column
            ' il range comprende anche la colonna di intestazione in modo da evitare i problemi che si presentano con l'ultima
            ' iterazione che in caso contrario avrebbe una sola cella piena
            ' Seleziono la prima cella vuota
            ActiveCell.Offset(0, 1).Select
            For cella_J = 0 To colonna_0 + NumeroFattori - 1 - ActiveCell.Column
                ActiveCell.Value = ActiveCell.Offset(0, -1).Value * Cells(riga_0 + NumeroFattori + 1, ActiveCell.Column).Value
                ActiveCell.Font.Color = -16776961
                If ActiveCell.Column = colonna_0 + NumeroFattori - 1 Then
                   ActiveCell.Offset(0, 2).Value = ActiveCell.Value - Cells(ultimaRigaPiena, ultimaColonnaPiena).Value
                End If
                With _
                Cells(Range(Ref_J).Row + 1 + ActiveCell.Row - riga_0, Range(Ref_I).Column + 1 + ActiveCell.Column - colonna_0)
                .Value = ActiveCell.Value - ActiveCell.Offset(0, -1).Value
                .Font.Color = -16776961
                End With
                ActiveCell.Offset(0, 1).Select
            Next cella_J
    Next cella_I
    
    ' Stima della riserva sinistri come somma degli OLL
    ActiveCell.Offset(2, 1).Select
    With ActiveCell
    .Value = Application.WorksheetFunction.Sum(Range(Cells(ActiveCell.Row - NumeroFattori, ActiveCell.Column), _
    Cells(ActiveCell.Row - 2, ActiveCell.Column)))
    .Font.Color = -16776961
    End With
    ActiveCell.Offset(-(NumeroFattori + 1), 0).Value = "O.L.L."
    ActiveCell.Offset(0, 1).Value = "Stima della riserva sinistri"
    Else
        If optCumulati Then
            ' Inserire un controllo che verifichi se i dati di partenza riguardano effettivamente pagamenti cumulati
            ' Stima dei fattori di sviluppo
            ActiveCell.Offset(NumeroFattori + 1, 0).Value = "Fattori di sviluppo"
            ActiveCell.Offset(NumeroFattori + 1, 1).Select
            colonna_0 = Range(Ref_I).Column + 1
            riga_0 = Range(Ref_J).Row + 1
            For cella_J = 0 To NumeroFattori - 2
                ActiveCell.Value = _
                Application.WorksheetFunction.Sum(Range(Cells(riga_0, colonna_0 + cella_J + 1), Cells(riga_0 + NumeroFattori - 2 - cella_J, colonna_0 + cella_J + 1))) / _
                Application.WorksheetFunction.Sum(Range(Cells(riga_0, colonna_0 + cella_J), Cells(riga_0 + NumeroFattori - 2 - cella_J, colonna_0 + cella_J)))
                ActiveCell.Offset(0, 1).Select
            Next cella_J
            ActiveCell.Offset(0, -NumeroFattori).Select
            
            'Completamento triangolo cumulati, calcolo OLL e Riserva sinistri
            For cella_I = 0 To NumeroFattori - 1
               For cella_J = 0 To NumeroFattori - 1
                      'Completamento triangolo dei cumulati
                    If cella_I + cella_J > NumeroFattori - 1 Then
                       With Cells(riga_0 + cella_I, colonna_0 + cella_J)
                        .Value = _
                        Cells(riga_0 + cella_I, colonna_0 + NumeroFattori - 1 - cella_I) * _
                        Application.WorksheetFunction.Product(Range( _
                        Cells(ActiveCell.Row, ActiveCell.Offset(0, NumeroFattori - cella_I).Column), _
                        Cells(ActiveCell.Row, ActiveCell.Offset(0, cella_J).Column)))
                        .Font.Color = -16776961
                        End With
                    End If
                    
                    'Calcolo Pagamenti incrementali
                    If cella_J = 0 Then
                        With _
                        ActiveCell.Offset(-(NumeroFattori - cella_I + 1), cella_J)
                        .Value = Cells(riga_0 + cella_I, colonna_0)
                        .Font.Color = -16776961
                        End With
                    Else
                        With _
                        ActiveCell.Offset(-(NumeroFattori - cella_I + 1), cella_J)
                        .Value = Cells(riga_0 + cella_I, colonna_0 + cella_J) - Cells(riga_0 + cella_I, colonna_0 + cella_J - 1)
                        .Font.Color = -16776961
                        End With
                    End If
                    
                Next cella_J
                ' Stima  OLL
                Cells(riga_0 + cella_I, colonna_0 + NumeroFattori + 1) = _
                Cells(riga_0 + cella_I, colonna_0 + NumeroFattori - 1) - _
                Cells(riga_0 + cella_I, colonna_0 + NumeroFattori - 1 - cella_I)
                
            Next cella_I
            With _
            Cells(riga_0 + NumeroFattori, colonna_0 + NumeroFattori + 1)
            .Value = Application.WorksheetFunction.Sum(Range(Cells(riga_0, colonna_0 + NumeroFattori + 1), _
            Cells(riga_0 + NumeroFattori - 1, colonna_0 + NumeroFattori + 1)))
            .Font.Color = -16776961
            End With
            Cells(riga_0 + NumeroFattori, colonna_0 + NumeroFattori + 2) = "Stima della riserva sinistri"
            Cells(riga_0 - 1, colonna_0 + NumeroFattori + 1) = "O. L. L."
        Else
            MsgBox ("Scegli il tipo di dati di partenza")
        End If
    End If
    Unload Riferimenti
End Sub



Private Sub optCumulati_Click()

End Sub

Private Sub optIncrementali_Click()

End Sub

Private Sub Ref_I_BeforeDragOver(Cancel As Boolean, ByVal Data As MSForms.DataObject, ByVal x As stdole.OLE_XPOS_CONTAINER, ByVal y As stdole.OLE_YPOS_CONTAINER, ByVal DragState As MSForms.fmDragState, Effect As MSForms.fmDropEffect, ByVal Shift As Integer)

End Sub

Private Sub Ref_J_BeforeDragOver(Cancel As Boolean, ByVal Data As MSForms.DataObject, ByVal x As stdole.OLE_XPOS_CONTAINER, ByVal y As stdole.OLE_YPOS_CONTAINER, ByVal DragState As MSForms.fmDragState, Effect As MSForms.fmDropEffect, ByVal Shift As Integer)

End Sub
