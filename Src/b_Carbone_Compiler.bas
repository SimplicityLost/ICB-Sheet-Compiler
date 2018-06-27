Attribute VB_Name = "b_Carbone_Compiler"
Function carbcomp()
 r = 1
    n = 0
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With
    Call MsgBox("You have to give this program complete focus or it breaks.", vbOKOnly)
    With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = True
        .Title = "Select files to process"
        .Show
         
        If .SelectedItems.Count = 0 Then Exit Function
         
        Set wsMaster = ActiveWorkbook
         
        For File = 1 To .SelectedItems.Count
            On Error GoTo err1:
            Filename = .SelectedItems.Item(File)
             
            If Right(Filename, 4) = ".xls" Or Right(Filename, 5) = ".xlsx" Then
                    period = "EOM"
                    monthin = Val(Mid(Filename, (InStr(1, Filename, "Carbone Files", vbTextCompare) + 14), 2))
                    yearin = Val(Mid(Filename, (InStr(1, Filename, "Carbone Files", vbTextCompare) + 20), 2))
                
                Do While ShiftPressed()
                    DoEvents
                Loop
                
                Application.DisplayAlerts = False
                Workbooks.Open Filename, False, True
                 
                Set xlsFiles = ActiveWorkbook

                n = xlsFiles.Sheets("SUMMARY").Range("F:F").Cells.SpecialCells(xlCellTypeConstants).Count
                xlsFiles.Sheets("summary").Cells.ClearOutline
                
                xlsFiles.Sheets("SUmmary").Rows("2:" & n).Copy Destination:=wsMaster.Sheets("Output").Range("A" & r).Offset(1, 0)
                For i = r + 1 To r + n - 1
                    wsMaster.Sheets("Output").Range("A" & i).Value = monthin
                    wsMaster.Sheets("Output").Range("Y" & i).Value = period
                    wsMaster.Sheets("Output").Range("Q" & i).Value = yearin
                    wsMaster.Sheets("Output").Range("N" & i).Value = "Unknown"
                    wsMaster.Sheets("Output").Range("O" & i).Value = "Unknown"
                    wsMaster.Sheets("Output").Range("P" & i).Value = "Unknown"
                Next i
                
                r = r + n
gotothenextfile:
                Application.DisplayAlerts = False
                xlsFiles.Close SaveChanges:=False 'close without saving
                Application.DisplayAlerts = True
            End If

        Next File 'go to the next file and repeat the process
        
        On Error GoTo 0
        
        With wsMaster.Sheets("Output")
            .Cells.ClearFormats
            
            .Columns("C").Value = .Columns("F").Value
            .Columns("H").Value = .Columns("G").Value
            .Columns("G").Value = .Columns("B").Value
            .Columns("F").Value = .Columns("W").Value
            .Columns("E").Value = .Columns("J").Value
            .Columns("J").Value = .Columns("M").Value
            .Columns("M").Value = .Columns("D").Value
            .Columns("D").Value = .Columns("C").Value
            .Columns("I").Value = .Columns("K").Value
            .Columns("K").Clear
            .Columns("B").Value = .Columns("Y").Value
     
            .Columns("R:Z").Delete
            
            Dim rng As Range, rngDelete As Range
            Set rng = .Range("D:D")
            .Cells.Sort key1:=rng, order1:=xlDescending
            .Rows(1).Insert (xlDown)
            .Cells.AutoFilter Field:=7, Criteria1:="="
            Set rngDelete = rng.SpecialCells(xlCellTypeVisible)
            rngDelete.EntireRow.Delete

            Set rng = .Range("A:A")
            .Cells.Sort key1:=rng, order1:=xlDescending
            .Rows(1).Insert (xlDown)
            .Cells.AutoFilter Field:=1, Criteria1:="N/A"
            Set rngDelete = rng.SpecialCells(xlCellTypeVisible)
            rngDelete.EntireRow.Delete
            
            
        End With
        
    End With
    
    Set wsMaster = Nothing
    Set xlsFiles = Nothing
    
    r = 0
    n = 0
    ur = 0
    
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
    
Exit Function

err1:
Resume gotothenextfile:
End Function

