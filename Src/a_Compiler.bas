Attribute VB_Name = "a_Compiler"
Option Explicit
Dim wsMaster As Workbook, xlsFiles As Workbook
Dim Filename As String
Dim File As Integer
Dim r As Long
Dim ur As Long
Dim n As Long
Dim monthin As String
Dim period As String
Dim i As Long
Dim storenum As String
Dim storeindex As Long
Dim yearin As Integer
Dim findcol As Range
Dim findnext As Range
Dim errorlist As String

'Declare API
Declare PtrSafe Function GetKeyState Lib "User32" _
(ByVal vKey As Integer) As Integer
Const SHIFT_KEY = 16

Public Function Consolidate()
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
                If (InStr(1, Filename, "EOM", vbTextCompare) > 0) Then period = "EOM"
                If (InStr(1, Filename, "MID", vbTextCompare) > 0) Then period = "MID"
                If (InStr(1, Filename, "-ICB", vbTextCompare) > 0) Then
                    storenum = Mid(Filename, (InStr(1, Filename, "-ICB", vbTextCompare) - 4), 4)
                    monthin = Mid(Filename, (InStr(1, Filename, "-ICB", vbTextCompare) + 4), 2)
                    yearin = Mid(Filename, (InStr(1, Filename, "-ICB", vbTextCompare) + 6), 2)
                Else
                    storenum = "N/A"
                    monthin = "N/A"
                End If
                
                Do While ShiftPressed()
                    DoEvents
                Loop

                Workbooks.Open Filename, 0, True
                 
                Set xlsFiles = ActiveWorkbook
                 
                'r = wsMaster.Sheets("Output").UsedRange.Rows.Count

                ur = xlsFiles.Sheets("DATA").UsedRange.Rows.Count
                n = xlsFiles.Sheets("DATA").Range("J:J").Cells.SpecialCells(xlCellTypeConstants).Count

                If xlsFiles.Sheets("DATA").Range("E1").Value <> "Control" Then
                    xlsFiles.Sheets("DATA").Columns("e").Insert (xlToRight)
                    xlsFiles.Sheets("DATA").Columns("j").Insert (xlToRight)
                    xlsFiles.Sheets("DATA").Columns("h").Delete
                End If
                    
                If xlsFiles.Sheets("DATA").Range("J1").Value <> "Amount" Then
                    Set findcol = xlsFiles.Sheets("DATA").Range("A1:Z1").Find("Amount", LookIn:=xlValues, SearchFormat:=False)
                    xlsFiles.Sheets("DATA").Columns("J").Insert (xlToRight)
                    xlsFiles.Sheets("DATA").Columns("J").Value = findcol.EntireColumn.Value
                End If

                If xlsFiles.Sheets("DATA").Range("H1").Value <> "Vendor Name" Then
                    Set findcol = xlsFiles.Sheets("DATA").Range("A1:Z1").Find("Vendor Name", LookIn:=xlValues, SearchFormat:=False)
                    xlsFiles.Sheets("DATA").Columns("H").Insert (xlToRight)
                    xlsFiles.Sheets("DATA").Columns("H").Value = findcol.EntireColumn.Value
                End If
                
 
                xlsFiles.Sheets("DATA").Rows("2:" & n).Copy Destination:=wsMaster.Sheets("Output").Range("A" & r).Offset(1, 0)
                For i = r + 1 To r + n - 1
                    wsMaster.Sheets("Output").Range("O" & i).Value = monthin
                    wsMaster.Sheets("Output").Range("P" & i).Value = period
                    wsMaster.Sheets("Output").Range("Q" & i).Value = storenum
                    wsMaster.Sheets("Output").Range("R" & i).Value = yearin
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
            .Columns("F").Delete
            .Columns("M").Delete
            .Columns("A:C").Insert (xlToRight)
            .Columns("A:C").Value = .Columns("P:R").Value
            .Columns("P:R").Delete
            .Columns("K").Insert (xlToRight)
     
            Dim rng As Range, rngDelete As Range
            Set rng = .Range("J:J")
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
            
            Set rng = .Range("E:E")
            .Cells.Sort key1:=rng, order1:=xlDescending
            .Rows(1).Insert (xlDown)
            .Cells.AutoFilter Field:=5, Criteria1:="<70000"
            Set rngDelete = rng.SpecialCells(xlCellTypeVisible)
            rngDelete.EntireRow.Delete
            
            Set rng = .Range("E:E")
            .Cells.Sort key1:=rng, order1:=xlDescending
            .Rows(1).Insert (xlDown)
            .Cells.AutoFilter Field:=5, Criteria1:=">90000"
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

Function ShiftPressed() As Boolean
'Returns True if shift key is pressed
    ShiftPressed = GetKeyState(SHIFT_KEY) < 0
End Function

Function OldConsolidate()
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
             
            If Right(Filename, 4) = ".xls" Then
                If (InStr(1, Filename, "EOM", vbTextCompare) > 0) Then period = "EOM"
                If (InStr(1, Filename, "MID", vbTextCompare) > 0) Then period = "MID"
                If (InStr(1, Filename, "-ICB", vbTextCompare) > 0) Then
                    storenum = Mid(Filename, (InStr(1, Filename, "-ICB", vbTextCompare) - 4), 4)
                    monthin = Mid(Filename, (InStr(1, Filename, "-ICB", vbTextCompare) + 4), 2)
                    yearin = Mid(Filename, (InStr(1, Filename, "-ICB", vbTextCompare) + 6), 2)
                Else
                    storenum = "N/A"
                    monthin = "N/A"
                End If
                
                Do While ShiftPressed()
                    DoEvents
                Loop

                Workbooks.Open Filename, 0, True
                 
                Set xlsFiles = ActiveWorkbook
                 

                'ur = Split(xlsFiles.Sheets("DETAIL").Range("C:C").SpecialCells(xlCellTypeLastCell).Address(True, False), "$")(1)
                
                Dim colnum
                Set findcol = xlsFiles.Sheets("Detail").Range("A1:CC6").Find("Grand Total", LookIn:=xlValues, SearchFormat:=False)
                colnum = Split(findcol.Address(True, False), "$")(0)
                n = xlsFiles.Sheets("DETAIL").Range(colnum & ":" & colnum).Find("*", Range(colnum & "1"), SearchDirection:=xlPrevious).Row
                
                'COPY IMPORTANT COLUMNS ONE AT A TIME
                xlsFiles.Sheets("DETAIL").Range("C13:C" & n).Copy Destination:=wsMaster.Sheets("Output").Range("J" & r).Offset(1, 0)
                xlsFiles.Sheets("DETAIL").Range("D13:D" & n).Copy Destination:=wsMaster.Sheets("Output").Range("L" & r).Offset(1, 0)
               
                xlsFiles.Sheets("DETAIL").Range(colnum & "13:" & colnum & n).Copy Destination:=wsMaster.Sheets("Output").Range("M" & r).Offset(1, 0)
                
                
                
                For i = r + 1 To r + n - 12
                    wsMaster.Sheets("Output").Range("A" & i).Value = monthin
                    wsMaster.Sheets("Output").Range("B" & i).Value = period
                    wsMaster.Sheets("Output").Range("C" & i).Value = storenum
                    wsMaster.Sheets("Output").Range("Q" & i).Value = yearin
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
    
            Dim rng As Range, rngDelete As Range
            Set rng = .Range("J:J")
            .Cells.Sort key1:=rng, order1:=xlDescending
            .Rows(1).Insert (xlDown)
            .Cells.AutoFilter Field:=10, Criteria1:="="
            Set rngDelete = rng.SpecialCells(xlCellTypeVisible)
            rngDelete.EntireRow.Delete
            
            Set rng = .Range("A:A")
            .Cells.Sort key1:=rng, order1:=xlDescending
            .Rows(1).Insert (xlDown)
            .Cells.AutoFilter Field:=1, Criteria1:="N/A"
            Set rngDelete = rng.SpecialCells(xlCellTypeVisible)
            rngDelete.EntireRow.Delete
            
            n = .Range("J:J").Find("*", Range("J1"), SearchDirection:=xlPrevious).Row
            
            Dim cell
            Dim tmp
            For Each cell In .Range("J1:J" & n)
            tmp = Split(cell.Value, ")")
            If InStr(1, cell.Value, ")") > 0 Then
                cell.Offset(0, 1).Value = Trim(tmp(1))
            Else: cell.Offset(0, 1).Value = Trim(tmp(0))
            End If
            
            Next cell
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

