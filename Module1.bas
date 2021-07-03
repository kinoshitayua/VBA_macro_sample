Attribute VB_Name = "Module1"
Option Explicit

Sub sample1()

    Application.ScreenUpdating = False
    
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
        
    Dim path As String
    Dim file As String
    Dim wb As Workbook
    
    path = ActiveWorkbook.path
    file = ThisWorkbook.Worksheets("").Range("")
    
    Set wb = Workbooks.Add
    wb.SaveAs Filename:=path & "\" & file
    ThisWorkbook.Activate

    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    
    If Range("") = "New" Then
        Worksheets("").Cells.Copy
        wb.Worksheets("").Range("").PasteSpecial
        Worksheets("").Cells.Copy
        wb.Worksheets("").Range("").PasteSpecial Paste:=xlPasteValues
        
    ElseIf Range("") = "Old" Then
        Worksheets("").Cells.Copy
        wb.Worksheets("").Range("").PasteSpecial
        Worksheets("").Cells.Copy
        wb.Worksheets("").Range("").PasteSpecial Paste:=xlPasteValues
        
    End If
    
    Application.CutCopyMode = False
    
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

    Dim rowmax As Long
    Dim i As Long

    rowmax = Cells(Rows.Count, "A").End(xlUp).row
    For i = rowmax To 2 Step -1
'    For i = 4 To rowmax
        With wb.Worksheets("")
                .Activate
                
                If .Cells(i, "G") = "-" Then
                    .Cells(i, "G").EntireRow.Delete
                End If
    
                    If Range("") = "New" Then
                        If .Cells(i, "A") = "Old" Then
                            .Cells(i, "A").EntireRow.Delete
                         End If
                
                    ElseIf Range("") = "Old" Then
                        If .Cells(i, "A") = "New" Then
                            .Cells(i, "A").EntireRow.Delete
                        End If
                    
                    End If
        
        End With
    
    Next i
    
    Columns("A").Delete

    '**************************************************************************
    
    wb.Save
    wb.Close
    ThisWorkbook.Worksheets("").Activate

    Application.ScreenUpdating = True

    Dim rc As Integer
    rc = MsgBox("", vbYesNo)
    If rc = vbYes Then
    Workbooks.Open path & "\" & file
    End If
    
End Sub
