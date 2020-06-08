Attribute VB_Name = "Tester"

Sub GenerateTable1()
    '
    ' GenerateTable1 Macro
    '
    '
    ActiveDocument.Tables.Add Range:=Selection.Range, NumRows:=6, NumColumns:= _
    8, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
    wdAutoFitFixed
    With Selection.Tables(1)
        If .Style <> "Table Grid" Then
            .Style = "Table Grid"
        End If
        .ApplyStyleHeadingRows = True
        .ApplyStyleLastRow = False
        .ApplyStyleFirstColumn = True
        .ApplyStyleLastColumn = False
        .ApplyStyleRowBands = True
        .ApplyStyleColumnBands = False
    End With
End Sub





Sub CreateNewTable()
    Dim docActive   As Document
    Dim tblNew      As Table
    Dim celTable    As Cell
    Dim i    As Integer
    Dim colNames
    
    Set docActive = ActiveDocument
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=0, End:=0), NumRows:=1, _
        NumColumns:=9)
        
    With tblNew
        .Style = "Table Grid"
        .ApplyStyleHeadingRows = True
        .ApplyStyleLastRow = False
        .ApplyStyleFirstColumn = True
        .ApplyStyleLastColumn = False
        .ApplyStyleRowBands = True
        .ApplyStyleColumnBands = False
        End With
    colNames = Array("Test Number", "Description", "Test Data", "Test Type", "Expected Value", "Actual Value", "Pass/Fail", "Cross reference", "func_name")
     
    i = 0
    
    For Each celTable In tblNew.Rows(1).Cells
        celTable.Range.Text = colNames(i)
        i = i + 1
    Next celTable
    
    tblNew.Columns(9).Width = CentimetersToPoints(0.42)
    tblNew.Columns(9).Borders(wdBorderRight).LineStyle = wdLineStyleNone
    tblNew.Columns(9).Borders(wdBorderTop).LineStyle = wdLineStyleNone
    tblNew.Columns(9).Borders(wdBorderBottom).LineStyle = wdLineStyleNone
    For Each celTable In tblNew.Columns(9).Cells
        celTable.Range.Font.Hidden = True
        celTable.Range.Font.Size = 1
    Next celTable
    
End Sub

Sub InsertTextInCell()
    Dim celTable    As Cell
    Dim tblNew      As Table
    If ActiveDocument.Tables.Count >= 1 Then
        Set tblNew = ActiveDocument.Tables(1)
        For Each celTable In tblNew.Range.Cells
            With celTable.Range
                .Delete
                .InsertAfter Text:="Cell 1,1"
            End With
        Next celTable
    End If
End Sub

Sub CreateTestTableFromCSV()

    Dim FilePath As String
    Dim LineFromFile As String
    Dim oNewRow
    Dim oTable As Table
    Dim data As String
    
    Dim oPara1 As Word.Paragraph
    Dim oParal2 As Word.Paragraph
    
    Dim oDoc As Word.Document
    Set oDoc = ActiveDocument
    
    Call CreateNewTable

    Set oTable = ActiveDocument.Tables(1)
    
    FilePath = "C:\Users\timmc\source\repos\PasswordChecker\PasswordChecker\testOutput\testResults.csv"
    row_number = 0
    Open FilePath For Input As #1
    Do Until EOF(1)
        Line Input #1, LineFromFile
        LineItems = Split(LineFromFile, ",")
        oTable.Rows.Add
        Set oNewRow = oTable.Rows(oTable.Rows.Count)
        oNewRow.Cells(1).Range.Text = LineItems(1)
        oNewRow.Cells(2).Range.Text = LineItems(2)
        oNewRow.Cells(3).Range.Text = LineItems(3)
        oNewRow.Cells(4).Range.Text = LineItems(4)
        oNewRow.Cells(5).Range.Text = LineItems(5)
        oNewRow.Cells(6).Range.Text = LineItems(6)
        oNewRow.Cells(7).Range.Text = LineItems(7)
        oNewRow.Cells(8).Range.Text = "Screenshot below " + LineItems(1)
        
        Set oPara1 = oDoc.Content.Paragraphs.Add
        
        With oPara1.Range
            .Style = ActiveDocument.Styles("Subtitle")
            .Text = "Test Case " + LineItems(1) + " Evidence"
        End With
        
        Set oPara12 = oDoc.Content.Paragraphs.Add
        oPara1.Range.InlineShapes.AddPicture FileName:= _
        LineItems(8), LinkToFile:=False, _
        SaveWithDocument:=True
        
            
    Loop
    
    Close #1
    

End Sub
