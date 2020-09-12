Attribute VB_Name = "RemoveAllFormatting"
Sub OneSubToRuleThemAll()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Dim wkSheet As Worksheet
    Dim deletedTabCount As Integer
    
    deletedTabCount = 0

    For Each wkSheet In Worksheets
        RemoveImages
        Select Case wkSheet.Name
            Case "Epsilon- TotalSource Plus"
                wkSheet.Activate
                Call FormatTSP
             Case "Epsilon- MarketTrends"
                wkSheet.Activate
                Call FormatMarketTrends
            Case "Epsilon- Online Behavioral"
                wkSheet.Activate
                Call FormatOnlineBehavioral
            Case "Epsilon- ShoppersVoice"
                wkSheet.Activate
                Call FormatShoppersVoice
            Case "Epsilon-MarketView"
                wkSheet.Activate
                Call FormatMarketView
            Case "Epsilon- Contextual Labels"
                wkSheet.Activate
                Call FormatContextualLabels
            Case "Inscape"
                wkSheet.Activate
                Call FormatInscape
            Case Else
                wkSheet.Delete
                deletedTabCount = deletedTabCount + 1
        End Select
        ClearFormatting
    Next wkSheet
        
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox deletedTabCount & " tabs deleted"
        

End Sub

Private Sub RemoveImages()

    Dim shape As Excel.shape
    For Each shape In ActiveSheet.Shapes
        shape.Delete
    Next


End Sub

Private Sub ClearFormatting()

    ActiveSheet.UsedRange.ClearFormats

End Sub

Private Sub RemoveMergedCells(ByVal rowIndex)
Attribute RemoveMergedCells.VB_ProcData.VB_Invoke_Func = " \n14"


    With ActiveSheet
        If .Cells(rowIndex, 1).MergeCells Then
            .Cells(rowIndex, 1).EntireRow.Delete
        End If
    End With
    
End Sub

Private Sub DeleteAllEmptyRows(ByVal rowIndex)

    With ActiveSheet
        If Application.CountA(Rows(rowIndex)) = 0 Then
            Rows(rowIndex).Delete
        End If
    End With
 
End Sub

Private Sub CopyValuesIntoBlankCells(ByVal rowIndex, colIndex)
      
    With ActiveSheet
        If .Cells(rowIndex, colIndex).Value = "" Then
            .Cells(rowIndex, colIndex).Value = .Cells(rowIndex - 1, colIndex).Value
        End If
    End With
        
End Sub

Private Sub ConvertTableToRange()

    ActiveSheet.ListObjects(1).Unlist
    
End Sub

Private Sub AppendBleedingCell(ByVal rowIndex)

    With ActiveSheet
        If .Cells(rowIndex, 1).Value = "" And .Cells(rowIndex, 2).Value <> "" Then
          .Cells(rowIndex - 1, 2).Value = .Cells(rowIndex - 1, 2).Value & .Cells(rowIndex, 2).Value
          .Cells(rowIndex, 2).ClearContents
        End If
    End With
 
End Sub

Private Sub CountDelimiters(ByVal rowIndex, delimMax)
    
    With ActiveSheet
        myTxt = .Cells(rowIndex, "A").Value
        delimCount = (Len(myTxt) - Len(Replace(myTxt, ">", "")))
            If delimCount > delimMax Then
                delimMax = delimCount
            End If
    End With

End Sub

Private Sub InsertHeaders(ByVal delimMax)

 'NOTE:Insert columns
    lastInsertColIndex = delimMax + 1
    firstInsertColIndex = 2
    Range(Columns(firstInsertColIndex), Columns(lastInsertColIndex)).Insert
    
    'NOTE:Update Headers
    rowIndex = 1
    colIndex = 1
    columnName = ActiveSheet.Cells(rowIndex, colIndex).Value
    For colIndex = 1 To lastInsertColIndex Step 1
        ActiveSheet.Cells(rowIndex, colIndex).Value = columnName & " " & colIndex
    Next colIndex

End Sub

Private Sub SplitCell(ByVal rowIndex)

    
    If Cells(rowIndex, 1) <> "" Then
         Cells(rowIndex, 1).TextToColumns _
          Destination:=ActiveSheet.UsedRange.Cells(rowIndex, 1), _
          DataType:=xlDelimited, _
          TextQualifier:=xlDoubleQuote, _
          ConsecutiveDelimiter:=False, _
          Tab:=True, _
          Semicolon:=False, _
          Comma:=False, _
          Space:=False, _
          Other:=True, _
          OtherChar:=">"
    End If
    

End Sub

Private Sub DeleteRowIfEmptyFirstCell(ByVal rowIndex)
    
    With ActiveSheet
        If .Cells(rowIndex, 1).Value = "" Then
            Rows(rowIndex).Delete
        End If
    End With
  
End Sub

Private Sub RemoveNAs(ByVal rowIndex, colIndex)
    
    With ActiveSheet
        If .Cells(rowIndex, colIndex).Value = "N/A" Then
            Cells(rowIndex, colIndex).Clear
        End If
    End With
   
End Sub

Private Sub FormatTSP()

    Dim lastIndex As New lastIndex
      
    
    lastIndex.lastRow = ActiveSheet.UsedRange.Rows.Count
    lastIndex.lastColumn = ActiveSheet.UsedRange.Columns.Count
    
    For rowIndex = lastIndex.lastRow To 1 Step -1
        RemoveMergedCells rowIndex
        DeleteAllEmptyRows rowIndex
    Next rowIndex
    
    For rowIndex = 1 To lastIndex.lastRow Step 1
        For colIndex = 1 To lastIndex.lastColumn Step 1
            CopyValuesIntoBlankCells rowIndex, colIndex
        Next colIndex
    Next rowIndex

End Sub

Private Sub FormatMarketTrends()

    Dim lastIndex As New lastIndex
      
    
    lastIndex.lastRow = ActiveSheet.UsedRange.Rows.Count
    lastIndex.lastColumn = ActiveSheet.UsedRange.Columns.Count

    ConvertTableToRange
    
    For rowIndex = lastIndex.lastRow To 1 Step -1
        RemoveMergedCells rowIndex
        If rowIndex <> 1 Then
            AppendBleedingCell rowIndex
        End If
        DeleteAllEmptyRows rowIndex
    Next rowIndex

    
End Sub

Private Sub FormatOnlineBehavioral()
    
    Dim lastIndex As New lastIndex
      
    
    lastIndex.lastRow = ActiveSheet.UsedRange.Rows.Count
    lastIndex.lastColumn = ActiveSheet.UsedRange.Columns.Count
    
    With ActiveSheet
        For rowIndex = lastIndex.lastRow To 1 Step -1
            'HACK:custom code for the stupid formatting of this page
            If .Cells(rowIndex, 1).Value = "Segments" And .Cells(rowIndex, 1).MergeCells Then
                .Cells(rowIndex, 1).MergeArea.UnMerge
            End If
            'NOTE:end custom code
            RemoveMergedCells rowIndex
            DeleteAllEmptyRows rowIndex
        Next rowIndex
    End With

End Sub

Private Sub FormatShoppersVoice()
    
    'NOTE: for testing
    'Application.ScreenUpdating = False
    'Application.DisplayAlerts = False

    Dim lastIndex As New lastIndex
    Dim delimMax As Integer
          
    
    lastIndex.lastRow = ActiveSheet.UsedRange.Rows.Count
    lastIndex.lastColumn = ActiveSheet.UsedRange.Columns.Count
    
    delimMax = 0
    
    For rowIndex = lastIndex.lastRow To 1 Step -1
        RemoveMergedCells rowIndex
        DeleteAllEmptyRows rowIndex
        CountDelimiters rowIndex, delimMax
    Next rowIndex
    
    Debug.Print delimMax

    InsertHeaders delimMax
    
    For rowIndex = 2 To lastIndex.lastRow Step 1
        SplitCell rowIndex
    Next rowIndex

End Sub

Function FormatMarketView()

    'NOTE: for testing
    'Application.ScreenUpdating = False
    'Application.DisplayAlerts = False

    Dim lastIndex As New lastIndex
    Dim delimMax As Integer
          
    
    lastIndex.lastRow = ActiveSheet.UsedRange.Rows.Count
    lastIndex.lastColumn = ActiveSheet.UsedRange.Columns.Count
    
    delimMax = 0
    
    For rowIndex = lastIndex.lastRow To 1 Step -1
        RemoveMergedCells rowIndex
        DeleteAllEmptyRows rowIndex
        CountDelimiters rowIndex, delimMax
    Next rowIndex
    
    Debug.Print delimMax

    InsertHeaders delimMax
    
    For rowIndex = lastIndex.lastRow To 1 Step -1
        SplitCell rowIndex
    Next rowIndex

End Function

Function FormatContextualLabels()

    Dim lastIndex As New lastIndex
          
    
    lastIndex.lastRow = ActiveSheet.UsedRange.Rows.Count
    lastIndex.lastColumn = ActiveSheet.UsedRange.Columns.Count
    
    For rowIndex = lastIndex.lastRow To 1 Step -1
        RemoveMergedCells rowIndex
        DeleteAllEmptyRows rowIndex
        DeleteRowIfEmptyFirstCell rowIndex
            For colIndex = 1 To lastIndex.lastColumn Step 1
                RemoveNAs rowIndex, colIndex
            Next colIndex
    Next rowIndex

End Function

Function FormatInscape()

    Dim lastIndex As New lastIndex
          
    
    lastIndex.lastRow = ActiveSheet.UsedRange.Rows.Count
    lastIndex.lastColumn = ActiveSheet.UsedRange.Columns.Count
    
    For rowIndex = lastIndex.lastRow To 1 Step -1
        RemoveMergedCells rowIndex
        DeleteAllEmptyRows rowIndex
    Next rowIndex

End Function
