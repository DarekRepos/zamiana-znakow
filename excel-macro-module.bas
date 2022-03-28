Attribute VB_Name = "Module1"
Option Explicit

Sub OKButton_Click1(control As IRibbonControl)
    Dim SelectedTextCells As Range
    Dim cell As Range
    
    On Error Resume Next

    Set SelectedTextCells = Selection.SpecialCells(xlConstants, xlTextValues)
   
    Application.ScreenUpdating = False
    
   If Selection.Count = 1 Then
     If Not ActiveCell.HasFormula Then
     ActiveCell.Value = UCase(ActiveCell.Value)
     End If
   Else
     If Selection.MergeCells = True Then
      If Not ActiveCell.HasFormula Then
       ActiveCell.Value = UCase(ActiveCell.Value)
       End If
     Else
        For Each cell In SelectedTextCells
        cell.Value = UCase(cell.Value)
        Next
     End If
   End If
End Sub

Sub OKButton_Click2(control As IRibbonControl)
    Dim SelectedTextCells As Range
    Dim cell As Range
 
    On Error Resume Next
    Set SelectedTextCells = Selection.SpecialCells(xlConstants, xlTextValues)

    Application.ScreenUpdating = False
 
   If Selection.Count = 1 Then
     If Not ActiveCell.HasFormula Then
     ActiveCell.Value = LCase(ActiveCell.Value)
     End If
   Else
     If Selection.MergeCells = True Then
      If Not ActiveCell.HasFormula Then
       ActiveCell.Value = LCase(ActiveCell.Value)
       End If
     Else
        For Each cell In SelectedTextCells
        cell.Value = LCase(cell.Value)
        Next
     End If
   End If
End Sub

Sub OKButton_Click3(control As IRibbonControl)
    Dim SelectedTextCells As Range
    Dim cell As Range

    On Error Resume Next
    Set SelectedTextCells = Selection.SpecialCells(xlCellTypeConstants, xlTextValues)

    Application.ScreenUpdating = False
    
    If Selection.Count = 1 Then
     If Not ActiveCell.HasFormula Then
        ActiveCell.Value = Application.WorksheetFunction.Proper(ActiveCell.Value)
     End If
    Else
    If Selection.MergeCells = True Then
      If Not ActiveCell.HasFormula Then
       ActiveCell.Value = Application.WorksheetFunction.Proper(ActiveCell.Value)
       End If
     Else
        For Each cell In SelectedTextCells
        cell.Value = Application.WorksheetFunction.Proper(cell.Value)
        Next
     End If
    End If
End Sub

Sub OKButton_Click4(control As IRibbonControl)
    Dim SelectedTextCells As Range
    Dim cell As Range
    Dim Text As String

    On Error Resume Next
    Set SelectedTextCells = Selection.SpecialCells(xlConstants, xlTextValues)

    Application.ScreenUpdating = False

    If Selection.Count = 1 Then
    If Not ActiveCell.HasFormula Then
        Text = ActiveCell.Value
        Text = UCase(Left(ActiveCell.Value, 1))
        Text = Text & LCase(Mid(ActiveCell.Value, 2, Len(ActiveCell.Value)))
        ActiveCell.Value = Text
        End If
    Else
    If Selection.MergeCells = True Then
      If Not ActiveCell.HasFormula Then
        Text = ActiveCell.Value
        Text = UCase(Left(ActiveCell.Value, 1))
        Text = Text & LCase(Mid(ActiveCell.Value, 2, Len(ActiveCell.Value)))
        ActiveCell.Value = Text
       End If
     Else
    For Each cell In SelectedTextCells
        Text = cell.Value
        Text = UCase(Left(cell.Value, 1))
        Text = Text & LCase(Mid(cell.Value, 2, Len(cell.Value)))
        cell.Value = Text
    Next
    End If
    End If
End Sub

Sub OKButton_Click5(control As IRibbonControl)
    Dim SelectedTextCells As Range
    Dim cell As Range
    Dim Text As String
    Dim i As Long

    On Error Resume Next
    Set SelectedTextCells = Selection.SpecialCells(xlConstants, xlTextValues)
    
    Application.ScreenUpdating = False

    If Selection.Count = 1 Then
     If Not ActiveCell.HasFormula Then
      Text = ActiveCell.Value
            For i = 1 To Len(Text)
              If Mid(Text, i, 1) Like "[A-Z]" Then
                 Mid(Text, i, 1) = LCase(Mid(Text, i, 1))
              Else
                 Mid(Text, i, 1) = UCase(Mid(Text, i, 1))
              End If
            Next i
      ActiveCell.Value = Text
      End If
    Else
         If Selection.MergeCells = True Then
      If Not ActiveCell.HasFormula Then
             Text = ActiveCell.Value
            For i = 1 To Len(Text)
              If Mid(Text, i, 1) Like "[A-Z]" Then
                 Mid(Text, i, 1) = LCase(Mid(Text, i, 1))
              Else
                 Mid(Text, i, 1) = UCase(Mid(Text, i, 1))
              End If
            Next i
      ActiveCell.Value = Text
       End If
         Else
       For Each cell In SelectedTextCells
        Text = cell.Value
            For i = 1 To Len(Text)
              If Mid(Text, i, 1) Like "[A-Z]" Then
                 Mid(Text, i, 1) = LCase(Mid(Text, i, 1))
              Else
                 Mid(Text, i, 1) = UCase(Mid(Text, i, 1))
              End If
            Next i
        cell.Value = Text
    Next
     End If
    
    
    End If
End Sub


