' This file contains macros to configure the Excel interface settings and add a drop-down list to the worksheet

' This subroutine is run when the workbook is opened to configure the Excel interface settings
Private Sub Workbook_Open()
    ' Hide the Ribbon toolbar
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
    
    ' Hide the Formula Bar
    Application.DisplayFormulaBar = False
    
    ' Toggle the display of the Status Bar
    Application.DisplayStatusBar = Not Application.DisplayStatusBar
    
    ' Hide workbook tabs
    ActiveWindow.DisplayWorkbookTabs = False
End Sub

' This subroutine adds a drop-down list with two options ("KS01" and "KS02") to cell B3 of the active worksheet
Sub AddDropdown()
    ' Define the cell where the drop-down list will be added
    Dim cell As Range
    Set cell = Range("B3")

    ' Create the drop-down list data
    Dim data(1) As String
    data(0) = "KS01"
    data(1) = "KS02"

    ' Add the drop-down list to the cell
    With cell.Validation
        .Delete ' remove any existing validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=Join(data, ",")
    End With
End Sub

