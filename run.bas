' This module contains a VBA subroutine that reads the value of cell B3 in the script worksheet,
' and calls one of two subroutines ("KS01_mass" or "KS02_mass") depending on the value.
'
' If the value in B3 is "KS01", the subroutine "KS01_mass" is called to run a mass script for creating cost centers.
' If the value in B3 is "KS02", the subroutine "KS02_mass" is called to run a mass script for modifying cost centers.
'
' To use this module, create a button or other user interface element in the Excel worksheet, and assign the "runMassScript" subroutine
' to be called when the button is clicked or the element is activated. Make sure that the worksheet contains a dropdown list in cell B3
' with the options "KS01" and "KS02" for selecting the desired mass update script.
'
' Author: Abel Tavares


Sub runMassScript()
    ' Declare a variable to hold the selected option
    Dim selectedOption As String
    
    ' Read the value of cell B3 and store it in the variable "selectedOption"
    selectedOption = Range("B3").Value
    
    ' Use a Select Case statement to check the value of "selectedOption" and call the appropriate subroutine
    Select Case selectedOption
        Case "KS01"
            KS01_mass ' Call the KS01_mass subroutine if "KS01" is selected
        Case "KS02"
            KS02_mass ' Call the KS02_mass subroutine if "KS02" is selected
        Case Else
            MsgBox "Invalid option selected. Please choose KS01 or KS02." ' Show a message box if an invalid option is selected
    End Select
End Sub