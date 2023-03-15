' Mass Maintenance for Tcode KS02 in SAP
' This script automates mass maintenance for cost centers in SAP using tcode KS02. It loops through a list of cost centers and updates the specified fields, which are defined in the worksheet. 
' If mandatory fields are not filled, the script will prompt the user to fill them before continuing. The results are logged in column A of the worksheet.
'
' Author: Abel Tavares
'
' Instructions:
' 1. Fill out the System Name in the first row of the worksheet.
' 2. Fill out the fields to be updated for each cost center in the respective columns.
' 3. Run the script by clicking the "Run Script" button.
' 4. Wait for the script to finish running. A message box will appear when the script has finished processing all cost centers.
' 5. Check the log in column A for any errors or warnings.
'
' Note: This script requires SAP GUI to be installed on the machine.


Option Explicit


Sub KS02_mass()


    ' Variables
    Dim sapGui As Object
    Dim applic As Object
    Dim connection As Object
    Dim session As Object
    Dim systemName As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim row As Long
    Dim t As Long
    

    ' Constants
    Const SAPLOGON_PATH As String = "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"

    ' Set reference to active workbook and worksheet
    Set ws = ActiveSheet

    ' Get system name from the worksheet
    systemName = ws.Cells(1, "B").Value


    ' Validate mandatory fields
    If systemName = "" Then
        MsgBox ("Please fill System Name")
        Exit Sub
    End If

    ' Get the last cost center row
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row

    ' Loop through each cost center and check if mandatory dates are filled
    For row = 7 To lastRow
        If Not IsEmpty(ws.Cells(row, "B")) And _
        (IsEmpty(ws.Cells(row, "C")) Or IsEmpty(ws.Cells(row, "D"))) Then 'validFromDate and validToDate
            MsgBox ("Please fill mandatory dates for cost center " & ws.Cells(row, "B").Value) & " change"
            Exit Sub
        End If
    Next row


    ' Check if SAP Logon is already running
    On Error Resume Next
    Set sapGui = GetObject("SAPGUI")
    On Error GoTo 0
    
    'Handle errors
    On Error GoTo ErrorHandler

    ' If SAP Logon is not running, start it
    If sapGui Is Nothing Then
        Shell SAPLOGON_PATH, vbHide
        Application.Wait Now + TimeValue("0:00:03")
        Set sapGui = GetObject("SAPGUI")
    End If

    ' Connect to SAP and get the session
    Set applic = sapGui.GetScriptingEngine
    Set connection = applic.OpenConnection(systemName, True)
    Set session = connection.Children(0)


    ' Maximize SAP window
    session.findById("wnd[0]").maximize

    ' Hide SAP window
    session.findById("wnd[0]").iconify

    ' Navigate to ks02 transaction
    session.findById("wnd[0]/tbar[0]/okcd").Text = "ks02"
    session.findById("wnd[0]").sendVKey 0

    ' Loop through the data rows and perform operations on SAP
    t = 7 ' start at row 7
    Do Until IsEmpty(ws.Cells(t, 2).Value)

        'Enter Cost Center ID
        session.findById("wnd[0]/usr/ctxtCSKSZ-KOSTL").Text = ws.Cells(t, 2).Value 'Cost Center
        session.findById("wnd[0]").sendVKey 0
        
        If session.ActiveWindow.name = "wnd[1]" Then
            session.findById("wnd[1]").sendVKey 0
        End If

        'Open analysis period
        session.findById("wnd[0]/mbar/menu[1]/menu[0]").Select
        session.findById("wnd[1]/tbar[0]/btn[6]").press
        
        'Enter Valid From and Valid To dates
        session.findById("wnd[2]/usr/ctxtRKMA2-DATAB").Text = ws.Cells(t, 3).Value 'Valid from
        session.findById("wnd[2]/usr/ctxtRKMA2-DATBI").Text = ws.Cells(t, 4).Value  'Valid to
        session.findById("wnd[2]/tbar[0]/btn[0]").press

        ' Perform change operations
        ' Basic data tab
        If Not IsEmpty(ws.Cells(t, 5).Value) Then
        session.findById("wnd[0]/usr/tabsTABSTRIP_EINZEL/tabpGRUN/ssubSUBSCREEN_EINZEL:SAPLKMA1:0300/txtCSKSZ-KTEXT").Text = ws.Cells(t, 5).Value 'Name
        End If

        If Not IsEmpty(ws.Cells(t, 6).Value) Then
        session.findById("wnd[0]/usr/tabsTABSTRIP_EINZEL/tabpGRUN/ssubSUBSCREEN_EINZEL:SAPLKMA1:0300/txtCSKSZ-LTEXT").Text = ws.Cells(t, 6).Value 'Description
        End If
        
        If Not IsEmpty(ws.Cells(t, 7).Value) Then
        session.findById("wnd[0]/usr/tabsTABSTRIP_EINZEL/tabpGRUN/ssubSUBSCREEN_EINZEL:SAPLKMA1:0300/ctxtCSKSZ-VERAK_USER").Text = ws.Cells(t, 7).Value 'User
        End If
        
        If Not IsEmpty(ws.Cells(t, 8).Value) Then
        session.findById("wnd[0]/usr/tabsTABSTRIP_EINZEL/tabpGRUN/ssubSUBSCREEN_EINZEL:SAPLKMA1:0300/txtCSKSZ-VERAK").Text = ws.Cells(t, 8).Value 'Person
        End If
        
        If Not IsEmpty(ws.Cells(t, 9).Value) Then
        session.findById("wnd[0]/usr/tabsTABSTRIP_EINZEL/tabpGRUN/ssubSUBSCREEN_EINZEL:SAPLKMA1:0300/ctxtCSKSZ-KOSAR").Text = ws.Cells(t, 9).Value 'Category
        End If

        If Not IsEmpty(ws.Cells(t, 10).Value) Then
        session.findById("wnd[0]/usr/tabsTABSTRIP_EINZEL/tabpGRUN/ssubSUBSCREEN_EINZEL:SAPLKMA1:0300/ctxtCSKSZ-KHINR").Text = ws.Cells(t, 10).Value 'Hierarchy
        End If
        
        If Not IsEmpty(ws.Cells(t, 12).Value) Then
        session.findById("wnd[0]/usr/tabsTABSTRIP_EINZEL/tabpGRUN/ssubSUBSCREEN_EINZEL:SAPLKMA1:0300/ctxtCSKSZ-BUKRS").Text = ws.Cells(t, 12).Value 'Company Code
        End If
        
        If Not IsEmpty(ws.Cells(t, 14).Value) Then
        session.findById("wnd[0]/usr/tabsTABSTRIP_EINZEL/tabpGRUN/ssubSUBSCREEN_EINZEL:SAPLKMA1:0300/ctxtCSKSZ-PRCTR").Text = ws.Cells(t, 14).Value 'Profit Center
        session.findById("wnd[0]").sendVKey 0
        End If
        
        If Not IsEmpty(ws.Cells(t, 11).Value) Then
        session.findById("wnd[0]/usr/tabsTABSTRIP_EINZEL/tabpGRUN/ssubSUBSCREEN_EINZEL:SAPLKMA1:0300/ctxtCSKSZ-FUNC_AREA").Text = ws.Cells(t, 11).Value 'Functional
        End If

        If Not IsEmpty(ws.Cells(t, 13).Value) Then
            session.findById("wnd[0]/usr/tabsTABSTRIP_EINZEL/tabpGRUN/ssubSUBSCREEN_EINZEL:SAPLKMA1:0300/ctxtCSKSZ-GSBER").Text = ws.Cells(t, 13).Value 'Business Area
        End If
        
        ' Select Control tab
        session.findById("wnd[0]/usr/tabsTABSTRIP_EINZEL/tabpKZEI").Select

        'Check if warning box pops up and confirm the changes in Basic data tab
        If session.ActiveWindow.name = "wnd[1]" Then
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        End If

        If Not IsEmpty(ws.Cells(t, 15).Value) Then
        session.findById("wnd[0]/usr/tabsTABSTRIP_EINZEL/tabpKZEI/ssubSUBSCREEN_EINZEL:SAPLKMA1:0310/chkCSKSZ-MGEFL").Selected = True 'Record Quantity
        End If

        If Not IsEmpty(ws.Cells(t, 16).Value) Then
        session.findById("wnd[0]/usr/tabsTABSTRIP_EINZEL/tabpKZEI/ssubSUBSCREEN_EINZEL:SAPLKMA1:0310/chkCSKSZ-BKZKP").Selected = True 'Lock Actual Primary Costs
        End If

        If Not IsEmpty(ws.Cells(t, 17).Value) Then
        session.findById("wnd[0]/usr/tabsTABSTRIP_EINZEL/tabpKZEI/ssubSUBSCREEN_EINZEL:SAPLKMA1:0310/chkCSKSZ-BKZKS").Selected = True 'Lock Actual Secondary Costs
        End If

        If Not IsEmpty(ws.Cells(t, 18).Value) Then
        session.findById("wnd[0]/usr/tabsTABSTRIP_EINZEL/tabpKZEI/ssubSUBSCREEN_EINZEL:SAPLKMA1:0310/chkCSKSZ-BKZER").Selected = True 'Lock Actual Revenues
        End If

        If Not IsEmpty(ws.Cells(t, 19).Value) Then
        session.findById("wnd[0]/usr/tabsTABSTRIP_EINZEL/tabpKZEI/ssubSUBSCREEN_EINZEL:SAPLKMA1:0310/chkCSKSZ-PKZKP").Selected = True 'Lock Plan Primary Costs
        End If

        If Not IsEmpty(ws.Cells(t, 20).Value) Then
        session.findById("wnd[0]/usr/tabsTABSTRIP_EINZEL/tabpKZEI/ssubSUBSCREEN_EINZEL:SAPLKMA1:0310/chkCSKSZ-PKZKS").Selected = True 'Lock Plan Secondary Costs
        End If

        If Not IsEmpty(ws.Cells(t, 21).Value) Then
            session.findById("wnd[0]/usr/tabsTABSTRIP_EINZEL/tabpKZEI/ssubSUBSCREEN_EINZEL:SAPLKMA1:0310/chkCSKSZ-PKZER").Selected = True 'Lock Plan Revenues
        End If

        If Not IsEmpty(ws.Cells(t, 22).Value) Then
            session.findById("wnd[0]/usr/tabsTABSTRIP_EINZEL/tabpKZEI/ssubSUBSCREEN_EINZEL:SAPLKMA1:0310/chkCSKSZ-BKZOB").Selected = True 'Lock Commitment Update
        End If

        ' Select Templates tab
        session.findById("wnd[0]/usr/tabsTABSTRIP_EINZEL/tabpTMPT").Select

        If Not IsEmpty(ws.Cells(t, 23).Value) Then
        session.findById("wnd[0]/usr/tabsTABSTRIP_EINZEL/tabpTMPT/ssubSUBSCREEN_EINZEL:SAPLKMA1:0350/ctxtCSKSZ-KALSM").Text = ws.Cells(t, 23).Value 'Costing Sheet
        End If

        ' Select Adress tab
        session.findById("wnd[0]/usr/tabsTABSTRIP_EINZEL/tabpADRE").Select

        If Not IsEmpty(ws.Cells(t, 24).Value) Then
            session.findById("wnd[0]/usr/tabsTABSTRIP_EINZEL/tabpADRE/ssubSUBSCREEN_EINZEL:SAPLKMA1:0320/ctxtCSKSZ-LAND1").Text = ws.Cells(t, 24).Value 'Country
        End If

        ' Select Cummunication tab
        session.findById("wnd[0]/usr/tabsTABSTRIP_EINZEL/tabpKOMM").Select

        If Not IsEmpty(ws.Cells(t, 25).Value) Then
            session.findById("wnd[0]/usr/tabsTABSTRIP_EINZEL/tabpKOMM/ssubSUBSCREEN_EINZEL:SAPLKMA1:0330/ctxtCSKSZ-SPRAS").Text = ws.Cells(t, 25).Value 'Language Key
        End If

        ' Select Add.fields tab
        session.findById("wnd[0]/usr/tabsTABSTRIP_EINZEL/tabp+CU1").Select

        If Not IsEmpty(ws.Cells(t, 26).Value) Then
        session.findById("wnd[0]").sendVKey 4
        session.findById("wnd[1]").iconify
        session.findById("wnd[1]/tbar[0]/btn[17]").press
        session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[0,24]").Text = ws.Cells(t, 27).Value 'Plant
        session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[1,24]").Text = ws.Cells(t, 26).Value 'Location
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        End If

        session.findById("wnd[0]/tbar[0]/btn[11]").press
        session.findById("wnd[0]").sendVKey 0
        
        ' Run Log
        ws.Cells(t, 1).Value = "Sucess"

        t = t + 1

    Loop


    ' Close SAP GUI and release all SAP-related objects
    session.findById("wnd[0]").Close
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    Set session = Nothing
    Set connection = Nothing
    Set applic = Nothing
    Set sapGui = Nothing

    MsgBox "Macro has finished processing: " & t - 7 & " cost centers changed."

    Exit Sub


ErrorHandler:
    ' Run Log
    ws.Cells(t, 1).Value = "Failed - Error " & Err.Number & ": " & Err.Description
    'Close SAP
    session.findById("wnd[0]").Close
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    Set session = Nothing
    Set connection = Nothing
    Set applic = Nothing
    Set sapGui = Nothing
    ' Give error
    MsgBox "Error " & Err.Number & ": " & Err.description
    Exit Sub
    

End Sub