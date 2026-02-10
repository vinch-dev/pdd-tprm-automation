' Logic: Late-bound SAP GUI automation with dynamic parameter injection
Sub Execute_SAP_Extraction_Framework()
    Dim SapGuiAuto As Object, App As Object, Connection As Object, Session As Object
    Dim ws As Worksheet: Set ws = ActiveWorkbook.Sheets("Automation_Control")
    
    ' Localized variables to prevent hardcoding paths
    Dim folderPath As String: folderPath = ws.Range("B10").Value
    Dim fileName As String: fileName = "Audit_Export_" & Format(Now, "HHmm") & ".xlsx"
    
    On Error GoTo ErrorHandler
    
    ' Attachment to existing SAP Session (Crucial for ground-level stability)
    Set SapGuiAuto = GetObject("SAPGUI")
    Set App = SapGuiAuto.GetScriptingEngine
    Set Connection = App.Children(0)
    Set Session = Connection.Children(0)

    ' Transaction Navigation
    Session.findById("wnd[0]/tbar[0]/okcd").Text = "Z_ERP_REPORT_CODE"
    Session.findById("wnd[0]").sendVKey 0

    ' Handling Multi-Value Selection Popups (Your specific logic)
    Session.findById("wnd[0]/usr/btn%_SELECTION_PUSH").press
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtFIELD[1,0]").Text = ws.Range("B8").Value
    Session.findById("wnd[1]").sendVKey 8 ' Confirm Multiple Selection

    ' Buffer for SAP Rendering (Prevents script from outrunning the GUI)
    Application.Wait (Now + TimeValue("0:00:05"))
    Session.findById("wnd[0]").sendVKey 8 ' Execute Report

    ' Export Sequence with specific SAP Layout Keys
    Session.findById("wnd[0]").sendVKey 33 ' Select Layout
    Session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cmbBOX").Key = "X"
    
    Session.findById("wnd[0]").sendVKey 43 ' Local File Export
    Session.findById("wnd[1]/usr/ctxtDY_PATH").Text = folderPath
    Session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = fileName
    Session.findById("wnd[1]").sendVKey 7 ' Start Download

    ' Trigger Python Transformation with CLI Arguments
    Dim cmd As String: cmd = "python.exe ""C:\Scripts\processor.py"" """ & folderPath & "\" & fileName & """"
    CreateObject("WScript.Shell").Run cmd, 1, False

CleanExit:
    Set Session = Nothing: Set App = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Automation Halted: " & Err.Description & vbCrLf & "Check SAP Scripting permissions.", vbCritical
End Sub