' SAP
SAP_BIN = "saplogon.exe"
SAP_GUI_PATH = "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\" & SAP_BIN
' Connections
lk0_sso = "1. Grupo Nutresa_ERP_PRD"
lk0_auth = "LK0 [EHP6 SP7 SIC Domestico - TESTING]"
' Auth
user = "username"
password = "pw"

Fecha_Entrega=date() -1
findString  = "/"
replaceWith = "."
startPos = 1
Fecha1 = Replace(Fecha_Entrega,findString,replaceWith,startPos)

Main

Sub Main()
        If Not FileExists(SAP_GUI_PATH) Then
                MsgBox "El archivo no existe en la ruta especificada.", vbExclamation, "Archivo no encontrado"
                Exit Sub
        End If

        ExecuteAndWaitForSAP

        ' SapGuiApplication
        Set root = GetObject("SAPGUI")
        Set Application = root.GetScriptingEngine

        ' Open sync connection
        '
        ' Without SSO
        ' Set Connection = Application.OpenConnection(lk0_auth, True)
        ' Set Session = Connection.Children(0)
        ' Session.findById("wnd[0]/usr/txtRSYST-BNAME").text = user
        ' Session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = password
        ' Session.findById("wnd[0]/tbar[0]/btn[0]").press
        '        
        ' SSO
        Set Connection = Application.OpenConnection(lk0_sso, True)
        Set Session = Connection.Children(0)

'         ' Open any tx
'         Session.StartTransaction("cohvpi")
'         session.findById("wnd[0]").sendVKey 0
'         session.findById("wnd[0]/tbar[1]/btn[17]").press
'         session.findById("wnd[1]/usr/txtV-LOW").text = "/ENT_PRO_PRD"
'         session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
'         session.findById("wnd[1]/usr/txtENAME-LOW").setFocus
'         session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 0
'         session.findById("wnd[1]/tbar[0]/btn[8]").press
'         session.findById("wnd[0]/tbar[1]/btn[8]").press
'         session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
'         session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectContextMenuItem "&PC"
'         session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
'         session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
'         session.findById("wnd[1]/tbar[0]/btn[0]").press
'         session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\accgarces\Desktop\MODELACIÓN CARPETA PERSONAL\Informe. Programa Producción. cumplimiento y ocupaciones\Entregas\Entreg" 'Ruta para guardar el archivo
'         session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = Fecha1 & ".txt"
'         session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
'         session.findById("wnd[1]/tbar[0]/btn[0]").press
'         session.findById("wnd[0]/tbar[0]/btn[3]").press
'         session.findById("wnd[0]").resizeWorkingPane 169,23,false
'         session.findById("wnd[0]/tbar[0]/btn[3]").press
End Sub

Sub ExecuteAndWaitForSAP()
        ' Run saplogon bin
        WScript.CreateObject("WScript.Shell").Run Chr(34) & SAP_GUI_PATH & Chr(34), 2

        ' Wait to be initialized
        isSapInitialized = False
        Do While Not isSapInitialized
                isSapInitialized = IsProcessRunning(SAP_BIN)
        Loop
        
        WScript.Sleep 3000
End Sub

Function IsProcessRunning(targetProcess)
        Set WMIService = GetObject("winmgmts:\\.\root\cimv2")
        query = "SELECT * FROM Win32_Process"
        Set items = WMIService.ExecQuery(query)

        For Each item In items
                If item.Name = targetProcess Then
                        IsProcessRunning = True
                        Exit Function
                End If
        Next

        IsProcessRunning = False
End Function

Function FileExists(filePath)
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FileExists(filePath) Then 
                FileExists = True 
        Else
                FileExists = False
        End If
End Function