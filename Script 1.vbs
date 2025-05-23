If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If

Dim rutaScript
rutaScript = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)

'rutaScript = Replace(rutaScript, "\", "/")
rutaScript = rutaScript & "\PDF\"


'VB conexion con Excel
   Dim objEcxcel
   Dim objSheet, intRow
   Dim i
   Dim x
   Set objExcel = GetObject(,"Excel.Application")
   Set objSheet = objExcel.ActiveWorkbook.ActiveSheet


session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm02"
session.findById("wnd[0]").sendVKey 0

For i=2 To objSheet.UsedRange.Rows.Count
   Estado = Trim(CStr(objSheet.Cells(i,4).Value))
   'MsgBox(Estado)

   If Estado = "Pendiente" Then

      Material = Trim(CStr(objSheet.Cells(i,1).Value))
      Centro = Trim(CStr(objSheet.Cells(i,2).Value))
      CarVal = Trim(CStr(objSheet.Cells(i,3).Value))

      'Se debe de tener preseleccionada la vista Contabilidad 1
      session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm02"
      session.findById("wnd[0]").sendVKey 0
      session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = Material
      session.findById("wnd[0]").sendVKey 0
      session.findById("wnd[1]").sendVKey 0
      session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = Centro
      session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").caretPosition = 4
      session.findById("wnd[1]").sendVKey 0

      On Error Resume Next
      session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:/CWM/SAPLMGD1:2800/subSUB1:SAPLCKMMAT:0010/tabsTABS/tabpPPLF/ssubSUBML:SAPLCKMMAT:0100/ctxtMBEW-BKLAS").text = CarVal
      session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:/CWM/SAPLMGD1:2800/subSUB1:SAPLCKMMAT:0010/tabsTABS/tabpPPLF/ssubSUBML:SAPLCKMMAT:0300/ctxtMBEW-BKLAS").text = CarVal
      session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:/CWM/SAPLMGD1:2800/subSUB1:SAPLCKMMAT:0010/tabsTABS/tabpPPLF/ssubSUBML:SAPLCKMMAT:0100/ctxtMBEW-BKLAS").setFocus
      session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:/CWM/SAPLMGD1:2800/subSUB1:SAPLCKMMAT:0010/tabsTABS/tabpPPLF/ssubSUBML:SAPLCKMMAT:0100/ctxtMBEW-BKLAS").caretPosition = 4
      On Error Goto 0
      session.findById("wnd[0]").sendVKey 0

      On Error Resume Next
         mensaje = session.findById("wnd[0]/sbar").text

         'mensaje = session.findById("wnd[0]/usr/subSUBSCREEN:SAPLSBAL_DISPLAY:0101/cntlSAPLSBAL_DISPLAY_CONTAINER/shellcont/shell/shellcont[1]/shell").GetCellValue(0, "T_MSG")

         
         
      On Error Goto 0

      If mensaje = "No se puede modificar categoría valoración, seleccione ""Visualizar error""" Then
         session.findById("wnd[0]/tbar[1]/btn[25]").press
         session.findById("wnd[0]/tbar[0]/btn[86]").press
         session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSPRI:0600/cmbPRIPAR_DYN-PRIMM2").setFocus
         session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSPRI:0600/cmbPRIPAR_DYN-PRIMM2").key = "X"
         session.findById("wnd[1]/tbar[0]/btn[13]").press


         Dim texto, shell
         texto = Material & "-" & Centro
         Set shell = CreateObject("WScript.Shell")
         ' Crear archivo temporal con el texto
         Set fso = CreateObject("Scripting.FileSystemObject")
         Set tempFile = fso.CreateTextFile(fso.GetSpecialFolder(2) & "\tmp_clip.txt", True)
         tempFile.Write texto
         tempFile.Close
         ' Usar clip.exe para copiarlo al portapapeles
         shell.Run "cmd /c type """ & fso.GetSpecialFolder(2) & "\tmp_clip.txt"" | clip", 0, True
         ' Eliminar archivo temporal si deseas
         fso.DeleteFile fso.GetSpecialFolder(2) & "\tmp_clip.txt"

            Set WshShell = CreateObject("WScript.Shell")
            WScript.Sleep 2000 ' Espera a que aparezca la ventana

            ' Escribe la ruta completa del archivo y presiona Enter
            WshShell.SendKeys rutaScript & texto &".pdf"
            WScript.Sleep 500
            WshShell.SendKeys "{ENTER}"

         'MsgBox(Material & "-" & Centro)

         session.findById("wnd[0]").sendVKey 3
         session.findById("wnd[0]").sendVKey 3

         objSheet.Cells(i, 4).Value = "No se puede modificar categoría valoración, seleccione ""Visualizar error""" & ". Se adjunta PDF " & texto &".pdf"
      
      Else
         objExcel.Cells(i,4).Value = mensaje

         session.findById("wnd[0]/tbar[0]/btn[15]").press
         session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
      End If

      estadoFinal = Trim(CStr(objSheet.Cells(i,4).Value))

      If estadoFinal = "" Then
         objSheet.Cells(i, 4).Value = "Se modifica Categoria de Valoracion"
      End If

   End If

Next 'i

MsgBox("Proceso terminado")
