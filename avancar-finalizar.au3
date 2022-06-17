#include <File.au3>
#include <Array.au3>
#include <String.au3>
#include <SAP2.au3>
#include <Date.au3>
#include <Excel.au3>

Local $sFile_xlsx = @DesktopDir & '/Avancar-Finalizar Notas.xlsx' ; Planilha das notas a avançar/finalizar NEOENERGIA PERNAMBUCO

lerExcel()
_log()
abrirCCS()
avancarMedidas()

Func lerExcel()

   ; Fecha Excel caso tenha algum aberto
	  RunWait(@ComSpec & " /c taskkill /IM excel.exe /F", "", @SW_HIDE)

   ; Cria um objeto Excel e abre o workbook da planilha AVANÇAR-FINALIZAR NOTAS
   Local $oExcel = _Excel_Open()
   Local $oWorkbook = _Excel_BookOpen($oExcel, $sFile_xlsx)
	  If @error Then
		 MsgBox($MB_SYSTEMMODAL, "Erro", "Erro ao abrir Excel")
		 _Excel_Close($oExcel)
		 Exit
	  EndIf

   ; Armazena as informações de NOTA, ALINEA e DATA

   Local $sPrimeira_nota = _Excel_RangeRead($oWorkbook, Default, "A2")
   Local $sSegunda_nota = _Excel_RangeRead($oWorkbook, Default, "A3")

   Local $sPrimeira_alinea = _Excel_RangeRead($oWorkbook, Default, "B2")
   Local $sSegunda_alinea = _Excel_RangeRead($oWorkbook, Default, "B3")

   Local $sPrimeira_data = _Excel_RangeRead($oWorkbook, Default, "C2")
   Local $sSegunda_data = _Excel_RangeRead($oWorkbook, Default, "C3")

	  ; NÃO HÁ NOTAS NA PLANILHA
	  If $sPrimeira_nota = '' and $sPrimeira_alinea = '' and $sPrimeira_data = '' Then
		 ; Mensagem de aviso
			MsgBox(0, 'Aviso', 'A planilha est� vazia ou n�o foram inseridas informa��es na primeira linha da tabela.')

		 ; Encerra o script
			Exit

	  ; H� APENAS UMA NOTA NA PLANILHA
	  ElseIf $sSegunda_nota = '' and $sSegunda_alinea = '' and $sSegunda_data = '' Then
		 Local $sNota = _Excel_RangeRead($oWorkbook, Default, "A2")
		 Local $sAlinea = _Excel_RangeRead($oWorkbook, Default, "B2")
		 Local $sData = _Excel_RangeRead($oWorkbook, Default, "C2")

		 ; Confirma se o preenchimento foi feito corretamente

			; Informa��o n�o preenchida
			   If $sNota = '' or $sAlinea = '' or $sData = '' Then
				  ; Mensagem de aviso
					 MsgBox(0, 'Aviso', 'Aten��o! Verifique se h� alguma informa��o n�o preenchida e rode o programa novamente.')

				  ; Encerra o script
					 Exit
			   EndIf

			; Alinea incorreta
			   If $sAlinea <> 'Deferido' and $sAlinea <> 'Deferido Parcial' and $sAlinea <> 'Solicita��o 90 dias' and $sAlinea <>  'Solicita��o 5 anos' and $sAlinea <> 'Suspens�o' and $sAlinea <> 'Cortado' and $sAlinea <> 'Indeferido' and $sAlinea <> 'A' and $sAlinea <> 'B' and $sAlinea <> 'C' and $sAlinea <> 'D' and $sAlinea <> 'E' and $sAlinea <> 'F' and $sAlinea <> 'G' and $sAlinea <> 'H' and $sAlinea <> 'I' and $sAlinea <> 'J' and $sAlinea <> 'K' and $sAlinea <> 'L' and $sAlinea <> 'M' and $sAlinea <> 'N' and $sAlinea <> 'O' and $sAlinea <> 'P' and $sAlinea <> 'Q' and $sAlinea <> 'R' and $sAlinea <> 'S' and $sAlinea <> 'T' and $sAlinea <> 'U' and $sAlinea <> 'V' and $sAlinea <> 'W' Then
				  ; Mensagem de aviso
					 MsgBox(0, 'Aviso', 'Aten��o! A al�nea inserida � inv�lida. Corrija e rode o programa novamente.')

				  ; Encerra o script
					 Exit
			   EndIf

		 Global $aNota[1]
		 Global $aAlinea[1]
		 Global $aData[1]

		 $aNota[0] = $sNota
		 $aAlinea[0] = $sAlinea
		 $aData[0] = $sData

	  ; H� DUAS NOTAS OU MAIS NA PLANILHA
	  Else
		 Global $aNota = _Excel_RangeRead($oWorkbook, Default, $oWorkbook.Sheets("Plan1").Range($oExcel.Application.Sheets("Plan1").Range("A2"), $oExcel.Application.Sheets("Plan1").Range("A2").End(-4121)))
		 If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Erro", "Erro - Coluna NOTA" & @CRLF & "@error = " & @error & ", @extended = " & @extended)

		 Global $aAlinea = _Excel_RangeRead($oWorkbook, Default, $oWorkbook.Sheets("Plan1").Range($oExcel.Application.Sheets("Plan1").Range("B2"), $oExcel.Application.Sheets("Plan1").Range("B2").End(-4121)))
		 If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Erro", "Erro - Coluna ALINEA" & @CRLF & "@error = " & @error & ", @extended = " & @extended)

		 Global $aData = _Excel_RangeRead($oWorkbook, Default, $oWorkbook.Sheets("Plan1").Range($oExcel.Application.Sheets("Plan1").Range("C2"), $oExcel.Application.Sheets("Plan1").Range("C2").End(-4121)))
		 If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Erro", "Erro - Coluna DATA" & @CRLF & "@error = " & @error & ", @extended = " & @extended)

		 ; Confirma se o preenhimento foi feito corretamente

			; Informa��o n�o preenchida
			If UBound($aNota) <> UBound($aAlinea) or UBound($aNota) <> UBound($aData) or UBound($aAlinea) <> UBound($aData) Then
			   ; Mensagem de aviso
				  MsgBox(0, 'Aviso', 'Aten��o! Verifique se h� alguma informa��o n�o preenchida e rode o programa novamente.')

			   ; Encerra o Script
				  Exit
			EndIf

			; Alinea incorreta
			For $i = 0 to UBound($aAlinea) - 1
			   If $aAlinea[$i] <> 'Deferido' and $aAlinea[$i] <> 'Deferido Parcial' and $aAlinea[$i] <> 'Solicita��o 90 dias' and $aAlinea[$i] <>  'Solicita��o 5 anos' and $aAlinea[$i] <> 'Suspens�o' and $aAlinea[$i] <> 'Cortado' and $aAlinea[$i] <> 'Indeferido' and $aAlinea[$i] <> 'A' and $aAlinea[$i] <> 'B' and $aAlinea[$i] <> 'C' and $aAlinea[$i] <> 'D' and $aAlinea[$i] <> 'E' and $aAlinea[$i] <> 'F' and $aAlinea[$i] <> 'G' and $aAlinea[$i] <> 'H' and $aAlinea[$i] <> 'I' and $aAlinea[$i] <> 'J' and $aAlinea[$i] <> 'K' and $aAlinea[$i] <> 'L' and $aAlinea[$i] <> 'M' and $aAlinea[$i] <> 'N' and $aAlinea[$i] <> 'O' and $aAlinea[$i] <> 'P' and $aAlinea[$i] <> 'Q' and $aAlinea[$i] <> 'R' and $aAlinea[$i] <> 'S' and $aAlinea[$i] <> 'T' and $aAlinea[$i] <> 'U' and $aAlinea[$i] <> 'V' and $aAlinea[$i] <> 'W' Then
				  ; Mensagem de aviso
					 MsgBox(0, 'Aviso', 'Aten��o! A al�nea inserida na linha ' & $i + 2 & ' � inv�lida. Corrija e rode o programa novamente.')

				  ; Encerra o script
					 Exit
			   EndIf
			Next
	  EndIf

   ; Fecha Excel
	  RunWait(@ComSpec & " /c taskkill /IM excel.exe /F", "", @SW_HIDE)

EndFunc

Func _log()

   ; Array do LOG
	  Global $aLog[UBound($aNota, $UBOUND_ROWS) + 2][4]

   ; Cabe�alho
	  $aLog[0][0] = '   Status  '
	  $aLog[0][1] = '    Nota    '
	  $aLog[0][2] = '        Al�nea       '
	  $aLog[0][3] = '   Observa��o '

	  $aLog[1][0] = '-----------'
	  $aLog[1][1] = '------------'
	  $aLog[1][2] = '---------------------'
	  $aLog[1][3] = '-----------------------'

   ; Preenche as colunas de Status, Nota e Al�nea
	  For $i = 2 to UBound($aLog, $UBOUND_ROWS) - 1
		 $aLog[$i][0] = '           '									; Status
		 $aLog[$i][1] = ' ' & $aNota[$i-2] & ' '						; Nota
		 If StringLen($aAlinea[$i-2]) = 1 Then							; Al�nea
			$aLog[$i][2] = '          ' & $aAlinea[$i-2] & '          '
		 ElseIf StringLen($aAlinea[$i-2]) = 7 Then
			$aLog[$i][2] = '      ' & $aAlinea[$i-2] & '        '
		 ElseIf StringLen($aAlinea[$i-2]) = 8 Then
			$aLog[$i][2] = '      ' & $aAlinea[$i-2] & '       '
		 ElseIf StringLen($aAlinea[$i-2]) = 9 Then
			$aLog[$i][2] = '      ' & $aAlinea[$i-2] & '      '
		 ElseIf StringLen($aAlinea[$i-2]) = 10 Then
			$aLog[$i][2] = '     ' & $aAlinea[$i-2] & '      '
		 ElseIf StringLen($aAlinea[$i-2]) = 11 Then
			$aLog[$i][2] = '     ' & $aAlinea[$i-2] & '     '
		 Elseif StringLen($aAlinea[$i-2]) = 16 Then
			$aLog[$i][2] = '  ' & $aAlinea[$i-2] & '   '
		 Elseif StringLen($aAlinea[$i-2]) = 18 Then
			$aLog[$i][2] = ' ' & $aAlinea[$i-2] & '  '
		 Elseif StringLen($aAlinea[$i-2]) = 19 Then
			$aLog[$i][2] = ' ' & $aAlinea[$i-2] & ' '
		 EndIf
	  Next

EndFunc

Func abrirCCS()

   ; Fecha qualquer SAP que esteja aberto
	  RunWait(@ComSpec & " /c taskkill /IM saplogon.exe /F", "", @SW_HIDE)

   ; Faz login no SAP CCS
	  Sleep(500)
	  $CMD = ' cd /d C:\Program Files (x86)\SAP\FrontEnd\SAPgui && ' & _
		 'sapgui brnep474 01'
	  RunWait('"' & @ComSpec & '" /c ' & $CMD , '', @SW_HIDE)	; executa o SAP no servidor da Celpe

	  While 1
		 If WinExists("SAP") Then
			Sleep(3000)
			WinActivate("SAP")
			RunWait(@Comspec & ' /c cscript.exe ' & @ScriptDir &"\fazerLogonCCS.vbs", "", @SW_HIDE)	; executa o VBscript que faz logon
			ExitLoop
		 EndIf
		 Sleep(1000)
	  WEnd

EndFunc

Func avancarMedidas()

   ; Vincula � transa��o IW52
	  _SAPSessAttach("[CLASS:SAP_FRONTEND_SESSION]", "IW52")
	  WinWait('Modificar nota servi�o: 1� tela')
	  WinActivate("Modificar nota servi�o: 1� tela")

   ; Entra em cada nota e avan�a as medidas de acordo com a Alinea
	  For $i = 0 to UBound($aNota) - 1

		 Global $bConcluido = False
		 Global $bMedida_correta = False
		 Global $bMedida_avancada = False

		 While 1
		 ; Continua apenas se a p�gina da IW52 estiver carregada
			WinWait('Modificar nota servi�o: 1� tela')

		 ; Insere o n�mero da nota
			$sap_session.findById("wnd[0]/usr/ctxtRIWO00-QMNUM").text = $aNota[$i]
			Sleep(200)
			Send("{ENTER}")

		 ; Confirma se a nota est� na medida correta

			; Seleciona a aba 'Medidas'
			   $sap_session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11").select

			; Identifica se a nota j� est� finalizada
			   If WinExists('Exibir nota de servi�o: Recl serv.Danos El�t') Then
				  If $aAlinea[$i] = 'A' or $aAlinea[$i] = 'B' or $aAlinea[$i] = 'C' or $aAlinea[$i] = 'D' or $aAlinea[$i] = 'E' or $aAlinea[$i] = 'F' or $aAlinea[$i] = 'G' or $aAlinea[$i] = 'H' or $aAlinea[$i] = 'I' or $aAlinea[$i] = 'J' or $aAlinea[$i] = 'K' or $aAlinea[$i] = 'L' or $aAlinea[$i] = 'M' or $aAlinea[$i] = 'N' or $aAlinea[$i] = 'O' or $aAlinea[$i] = 'P' or $aAlinea[$i] = 'Q' or $aAlinea[$i] = 'R' or $aAlinea[$i] = 'S' or $aAlinea[$i] = 'T' or $aAlinea[$i] = 'U' or $aAlinea[$i] = 'V' or $aAlinea[$i] = 'W' or $aAlinea[$i] = 'Cortado'  or $aAlinea[$i] = 'Suspens�o' or $aAlinea[$i] = 'Indeferido' Then
					 $bMedida_avancada = True
				  Endif
				  _SAPVKeysSend("F3")
				  ExitLoop
			   EndIf

			; Identifica a primeira linha vazia
			   For $j = 0 To 20
				  If $sap_session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/ctxtVIQMSM-MNGRP[1," & $j & "]").text = "" Then
					 $iLinha_vazia = $j
					 ExitLoop
				  EndIf
			   Next

			; Identifica se a nota foi uma improcedente finalizada ou se est� em RECLIMPR
			   If $aAlinea[$i] = 'C' or $aAlinea[$i] = 'D' or $aAlinea[$i] = 'E' or $aAlinea[$i] = 'F' or $aAlinea[$i] = 'G' or $aAlinea[$i] = 'H' or $aAlinea[$i] = 'I' or $aAlinea[$i] = 'J' or $aAlinea[$i] = 'K' or $aAlinea[$i] = 'L' or $aAlinea[$i] = 'M' or $aAlinea[$i] = 'N' or $aAlinea[$i] = 'O' or $aAlinea[$i] = 'P' or $aAlinea[$i] = 'Q' or $aAlinea[$i] = 'R' or $aAlinea[$i] = 'S' or $aAlinea[$i] = 'T' or $aAlinea[$i] = 'U' or $aAlinea[$i] = 'V' or $aAlinea[$i] = 'W' or $aAlinea[$i] = 'Cortado' or $aAlinea[$i] = 'Suspens�o' or $aAlinea[$i] = 'Indeferido' Then
				  If $sap_session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/ctxtVIQMSM-MNGRP[1," & $iLinha_vazia - 1 & "]").text = "RECLIMPR" Then
					 $bMedida_correta = True
				  EndIf
			   Endif

			; Identifica se a nota foi uma alinea A finalizada ou se est� em COMUNCLT
			   If $aAlinea[$i] = 'A' Then
				  If $sap_session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/ctxtVIQMSM-MNGRP[1," & $iLinha_vazia - 1 & "]").text = "COMUNCLT" Then
					 $bMedida_correta = True
				  EndIf
			   EndIf

			; Identifica se a nota foi uma alinea B finalizada ou se est� em NOTINTER
			   If $aAlinea[$i] = 'B' Then
				  If $sap_session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/ctxtVIQMSM-MNGRP[1," & $iLinha_vazia - 1 & "]").text = "NOTINTER" Then
					 $bMedida_correta = True
				  EndIf
			   EndIf

			; Identifica se a nota est� em RECLPROC ou se foi uma RECLPROC avan�ada para PAGAUTOR
			   If $aAlinea[$i] = 'Deferido' or $aAlinea[$i] = 'Deferido parcial' Then
				  If $sap_session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/ctxtVIQMSM-MNGRP[1," & $iLinha_vazia - 1 & "]").text = "RECLPROC" Then
					 $bMedida_correta = True
				  EndIf

				  If $sap_session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/ctxtVIQMSM-MNGRP[1," & $iLinha_vazia - 1 & "]").text = "PAGAUTOR" Then
					 $bMedida_avancada = True
				  EndIf
			   EndIf

			; Identifica se a nota est� em SOLICDOC ou se foi uma SOLCIDOC avan�ada para NOTINTER
			   If $aAlinea[$i] = 'Solicita��o 90 dias' or $aAlinea[$i] = 'Solicita��o 5 anos' Then
				  If $sap_session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/ctxtVIQMSM-MNGRP[1," & $iLinha_vazia - 1  & "]").text = "SOLICDOC" Then
					 $bMedida_correta = True
				  EndIf

				  If $sap_session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/ctxtVIQMSM-MNGRP[1," & $iLinha_vazia - 1  & "]").text = "NOTINTER" Then
					 $bMedida_avancada = True
				  EndIf
			   EndIf

		 ; Se a medida n�o estiver na correta, pula para a pr�xima nota
			If $bMedida_correta = False Then
			   _SAPVKeysSend("F3")
			   ExitLoop
			EndIf

		 ; Seleciona a aba 'Informa��es Adicionais'
			$sap_session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB21").select

		 If $aAlinea[$i] = 'Deferido' or $aAlinea[$i] = 'Deferido parcial' or $aAlinea[$i] = 'A' or $aAlinea[$i] = 'B' or $aAlinea[$i] = 'C' or $aAlinea[$i] = 'D' or $aAlinea[$i] = 'E' or $aAlinea[$i] = 'F' or $aAlinea[$i] = 'G' or $aAlinea[$i] = 'H' or $aAlinea[$i] = 'I' or $aAlinea[$i] = 'J' or $aAlinea[$i] = 'K' or $aAlinea[$i] = 'L' or $aAlinea[$i] = 'M' or $aAlinea[$i] = 'N' or $aAlinea[$i] = 'O' or $aAlinea[$i] = 'P' or $aAlinea[$i] = 'Q' or $aAlinea[$i] = 'R' or $aAlinea[$i] = 'S' or $aAlinea[$i] = 'T' or $aAlinea[$i] = 'U' or $aAlinea[$i] = 'V' or $aAlinea[$i] = 'W' or $aAlinea[$i] = 'Cortado' or $aAlinea[$i] = 'Suspens�o' or $aAlinea[$i] = 'Indeferido' Then
			; Insere a 'Data de Envio da Carta ao Cliente'
			   $sap_session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB21/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:1200/ctxtE_ZCTNSCE-DATA_ENVI").text = $aData[$i]

			; Insere a 'Hora de Envio da Carta ao Cliente'
			   $sap_session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB21/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:1200/ctxtE_ZCTNSCE-HORA_ENVI").text = ""

			; Insere a 'Data de Recebimento da Carta do Cliente'
			   $sap_session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB21/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:1200/ctxtE_ZCTNSCE-DATA_REC").text = $aData[$i]

			; Insere a 'Hora de Recebimento da Carta do Cliente'
			   $sap_session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB21/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:1200/ctxtE_ZCTNSCE-HORA_REC").text = ""
		 Endif

		 If $aAlinea[$i] = 'B' Then
			; Insere a 'Data de Entrega da Documenta��o'
			   $sap_session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB21/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:1200/ctxtE_ZCTNSCE-DATA_DOC").text = $aData[$i]

			; Insere a 'Hora de Entrega da Documenta��o'
			   $sap_session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB21/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:1200/ctxtE_ZCTNSCE-HORA_DOC").text = ""
		 EndIf

		 If $aAlinea[$i] = 'Solicita��o 90 dias' or $aAlinea[$i] = 'Solicita��o 5 anos' Then
			; Insere a 'Data de Solicita��o da Documenta��o'
			   $sap_session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB21/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:1200/ctxtE_ZCTNSCE-DATA_SOL").text = $aData[$i]

			; Insere a 'Hora de Solicita��o da Documenta��o'
			   $sap_session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB21/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:1200/ctxtE_ZCTNSCE-HORA_SOL").text = ""
		 EndIf

		 ; Seleciona aba 'Medidas'
			$sap_session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11").select

		 ; Identifica��o da linha da medida

			If $aAlinea[$i] = 'Deferido' or $aAlinea[$i] = 'Deferido parcial' Then
			   ; Identifica qual a linha da medida RECLPROC
				  For $j = 0 To 20
					 If $sap_session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/ctxtVIQMSM-MNGRP[1," & $j & "]").text = "RECLPROC" Then
						$iLinha = $j
						ExitLoop
					 EndIf
				  Next
			   EndIf

			If $aAlinea[$i] = 'C' or $aAlinea[$i] = 'D' or $aAlinea[$i] = 'E' or $aAlinea[$i] = 'F' or $aAlinea[$i] = 'G' or $aAlinea[$i] = 'H' or $aAlinea[$i] = 'I' or $aAlinea[$i] = 'J' or $aAlinea[$i] = 'K' or $aAlinea[$i] = 'L' or $aAlinea[$i] = 'M' or $aAlinea[$i] = 'N' or $aAlinea[$i] = 'O' or $aAlinea[$i] = 'P' or $aAlinea[$i] = 'Q' or $aAlinea[$i] = 'R' or $aAlinea[$i] = 'S' or $aAlinea[$i] = 'T' or $aAlinea[$i] = 'U' or $aAlinea[$i] = 'V' or $aAlinea[$i] = 'W' or $aAlinea[$i] = 'Cortado' or $aAlinea[$i] = 'Suspens�o' or $aAlinea[$i] = 'Indeferido' Then
			   ; Identifica qual a linha da medida RECLIMPR
				  For $j = 0 To 20
					 If $sap_session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/ctxtVIQMSM-MNGRP[1," & $j & "]").text = "RECLIMPR" Then
						$iLinha = $j
						ExitLoop
					 EndIf
				  Next
			Endif

			If $aAlinea[$i] = 'A' Then
			   ; Identifica qual a linha da medida COMUNCLT
				  For $j = 0 To 20
					 If $sap_session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/ctxtVIQMSM-MNGRP[1," & $j & "]").text = "COMUNCLT" Then
						$iLinha = $j
						ExitLoop
					 EndIf
				  Next
			EndIf

			If $aAlinea[$i] = 'B' Then
			   ; Identifica qual a linha da medida AGUARDOC
				  For $j = 0 To 20
					 If $sap_session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/ctxtVIQMSM-MNGRP[1," & $j & "]").text = "AGUARDOC" Then
						$iLinha = $j
						ExitLoop
					 EndIf
				  Next

			   ; Identifica qual a linha da medida NOTINTER
				  For $j = 0 To 20
					 If $sap_session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/ctxtVIQMSM-MNGRP[1," & $j & "]").text = "NOTINTER" Then
						$iLinha2 = $j
						ExitLoop
					 EndIf
				  Next
			EndIf

			If $aAlinea[$i] = 'Solicita��o 90 dias' or $aAlinea[$i] = 'Solicita��o 5 anos' Then
			   ; Identifica qual a linha da medida SOLICDOC
				  For $j = 0 To 20
					 If $sap_session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/ctxtVIQMSM-MNGRP[1," & $j & "]").text = "SOLICDOC" Then
						$iLinha = $j
						ExitLoop
					 EndIf
				  Next
			EndIf

		 ; Avan�o/Finaliza��o de notas

		 If $aAlinea[$i] = 'Deferido' or $aAlinea[$i] = 'Deferido parcial' or $aAlinea[$i] = 'C' or $aAlinea[$i] = 'D' or $aAlinea[$i] = 'E' or $aAlinea[$i] = 'F' or $aAlinea[$i] = 'G' or $aAlinea[$i] = 'H' or $aAlinea[$i] = 'I' or $aAlinea[$i] = 'J' or $aAlinea[$i] = 'K' or $aAlinea[$i] = 'L' or $aAlinea[$i] = 'M' or $aAlinea[$i] = 'N' or $aAlinea[$i] = 'O' or $aAlinea[$i] = 'P' or $aAlinea[$i] = 'Q' or $aAlinea[$i] = 'R' or $aAlinea[$i] = 'S' or $aAlinea[$i] = 'T' or $aAlinea[$i] = 'U' or $aAlinea[$i] = 'V' or $aAlinea[$i] = 'W' or $aAlinea[$i] = 'Cortado' or $aAlinea[$i] = 'Suspens�o' or $aAlinea[$i] = 'Indeferido' Then
			; Avan�a de RECLPROC ou RECLIMPR para COMUNCLT
			   $sap_session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER").getAbsoluteRow($iLinha).selected = true
			   $sap_session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/btnFC_ERLEDIGT").press
			   $sap_session.findById("wnd[0]/shellcont/shell").selectItem("0010","Column01")
			   $sap_session.findById("wnd[0]/shellcont/shell").ensureVisibleHorizontalItem("0010","Column01")
			   $sap_session.findById("wnd[0]/shellcont/shell").clickLink("0010","Column01")

			; Avan�a de COMUNCLT para PAGAUTOR ou FINLNOT
			   $sap_session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER").getAbsoluteRow($iLinha + 1).selected = true
			   $sap_session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/btnFC_ERLEDIGT").press
			   $sap_session.findById("wnd[0]/shellcont/shell").selectItem("0010","Column01")
			   $sap_session.findById("wnd[0]/shellcont/shell").ensureVisibleHorizontalItem("0010","Column01")
			   $sap_session.findById("wnd[0]/shellcont/shell").clickLink("0010","Column01")
		 Endif

		 If $aAlinea[$i] = 'A' Then
			; Avan�a de COMUNCLT para FINLNOT
			   $sap_session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER").getAbsoluteRow($iLinha).selected = true
			   $sap_session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/btnFC_ERLEDIGT").press
			   $sap_session.findById("wnd[0]/shellcont/shell").selectItem("0010","Column01")
			   $sap_session.findById("wnd[0]/shellcont/shell").ensureVisibleHorizontalItem("0010","Column01")
			   $sap_session.findById("wnd[0]/shellcont/shell").clickLink("0010","Column01")
		 EndIf

		 If $aAlinea[$i] = 'B' Then
			; Avan�a de AGUARDOC e NOTINTER para PARATEND
			   $sap_session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER").getAbsoluteRow($iLinha).selected = true
			   $sap_session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER").getAbsoluteRow($iLinha2).selected = true
			   $sap_session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/btnFC_ERLEDIGT").press
			   $sap_session.findById("wnd[0]/shellcont/shell").selectItem("0010","Column01")
			   $sap_session.findById("wnd[0]/shellcont/shell").ensureVisibleHorizontalItem("0010","Column01")
			   $sap_session.findById("wnd[0]/shellcont/shell").clickLink("0010","Column01")

			; Altera PARATEND para RECLIMPR e avan�a para COMUNCLT
			   $sap_session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER").getAbsoluteRow($iLinha2 + 1).selected = true
			   $sap_session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/ctxtVIQMSM-MNGRP[1," & $iLinha2 + 1 & "]").text = "RECLIMPR"
			   $sap_session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/ctxtVIQMSM-MNCOD[2," & $iLinha2 + 1 & "]").text = "0002"
			   $sap_session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/btnFC_ERLEDIGT").press
			   $sap_session.findById("wnd[0]/shellcont/shell").selectItem("0010","Column01")
			   $sap_session.findById("wnd[0]/shellcont/shell").ensureVisibleHorizontalItem("0010","Column01")
			   $sap_session.findById("wnd[0]/shellcont/shell").clickLink("0010","Column01")

			; Avan�a de COMUNCLT para FINLNOT
			   $sap_session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER").getAbsoluteRow($iLinha2 + 2).selected = true
			   $sap_session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/txtVIQMSM-QSMNUM[0," & $iLinha2 + 2 & "]").setFocus
			   $sap_session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/txtVIQMSM-QSMNUM[0," & $iLinha2 + 2 & "]").caretPosition = 0
			   $sap_session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/btnFC_ERLEDIGT").press
			   $sap_session.findById("wnd[0]/shellcont/shell").selectItem("0010","Column01")
			   $sap_session.findById("wnd[0]/shellcont/shell").ensureVisibleHorizontalItem("0010","Column01")
			   $sap_session.findById("wnd[0]/shellcont/shell").clickLink("0010","Column01")
		 EndIf

		 If $aAlinea[$i] = 'Solicita��o 90 dias' or $aAlinea[$i] = 'Solicita��o 5 anos' Then
			; Avan�a de SOLICDOC para NOTINTER
			   $sap_session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER").getAbsoluteRow($iLinha).selected = true
			   $sap_session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/btnFC_ERLEDIGT").press
			   $sap_session.findById("wnd[0]/shellcont/shell").selectItem("0010","Column01")
			   $sap_session.findById("wnd[0]/shellcont/shell").ensureVisibleHorizontalItem("0010","Column01")
			   $sap_session.findById("wnd[0]/shellcont/shell").clickLink("0010","Column01")
		 Endif

		 ; Grava nota
			$sap_session.findById("wnd[0]/tbar[0]/btn[11]").press

		 $bConcluido = True

		 ExitLoop
		 WEnd

	  If $bConcluido = True Then
		 $aLog[$i+2][0] = ' Concluido '
	  Else
		 $aLog[$i+2][0] = '   Pulada  '
	  EndIf

	  If $bMedida_correta = False and $bMedida_avancada = True Then
		 $aLog[$i+2][3] = ' Medida j� avan�ada'
	  EndIf

	  If $bMedida_correta = False and $bMedida_avancada = False Then
		 $aLog[$i+2][3] = ' Medida n�o corresponde a Al�nea inserida'
	  EndIf

	  _FileWriteFromArray(@ScriptDir & '\LOG.txt', $aLog)

	  Next

   MsgBox(0, 'Aviso', 'Conclu�do! Confira o LOG aberto no bloco de notas.')

   ; Fecha SAP
	  ;RunWait(@ComSpec & " /c taskkill /IM saplogon.exe /F", "", @SW_HIDE)
EndFunc



















