#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=static\icon.ico
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Res_Description=Send User Message Popup : utilitaire pour envoyer des messages flash (popup) a un utilisateur
#AutoIt3Wrapper_Res_Fileversion=1.0.0.0
#AutoIt3Wrapper_Res_ProductName=sump
#AutoIt3Wrapper_Res_ProductVersion=1.0.0
#AutoIt3Wrapper_Res_CompanyName=CNAMTS/CPAM_ARTOIS/BEI
#AutoIt3Wrapper_Res_LegalCopyright=bei.cpam-artois@assurance-maladie.fr
#AutoIt3Wrapper_Res_Language=1036
#AutoIt3Wrapper_Res_Compatibility=Win7
#AutoIt3Wrapper_Res_Field=AutoIt Version|%AutoItVer%
#AutoIt3Wrapper_AU3Check_Stop_OnWarning=Y
#AutoIt3Wrapper_AU3Check_Parameters=-q -d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7
#AutoIt3Wrapper_UseX64=Y
#Au3Stripper_Parameters=/MO /RSLN
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

; #INDEX# =======================================================================================================================
; Title .........: pmfinfos
; AutoIt Version : 3.3.14.5
; Language ......: French
; Description ...: Script .au3
; Author(s) .....: yann.daniel@assurance-maladie.fr
; ===============================================================================================================================

; #ENVIRONMENT# =================================================================================================================
; AutoIt3Wrapper
; Includes YD
#include "C:\Users\DANIEL-03598\Autoit_dev\Include\YDGVars.au3"
#include "C:\Users\DANIEL-03598\Autoit_dev\Include\YDLogger.au3"
#include "C:\Users\DANIEL-03598\Autoit_dev\Include\YDTool.au3"
; Includes Constants
#include <StaticConstants.au3>
#Include <WindowsConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <MsgBoxConstants.au3>
; Includes
#include <String.au3>
; Options
;~ AutoItSetOption("MustDeclareVars", 1)
;~ AutoItSetOption("WinTitleMatchMode", 2)
;~ AutoItSetOption("WinDetectHiddenText", 1)
;~ AutoItSetOption("MouseCoordMode", 0)
;~ AutoItSetOption("TrayMenuMode", 3)
;~ OnAutoItExitRegister("_YDTool_ExitApp")
; ===============================================================================================================================

; #VARIABLES# ===================================================================================================================
_YDGVars_Set("sAppName", _YDTool_GetAppWrapperRes("ProductName"))
_YDGVars_Set("sAppDesc", _YDTool_GetAppWrapperRes("Description"))
_YDGVars_Set("sAppVersion", _YDTool_GetAppWrapperRes("ProductVersion"))
_YDGVars_Set("sAppContact", _YDTool_GetAppWrapperRes("LegalCopyright"))
_YDGVars_Set("sAppVersionV", "v" & _YDGVars_Get("sAppVersion"))
_YDGVars_Set("sAppTitle", _YDGVars_Get("sAppName") & " - " & _YDGVars_Get("sAppVersionV"))
_YDGVars_Set("sAppDirDataPath", @ScriptDir & "\data")
_YDGVars_Set("sAppDirStaticPath", @ScriptDir & "\static")
_YDGVars_Set("sAppDirLogsPath", @ScriptDir & "\logs")
_YDGVars_Set("sAppDirVendorPath", @ScriptDir & "\vendor")
_YDGVars_Set("sAppIconPath", @ScriptDir & "\static\icon.ico")
_YDGVars_Set("sAppConfFile", @ScriptDir & "\conf.ini")
_YDGVars_Set("iAppNbDaysToKeepLogFiles", 15)

_YDLogger_Init()
_YDLogger_LogAllGVars()
; ===============================================================================================================================

; #MAIN SCRIPT# =================================================================================================================
If Not _YDTool_IsSingleton() Then Exit
;------------------------------
; On supprime les anciens fichiers de log
_YDTool_DeleteOldFiles(_YDGVars_Get("sAppDirLogsPath"), _YDGVars_Get("iAppNbDaysToKeepLogFiles"))
;------------------------------
; On cree le repertoire data s il n existe pas
_YDTool_CreateFolderIfNotExist(_YDGVars_Get("sAppDirDataPath"))
;------------------------------
_Main()
; #MAIN SCRIPT# =================================================================================================================

; #MAIN LOOP# ====================================================================================================================
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Description ...: Traitement principal
; Syntax ........: _Main()
; Parameters ....:
; Return values .:
; Author ........: yann.daniel@assurance-maladie.fr
; Last Modified .: 08/06/2021
; Notes .........:
;================================================================================================================================
Func _Main()
	Local $sFuncName = "_Main"
	Local $sUserName, $sInputPMF, $sMessage, $sModelName, $sModelFileContent, $sMessageFilePath
	Local $hFilehandle
	; ---------------
	; GUI
	GUICreate(_YDGVars_Get("sAppTitle"), 600, 200, (@DesktopWidth - 600) / 2, (@DesktopHeight - 220) / 2)
	; ---------------
	; Nom du PMF ou IP
	Local $sClipBoard = ClipGet()
	If StringInStr($sClipBoard, "S116201") <> 0 Or StringInStr($sClipBoard, "P116201") <> 0 Then
		_YDLogger_Log("Recuperation du clipboard : " & $sClipBoard, $sFuncName)
	Else
		$sClipBoard = ""
	EndIf
	GUICtrlCreateLabel("Nom du PMF : ", 10, 20, 100, 20)
	Local $idInputPMF = GUICtrlCreateInput($sClipBoard, 85, 18, 110, 20)
	; ---------------
	; Modeles
	GUICtrlCreateLabel("Modèle : ", 235, 20, 100, 20)
	Local $idComboModel = GUICtrlCreateCombo("", 280, 18, 280, 20, BitOR($CBS_DROPDOWNLIST,$CBS_AUTOHSCROLL))
	GUICtrlCreateIcon("shell32.dll", 7, 565, 20, 16, 16)
    GUICtrlSetState( -1, $GUI_DISABLE)
    Local $idButtonModelSave = GUICtrlCreateButton("", 565, 20, 16, 16, $WS_CLIPSIBLINGS)
	Local $aModels = _FileListToArray(_YDGVars_Get("sAppDirDataPath"), "*.txt", $FLTA_FILES)
	; On sort si tableau vide ou erreur
    If  @error <> 0  Or Not IsArray($aModels) Or UBound($aModels) = 0 Then
        _YDTool_SetMsgBoxError("Impossible de recuperer la liste des modeles !", $sFuncName)
		Exit
    EndIf
    _YDLogger_Var("$UBound($aModels)", UBound($aModels), $sFuncName, 2)
    ; Si tout est OK, on boucle sur le tableau
	Local $sModelList = ""
    For $i = 1 To UBound($aModels) - 1
		_YDLogger_Var("$aModels[" & $i & "]",  $aModels[$i], $sFuncName, 2)
		If $i = 1 Then
        	$sModelList = $aModels[$i]
		Else
			$sModelList = $sModelList & "|" & $aModels[$i]
		EndIf
    Next
	_YDLogger_Var("$sModelList", $sModelList, $sFuncName, 2)
    GUICtrlSetData($idComboModel, $sModelList, $aModels[1])
	; ---------------
	; Message
	GUICtrlCreateLabel("Message : ", 10, 50, 100, 20)
	$sModelFileContent = FileRead(_YDGVars_Get("sAppDirDataPath") & "\" & $aModels[1])
	Local $idInputEditMessage = GUICtrlCreateEdit($sModelFileContent, 10, 70, 570, 80)
	; ---------------
	; Bouton d envoi
	Local $idButtonSend = GUICtrlCreateButton("Envoyer le message", 220, 155, 130, 40)
	GUISetState()
	; ---------------
	; Main Loop
	While 1
		Local $iMsg = GuiGetMsg()
		Select
			Case $iMsg = $GUI_EVENT_CLOSE
				ExitLoop
			Case $iMsg = $idComboModel
				$sModelName = GUICtrlRead($idComboModel)
				_YDLogger_Var("$sModelName", $sModelName, $sFuncName, 2)
				$sModelFileContent = FileRead(_YDGVars_Get("sAppDirDataPath") & "\" & $sModelName)
				GUICtrlSetData($idInputEditMessage, $sModelFileContent)
			Case $iMsg = $idButtonModelSave
				$sModelName = GUICtrlRead($idComboModel)
				$sMessage = GUICtrlRead($idInputEditMessage)
				; On ecrase le modele
				$hFilehandle = FileOpen(_YDGVars_Get("sAppDirDataPath") & "\" & $sModelName, $FO_OVERWRITE)
				FileWrite($hFilehandle, $sMessage)
				FileClose($hFilehandle)
				If Not @error Then
					MsgBox(BitOR($MB_SYSTEMMODAL,$MB_ICONINFORMATION), "Sauvegarde du modele", "Le modele <" & $sModelName & "> a ete sauvegarde")
					_YDLogger_Log("Sauvegarde du modele " & $sModelName, $sFuncName)
				Else
					_YDTool_SetMsgBoxError("Erreur lors de la sauvegarde du modele " & $sModelName, $sFuncName)
				EndIf
			Case $iMsg = $idButtonSend
				$sInputPMF = GUICtrlRead($idInputPMF)
				_YDLogger_Var("$sInputPMF", $sInputPMF, $sFuncName, 2)
				$sMessage = GUICtrlRead($idInputEditMessage)
				_YDLogger_Var("$sMessage", $sMessage, $sFuncName, 2)
				; On verifie si ping OK
				If Not _YDTool_IsPing($sInputPMF) Then
					_YDTool_SetMsgBoxError("Pas de reponse au ping pour " & $sInputPMF & " !", $sFuncName)
				Else
					; On verifie si user connecte
					$sUserName = _YDTool_GetHostLoggedUserName($sInputPMF)
					If $sUserName == "" Then
						_YDTool_SetMsgBoxError("Aucun utilisateur connecte pour " & $sInputPMF & " !", $sFuncName)
					Else
						Local $idMsgBoxConfirm = MsgBox($MB_YESNO, "Confirmation", "Etes-vous sur de vouloir envoyer ce message a " & $sUserName & " ?")
						If $idMsgBoxConfirm = $IDYES Then
							; On ecrit le message dans un fichier temporaire
							$sMessageFilePath = @ScriptDir & "\msg.tmp"
							_YDLogger_Log("Ecriture du message dans le fichier " & $sMessageFilePath, $sFuncName, 2)
							_YDTool_DeleteFileIfExist($sMessageFilePath)
							$hFilehandle = FileOpen($sMessageFilePath, $FO_OVERWRITE)
							FileWrite($hFilehandle, $sMessage)
							FileClose($hFilehandle)
							_YDLogger_Log("Envoi du message a " & $sUserName & " sur " & $sInputPMF, $sFuncName)
							Local $sMsgCommand = "msg " & $sUserName & " /server:" & $sInputPMF & " /V /W < " & $sMessageFilePath
							_YDLogger_Var("$sMsgCommand", $sMsgCommand, $sFuncName)
							Local $iReturn = RunWait(@ComSpec & " /c " & $sMsgCommand)
							_YDLogger_Var("$iReturn", $iReturn, $sFuncName, 2)
							If $iReturn = 0 Then
								MsgBox(BitOR($MB_SYSTEMMODAL,$MB_ICONINFORMATION), "Envoi du message", "Le message a bien ete acquitte par " & $sUserName)
							Else
								_YDTool_SetMsgBoxError("Le message n a pas ete envoye ou non acquitte par a " & $sUserName & " !", $sFuncName)
							EndIf
							_YDTool_DeleteFileIfExist($sMessageFilePath)
						EndIf
					EndIf
				EndIf
		EndSelect
		;------------------------------
		Sleep(10)
	WEnd
	Exit
EndFunc
