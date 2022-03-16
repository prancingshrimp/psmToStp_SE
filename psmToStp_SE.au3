#include <GUIConstantsEx.au3>
#include <MsgBoxConstants.au3>
#include <File.au3>
#include <Array.au3>

Opt('MustDeclareVars', 1)

Func _ErrFunc($oError)
	MsgBox(BitOR(0,16), "Error", "There is a problem with Solid Edge!")
	Exit
EndFunc

Local $gui, $guiwidth, $guiheight, $buttonwidth, $buttonheight
Local $button_convert, $button_quit
Local $path, $txt_path, $label
Local $signal

$buttonwidth = 120
$buttonheight = $buttonwidth/4
$guiwidth = 2*$buttonwidth+20
$guiheight = 2*$buttonheight+25

$gui = GUICreate("psmToStp",$guiwidth,$guiheight,2,2)
$label = GUICtrlCreateLabel("Please select folder with psm files.",5,1*$buttonheight+15,400)
GUICtrlCreateLabel("v0.1",5,2*$buttonheight+5,400)
$button_convert = GUICtrlCreateButton("Select folder",5,5,$buttonwidth,$buttonheight)
$button_quit = GUICtrlCreateButton("Quit",$buttonwidth+15,5,$buttonwidth,$buttonheight)

GUISetState(@SW_SHOW)

While 1
   $signal = GUIGetMsg()
   Select
	  Case $signal = $GUI_EVENT_CLOSE
		 ExitLoop

	  Case $signal = $button_quit
		 ExitLoop

	   Case $signal = $button_convert

		 $txt_path = "T:\16_Technik"
		 $path = FileSelectFolder("Choose the destination folder", $txt_path)
		 If @error Then
			MsgBox(BitOR(0,16), "Error", "No folder has been selected!")
			ContinueLoop
		 EndIf

		 convert($path, $label)
	EndSelect
WEnd
GUIDelete($gui)


Func convert($path, $label)

	Local $size, $fileArray
	$fileArray = _FileListToArrayRec($path, "*.psm", $FLTAR_FILES, $FLTAR_NORECUR, $FLTAR_SORT)
	If @error Then
		If @extended = 9 Then
			MsgBox(BitOR(0,16), "Error", "No psm files found!")
			Return
		EndIf
	EndIf
	$size = $fileArray[0]


	For $i = 1 To $size
		$fileArray[$i] = StringReplace($fileArray[$i],".psm","")
	Next

	GUICtrlSetData($label, "Starting Solid Edge. Please wait... ")

	Local $oErrorHandler = ObjEvent("AutoIt.Error", "_ErrFunc")
	Local $oEdge = ObjCreate("SolidEdge.Application")
	If @error Then
		MsgBox(BitOR(0,16), "_Error_", "There is a problem with Solid Edge!")
		Exit
	EndIf

	$oEdge.Visible = True
	$oEdge.DisplayAlerts = False

	For $i = 1 To $size
		GUICtrlSetData($label, "Working on no. " & $i & " out of " & $size & " psm files.")
		Local $objDoc = $oEdge.Documents.Open($path & "\" & $fileArray[$i] & ".psm")
;~ 		$oEdge.DoIdle()
		$objDoc.SaveAs($path & "\" & $fileArray[$i] & ".stp")
		$objDoc.Close(False)
	Next

	$oEdge.Quit()

	MsgBox(64, "", "Done!")
	GUICtrlSetData($label, "Please select folder with psm files.")
EndFunc