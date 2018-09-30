#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=time-management-clock-small.ico
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_SaveSource=y
#AutoIt3Wrapper_Res_Language=1033
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#cs ----------------------------------------------------------------------------

    AutoIt Version: 3.3.14.5
    Author:         PZ

    Script Function: Daily Time Code Tracker

   TODO:
   - 
   
   


#ce ----------------------------------------------------------------------------

#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <ListViewConstants.au3>
#include <Excel.au3>
#include <GuiListView.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <GuiEdit.au3>
#include <Constants.au3>
#include <Misc.au3>
#include <Date.au3>

$Form1 = GUICreate("Daily Time Charge Code", 505, 429, 192, 124)

#Region DATE
$tCur = GUICtrlCreateDate("", 150, 377, 200, 20, $DTS_SHORTDATEFORMAT)
#EndRegion DATE

$Label2 = GUICtrlCreateLabel("*Time", 180, 20, 200, 20)
$Input2 = GUICtrlCreateInput("", 180, 40, 201, 22) ;time Input
$Label1 = GUICtrlCreateLabel("Details", 180, 64, 200, 20)
$Input1 = GUICtrlCreateEdit("", 180, 84, 200, 80) ;details input
$Input3 = GUICtrlCreateInput("", 60, 125, 60, 22)  ;other time code input

#Region TIMECODES
$Timecode1 = GUICtrlCreateRadio( "ECA Certs", 10, 20)
$Timecode2 = GUICtrlCreateRadio( "Loaner Laptops", 10, 40)
$Timecode3 = GUICtrlCreateRadio( "Printers", 10, 60)
$Timecode4 = GUICtrlCreateRadio( "Conference Rooms", 10, 80)
$Timecode5 = GUICtrlCreateRadio( "Ticket/Call", 10, 100)
$Timecode6 = GUICtrlCreateRadio( "Other", 10, 122)
#EndRegion

$ListView1 = GUICtrlCreateListView("Task|Time Code|Time|Details", 8, 184, 485, 185)
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 0, 150)
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 1, 73)
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 2, 73)
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 3, 185)
$Button2 = GUICtrlCreateButton("Export to file", 400, 376, 89, 40, $BS_NOTIFY) 	;Save button
GUICtrlSetCursor(-1, 0)
$Button4 = GUICtrlCreateButton("Add Entry", 400, 39, 89, 125, $BS_NOTIFY) 		;Add button
GUICtrlSetCursor(-1, 0)

GUICtrlSetState($Input3, $GUI_DISABLE)
GUICtrlSetState($Timecode1, $GUI_CHECKED)
$dummy = GUICtrlCreateDummy ()
GUISetState(@SW_SHOW)

$datetoformat = _NowDate()

While 1

   $nMsg = GUIGetMsg()

   Switch $nMsg
	  Case $GUI_EVENT_CLOSE
		Exit
	  Case $Form1
		ToolTip ("")

	Case $tCur
		$datetoformat = GUICtrlRead($tCur)
		FormatDate()

	Case $Button2
		$fname = FormatDate() & "_Time_Charge_Codes"
		$exfile = FileSaveDialog ( "Name of exported document.", "", "Excel file (*.xlsx)", 18, $fname, $Form1 )
		If $exfile = "" Then

		Else
			$excel = StringTrimRight ( $exfile, 4 ) & "xlsx"
			$num = _GUICtrlListView_GetItemCount ( $ListView1 )
			Local $array[$num + 1][4]
			$array[0][0] = "Task"
			$array[0][1] = "Charge Code"
			$array[0][2] = "Time"
			$array[0][3] = "Details"
			For $y = 0 To $num - 1 Step 1
			   $listarray = _GUICtrlListView_GetItemTextArray ( $ListView1, $y )
			   $array[$y + 1][0] = $listarray[1]
			   $array[$y + 1][1] = $listarray[2]
			   $array[$y + 1][2] = $listarray[3]
			   $array[$y + 1][3] = $listarray[4]
			Next
			$oExcel = _Excel_Open ()
			$oWorksheet = _Excel_BookNew ( $oExcel )
			_Excel_RangeWrite ( $oWorksheet, Default, $array )
			_Excel_BookSaveAs ( $oWorksheet, $excel )
			;_Excel_Export ( $oExcel, $oWorksheet, $exfile, Default, Default, Default, Default, Default, True )
			_Excel_Close ( $oExcel )
		EndIf

	Case $Input2
		If _IsPressed ( "0D" ) Then
			AddItem()
		EndIf

	Case $Button4
		AddItem()

	Case $Timecode6
		If GUICtrlRead($Timecode6) = $GUI_CHECKED Then
			GUICtrlSetState($Input3, $GUI_ENABLE)
		Else
			GUICtrlSetState($Input3, $GUI_DISABLE)
		EndIf

   	EndSwitch
WEnd

Func AddItem()
	If GUICtrlRead($Input2) = "" Then
			$sToolTipAnswer = ToolTip("You need to enter a time value.", Default, Default, "Enter time")
			Sleep(2000)
			ToolTip ("")
		Else
			ToolTip ("")
			Local $checkedTimeCode
			Local $inputtime = GUICtrlRead($Input2, "")
			Local $details = GUICtrlRead($Input1)
			$checkedTimeCode = GetTimeCode()
			GUICtrlCreateListViewItem($checkedTimecode[1] & "|" & $checkedTimecode[0] & "|" & $inputtime & "|" & $details, $ListView1)
			GUICtrlSetData ( $Input1, "" )
			GUICtrlSetData ( $Input2, "" )
	EndIf
EndFunc

Func GetTimecode()
   	Local $tcode
   	Local $tname
   	Local $timecode[2]
   	If GUICtrlRead($Timecode1) = $GUI_CHECKED Then
	  	$tcode = 20983
	  	$tname = "ECA Certs"
   	ElseIf GUICtrlRead($Timecode2) = $GUI_CHECKED Then
		$tcode = 20984
		$tname = "Loaner Laptops"
	ElseIf GUICtrlRead($Timecode3) = $GUI_CHECKED Then
		$tcode = 20986
		$tname = "Printers"
	ElseIf GUICtrlRead($Timecode4) = $GUI_CHECKED Then
		$tcode = 20987
		$tname = "Conference Rooms"
	ElseIf GUICtrlRead($Timecode5) = $GUI_CHECKED Then
		$tcode = 21019
		$tname = "Ticket/Call"
	ElseIf GUICtrlRead($Timecode6) = $GUI_CHECKED Then
		$tcode = GUICtrlRead($Input3)
		$tname = "Other"
	EndIf
	$timecode[0] = $tcode
	$timecode[1] = $tname
	Return $timecode
EndFunc

;format example YYYYMMDD
Func FormatDate()
	$split = StringSplit($datetoformat,"/n")
	$YYYY = StringRight($split[3],4)
	$DD = StringLeft($split[2],2)
	$MM = StringLeft($split[1],2)
	$dateformated = $YYYY & $MM & $DD
	Return $dateformated
EndFunc