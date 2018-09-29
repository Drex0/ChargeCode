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
   - File Export date as default name
   - Remove task input or add an array column for it
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

$Label1 = GUICtrlCreateLabel("Project/ticket working on", 200, 10, 200, 29)
$Input1 = GUICtrlCreateInput("", 200, 33, 201, 22) ;task input
$Label2 = GUICtrlCreateLabel("*Time", 200, 63, 200, 29)
$Input2 = GUICtrlCreateInput("", 200, 86, 201, 22) ;time Input

$Input3 = GUICtrlCreateInput("", 60, 115, 60, 22)  ;other time code input

;$Label3 = GUICtrlCreateLabel("Todays Entries", 10, 162, 250, 29)
;GUICtrlSetFont(-1, 15, 400, 0, "" )
#Region TIMECODES
$Timecode1 = GUICtrlCreateRadio( "ECA Certs", 10, 10)
$Timecode2 = GUICtrlCreateRadio( "Loaner Laptops", 10, 30)
$Timecode3 = GUICtrlCreateRadio( "Printers", 10, 50)
$Timecode4 = GUICtrlCreateRadio( "Conference Rooms", 10, 70)
$Timecode5 = GUICtrlCreateRadio( "Ticket/Call", 10, 90)
$Timecode6 = GUICtrlCreateRadio( "Other", 10, 110)
#EndRegion

GUICtrlSetState($Input3, $GUI_DISABLE)
GUICtrlSetState($Timecode1, $GUI_CHECKED)
$ListView1 = GUICtrlCreateListView("Task|Time|Time Code", 8, 194, 485, 175)
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 0, 150)
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 1, 73)
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 2, 110)
$Button2 = GUICtrlCreateButton("Export to file", 400, 376, 89, 41, $BS_NOTIFY) 	;Save button
GUICtrlSetCursor(-1, 0)
$Button4 = GUICtrlCreateButton("Add Entry", 400, 142, 89, 41, $BS_NOTIFY) 		;Add button
GUICtrlSetCursor(-1, 0)
$dummy = GUICtrlCreateDummy ()
GUISetState(@SW_SHOW)

#Region DATE
GUICtrlCreateDate("", 5, 162, 200, 20)
GUICtrlSetTip(-1, '#Region DATE')
GUICtrlCreateLabel("(Date control expands into a calendar)", 10, 305, 200, 20)
GUICtrlSetTip(-1, '#Region DATE - Label')
#EndRegion DATE

$tCur = _NowDate()
;GUICtrlSetData($Label3, $tCur)

While 1

   $nMsg = GUIGetMsg()

   Switch $nMsg
	  Case $GUI_EVENT_CLOSE
		 Exit
	  Case $Form1
		 ToolTip ("")

	  Case $Button2
		 $exfile = FileSaveDialog ( "Name of exported document.", "", "PDF file (*.pdf)", 18, "DailyTimeCodes_", $Form1 )
		 If $exfile = "" Then

		 Else
			$excel = StringTrimRight ( $exfile, 3 ) & "xlsx"
			$num = _GUICtrlListView_GetItemCount ( $ListView1 )
			Local $array[$num + 1][3]
			$array[0][0] = "Task"
			$array[0][1] = "Time"
			$array[0][2] = "Time Code"
			For $y = 0 To $num - 1 Step 1
			   $listarray = _GUICtrlListView_GetItemTextArray ( $ListView1, $y )
			   $array[$y + 1][0] = $listarray[1]
			   $array[$y + 1][1] = $listarray[2]
			   $array[$y + 1][2] = $listarray[3]
			Next
			$oExcel = _Excel_Open ()
			$oWorksheet = _Excel_BookNew ( $oExcel )
			_Excel_RangeWrite ( $oWorksheet, Default, $array )
			_Excel_BookSaveAs ( $oWorksheet, $excel )
			_Excel_Export ( $oExcel, $oWorksheet, $exfile, Default, Default, Default, Default, Default, True )
			_Excel_Close ( $oExcel )
		 EndIf

	  Case $Button4
		 If GUICtrlRead($Input2) = "" Then
			$sToolTipAnswer = ToolTip("You need to enter a time value.", Default, Default, "Enter time")
		 Else
			ToolTip ("")
			Local $checkedTimeCode
			Local $inputtime = GUICtrlRead($Input2, "")
			$checkedTimeCode = GetTimeCode()
			If GUICtrlRead($Input1, "") Then
			   GUICtrlCreateListViewItem(GUICtrlRead($Input1) & "|" & $inputtime & "|" & $checkedTimecode[0], $ListView1)
			Else
			   GUICtrlCreateListViewItem($checkedTimecode[1] & "|" & $inputtime & "|" & $checkedTimecode[0], $ListView1)
			EndIf
			   GUICtrlSetData ( $Input1, "" )
			   GUICtrlSetData ( $Input2, "" )
		 EndIf

	  Case $Timecode6
		 If GUICtrlRead($Timecode6) = $GUI_CHECKED Then
			GUICtrlSetState($Input3, $GUI_ENABLE)
		 Else
			GUICtrlSetState($Input3, $GUI_DISABLE)
		 EndIf

   EndSwitch
WEnd

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
	  $tname = "Ticket\Call"
   ElseIf GUICtrlRead($Timecode6) = $GUI_CHECKED Then
	  $tcode = 00000
	  $tname = "Other"
   EndIf
   $timecode[0] = $tcode
   $timecode[1] = $tname
   Return $timecode
EndFunc