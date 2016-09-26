#include <IE.au3>
#include <Date.au3>

$oIE = _IECreate("http://mitchell.ultipro.com",0,1,0)
Login()
$TempFind = "Mitchell International"
WaitWindowVisible($TempFind)
$oIE1 = _IECreate("https://mitchell.ultipro.com/pages/VIEW/UTASSO.aspx?USParams=PK=ESS!MenuID=1410!PageRerId=1410!ParentRerId=72",1)
$TempFind = "ULTIMATE SOFTWARE"
$found = WaitWindowVisible($TempFind)
ConsoleWrite($found)
$oIE2 = _IEAttach($found,"windowtitle")
_IENavigate($oIE2,"https://mitchell.ultiprotime.com/action/dailytimesheet.action?action=SelectTimesheetAction&clearUIPath=true")
_IELoadWait($oIE2)
ConsoleWrite("\n 1")
$Go = _IEGetObjByName($oIE2,"Go")
ConsoleWrite(" -->"&$Go.value)
_IEAction ($Go, "focus")
_IEAction ($Go, "click")
Sleep(3000)
_IELoadWait($oIE2)
$o_form = _IEFormGetObjByName ($oIE2, "page_form")
$T0 = _IEGetObjByName($o_form,"c_9999_0")
while $T0 = Null
   $T0 = _IEGetObjByName($o_form,"c_9999_0")
   ConsoleWrite("Hit")
WEnd

;$T0.value="9:00"
;$T0 = _IEGetObjByName($oIE2,"ts_c_9999_0")
;$T0.value="AM"
;$T0 = _IEGetObjByName($oIE2,"c_9999_1")
;$T0.value="12:30"
;$T0 = _IEGetObjByName($oIE2,"ts_c_9999_1")
;$T0.value="PM"
;$T0 = _IEGetObjByName($oIE2,"c_9999_2")
;$T0.value="1:30"
;$T0 = _IEGetObjByName($oIE2,"ts_c_9999_2")
;$T0.value="PM"
;$T0 = _IEGetObjByName($oIE2,"c_9999_3")
;$T0.value="6:30"
;$T0 = _IEGetObjByName($oIE2,"ts_c_9999_3")
;$T0.value="PM"
;$Apply = _IEGetObjByName($oIE2,"APPLY_0")
;_IEAction ($Apply, "click")
;$Apply = _IEGetObjByName($oIE2,"APPLY_1")
;_IEAction ($Apply, "click")
;$Apply = _IEGetObjByName($oIE2,"APPLY_2")
;_IEAction ($Apply, "click")
;$Apply = _IEGetObjByName($oIE2,"APPLY_3")
;_IEAction ($Apply, "click")
;$Apply = _IEGetObjByName($oIE2,"APPLY_4")
;_IEAction ($Apply, "click")

_IEImgClick($oIE2,"/images/icon_expand_all.gif")
Local $iWeekday = _DateToDayOfWeekISO(@YEAR, @MON, @MDAY)
$i=0
while $iWeekday <> 1
   $i+=1
   $iWeekday = _DateToDayOfWeekISO(@YEAR, @MON, @MDAY-$i)
WEnd
$mday=@MDAY-$i
ConsoleWrite(@LF&@YEAR&@MON&$mday)
;MsgBox(true,"Click",$oIE)
For $i=0 To 4
   For $j=0 To 2
	  For $k=0 To 1
		 $timeBox = _IEGetObjByName($oIE2,"X"&$i&"0"&$j&"0"&$k)
		 ;$timeBox_minuites = _IEGetObjByName($oIE2,"X"&$i&"0"&$j&"0"&$k&"_minutes")
		 ConsoleWrite(@LF&"X"&$i&"0"&$j&"0"&$k&": "&$timeBox.value)
		 if $j=0 Then
			if $k=0 Then
			   $timeBox.value=@YEAR&@MON&$mday+$i&" 090000"
			   ;$timeBox_minuites.value=0
			Else
			   $timeBox.value=@YEAR&@MON&$mday+$i&" 123000"
			   ;$timeBox_minuites.value=30
			EndIf
		 ElseIf $j=1 Then
			if $k=0 Then
			   $timeBox.value=@YEAR&@MON&$mday+$i&" 123000"
			   ;$timeBox_minuites.value=30
			Else
			   $timeBox.value=@YEAR&@MON&$mday+$i&" 130000"
			   ;$timeBox_minuites.value=0
			EndIf
		 Else
			if $k=0 Then
			   $timeBox.value=@YEAR&@MON&$mday+$i&" 130000"
			   ;$timeBox_minuites.value=0
			Else
			   $timeBox.value=@YEAR&@MON&$mday+$i&" 183000"
			   ;$timeBox_minuites.value=0
			EndIf
		 EndIf
	  Next
   Next
$Feild = _IEGetObjByName($oIE2,"X"&$i&"0003")	;X00003
_IEAction ($Feild, "click")
$Feild.value="WRK"
$Feild = _IEGetObjByName($oIE2,"X"&$i&"0103")	;X00103
_IEAction ($Feild, "click")
$Feild.value="MEAL"
$Feild = _IEGetObjByName($oIE2,"X"&$i&"0203")	;X00203
_IEAction ($Feild, "click")
$Feild.value="WRK"
$Feild = _IEGetObjByName($oIE2,"X"&$i&"0005")	;X00005
_IEAction ($Feild, "click")
$Feild.value="0600720000"
$Feild = _IEGetObjByName($oIE2,"X"&$i&"0205")	;X00205
_IEAction ($Feild, "click")
$Feild.value="0600720000";
$Feild = _IEGetObjByName($oIE2,"X"&$i&"0006")	;X00006
_IEAction ($Feild, "click")
$Feild.value="IMPLEMENTATION"
$Feild = _IEGetObjByName($oIE2,"X"&$i&"0206")	;X00206
_IEAction ($Feild, "click")
$Feild.value="IMPLEMENTATION"
Next
;SAVE
MsgBox(true,"Click",$oIE)
$save = _IEGetObjByName($oIE2,"SubmitTopButton")
_IEAction ($save, "click")
_IELoadWait($oIE2)

;$Del = _IEGetObjByName($oIE2,"X0000TRASH")
;$Del.value="Y"
;$Del = _IEGetObjByName($oIE2,"X1000TRASH")
;$Del.value="Y"
;$Del = _IEGetObjByName($oIE2,"X2000TRASH")
;$Del.value="Y"
;$Del = _IEGetObjByName($oIE2,"X3000TRASH")
;$Del.value="Y"
;$Del = _IEGetObjByName($oIE2,"X4000TRASH")
;$Del.value="Y"

;SAVE
;$save = _IEGetObjByName($oIE2,"SubmitTopButton")
;_IEAction ($save, "click")
;_IELoadWait($oIE2)

$Feild = Null
while $Feild = Null
   $Feild = _IEGetObjByName($oIE2,"X40103")
   ConsoleWrite("Hit")
WEnd

;$Feild = _IEGetObjByName($oIE2,"X00103")	;X00103
;_IEAction ($Feild, "click")
;$Feild.value="MEAL"
;$Feild = _IEGetObjByName($oIE2,"X10103")	;X00103
;_IEAction ($Feild, "click")
;$Feild.value="MEAL"
;$Feild = _IEGetObjByName($oIE2,"X20103")	;X00103
;_IEAction ($Feild, "click")
;$Feild.value="MEAL"
;$Feild = _IEGetObjByName($oIE2,"X30103")	;X00103
;_IEAction ($Feild, "click")
;$Feild.value="MEAL"
;$Feild = _IEGetObjByName($oIE2,"X40103")	;X00103
;_IEAction ($Feild, "click")
;$Feild.value="MEAL"

;$Feild = _IEGetObjByName($oIE2,"X00005")	;X00005
;_IEAction ($Feild, "click")
;$Feild.value="0600720000"
;$Feild = _IEGetObjByName($oIE2,"X00205")	;X00205
;_IEAction ($Feild, "click")
;$Feild.value="0600720000";
;$Feild = _IEGetObjByName($oIE2,"X10005")	;X00005
;_IEAction ($Feild, "click")
;$Feild.value="0600720000"
;$Feild = _IEGetObjByName($oIE2,"X10205")	;X00205
;_IEAction ($Feild, "click")
;$Feild.value="0600720000";
;$Feild = _IEGetObjByName($oIE2,"X20005")	;X00005
;_IEAction ($Feild, "click")
;$Feild.value="0600720000"
;$Feild = _IEGetObjByName($oIE2,"X20205")	;X00205
;_IEAction ($Feild, "click")
;$Feild.value="0600720000";
;$Feild = _IEGetObjByName($oIE2,"X30005")	;X00005
_;IEAction ($Feild, "click")
;$Feild.value="0600720000"
;$Feild = _IEGetObjByName($oIE2,"X30205")	;X00205
;_IEAction ($Feild, "click")
;$Feild.value="0600720000";
;$Feild = _IEGetObjByName($oIE2,"X40005")	;X00005
;_IEAction ($Feild, "click")
;$Feild.value="0600720000"
;$Feild = _IEGetObjByName($oIE2,"X40205")	;X00205
;_IEAction ($Feild, "click")
;$Feild.value="0600720000";
;
;$Feild = _IEGetObjByName($oIE2,"X00006")	;X00006
;_IEAction ($Feild, "click")
;$Feild.value="IMPLEMENTATION"
;$Feild = _IEGetObjByName($oIE2,"X00206")	;X00206
;_IEAction ($Feild, "click")
;$Feild.value="IMPLEMENTATION"
;$Feild = _IEGetObjByName($oIE2,"X10006")	;X00006
;_IEAction ($Feild, "click")
;$Feild.value="IMPLEMENTATION"
;$Feild = _IEGetObjByName($oIE2,"X10206")	;X00206
;_IEAction ($Feild, "click")
;$Feild.value="IMPLEMENTATION"
;$Feild = _IEGetObjByName($oIE2,"X20006")	;X00006
;_IEAction ($Feild, "click")
;$Feild.value="IMPLEMENTATION"
;$Feild = _IEGetObjByName($oIE2,"X20206")	;X00206
;_IEAction ($Feild, "click")
;$Feild.value="IMPLEMENTATION"
;$Feild = _IEGetObjByName($oIE2,"X30006")	;X00006
;_IEAction ($Feild, "click")
;$Feild.value="IMPLEMENTATION"
;$Feild = _IEGetObjByName($oIE2,"X30206")	;X00206
;_IEAction ($Feild, "click")
;$Feild.value="IMPLEMENTATION"
;$Feild = _IEGetObjByName($oIE2,"X40006")	;X00006
;_IEAction ($Feild, "click")
;$Feild.value="IMPLEMENTATION"
;$Feild = _IEGetObjByName($oIE2,"X40206")	;X00206
;_IEAction ($Feild, "click")
;$Feild.value="IMPLEMENTATION"

;SAVE
$save = _IEGetObjByName($oIE2,"SubmitTopButton")
_IEAction ($save, "click")


Func Login()
    ; Retrieve a list of window handles.
    $flag = False
    ; Loop through the array displaying only visable windows with a title.
	While $flag = False
	   Local $aList = WinList("Windows Security")
    For $i = 1 To $aList[0][0]
        If $aList[$i][0] <> "" And BitAND(WinGetState($aList[$i][1]), 2) Then
		   if $aList[$i][0]="Windows Security" Then
			WinActivate($aList[$i][0])
			Send($CmdLine[1])
			Send("{TAB}")
			Send($CmdLine[2])
			Send("{Enter}")
			$flag = True
		 EndIf
	  EndIf
   Next
WEnd
EndFunc

Func WaitWindowVisible($find)
   $flag = False
    ; Loop through the array displaying only visable windows with a title.
	While $flag = False
	   Local $aList = WinList()
	  For $i = 1 To $aList[0][0]
		 If $aList[$i][0] <> "" And BitAND(WinGetState($aList[$i][1]), 2) Then
			if StringInStr($aList[$i][0],$find) <> 0 Then
				  $flag = True
				  ExitLoop 2
			   EndIf
			EndIf
		 Next
	  WEnd
	  Return $aList[$i][0]
EndFunc
