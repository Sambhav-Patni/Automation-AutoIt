#include <Array.au3>
#include <Excel.au3>
#include <MsgBoxConstants.au3>

Local $oWorkbook = _ExcelBookOpen($CmdLine[1])
If @error Then
    MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead Example", "Error opening workbook '" & @ScriptDir & "\Extras\_Excel1.xls'." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
    _Excel_Close($oExcel)
    Exit
EndIf

Run("C:\WINDOWS\system32\cmd.exe")
WinWaitActive("C:\WINDOWS\system32\cmd.exe")
send('cd C:\Documents and Settings\sp102532\Desktop\StdPkg_ReProcessing_Tool' & "{ENTER}")
send('c:' & "{ENTER}")
For $i=2 To 102
   send('java -Dlog4j.configuration=file:"./settings/log4j.xml" -DSettingsFilePath="./settings" -jar stdpkg-reprocess-tool-handler-1.0.0.0.jar' & "{ENTER}")
   send('I9' & "{ENTER}")
   ; *****************************************************************************
   ; Read data from a single cell on the active sheet of the specified workbook
   ; *****************************************************************************
   Local $sResult = _ExcelReadCell($oWorkbook, "K"&$i)
   Local $sResult1 = _ExcelReadCell($oWorkbook, "D"&$i)
   If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead Example 1", "Error reading from workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
   ;MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead Example 1", "Data successfully read." & @CRLF & "Value of cell A1: " & $sResult)
   send($sResult & "{ENTER}")
   send($sResult1 & "{ENTER}")
   Sleep(30000)
Next