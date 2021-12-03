#NoEnv ; For security
#SingleInstance force
;==================================================================
; Current date is ";d0"   DatePicker GUI is ;dp
; Dates in the future are ";dn" where n = number of days.  (Max 9)
; Dates in the past are ";ddn" where n = number of day. (Max 9)
:?*:;dd9::
:?*:;dd8::
:?*:;dd7::
:?*:;dd6::
:?*:;dd5::
:?*:;dd4::
:?*:;dd3::
:?*:;dd2::
:?*:;dd1::
:?*:;d0::
:?*:;d1::
:?*:;d2::
:?*:;d3::
:?*:;d4::
:?*:;d5::
:?*:;d6::
:?*:;d7::
:?*:;d8::
:?*:;d9::

   StringReplace,nOffset,A_ThisHotkey,:?*:;d 
   StringReplace,nOffset,nOffset,d,- ; This first part condenced with help forum members :)

Offset += %nOffset%, days ; Puts offset into date format.
SetTitleMatchMode, 2
IfWinActive, GoalView 
{
   FormatTime, MyDate, %OffSet%, M/d/yyyy
}
else 
{
	IniRead, dateFormat, DateFormat.txt, Date tool, Format, M/d/yyyy
    FormatTime, MyDate, %OffSet%, %dateFormat%
 }
SendInput {U+1F4C5}%MyDate%   ; This types out the date.

FormatTime, DOWtoday ,,WDay ;====== This is all for the tooltip/popup.====
DOWsum := $ DOWtoday + nOffset ; "DOW" is "day or week," not "Dow Jones."
if (DOWsum > 7) 
	MySuffix = `, next week
else if (DOWsum > 14)
	MySuffix = `, week after next
else if (DOWsum < -1)
	MySuffix = `, last week
else if (DOWsum < -8)
	MySuffix = `, week before last
else
	MySuffix = 
myToolTipX := A_CaretX + 10 ; For position of tooltip.
myToolTipY := A_CaretY + 25

FormatTime, DayOfWeek, %OffSet%, dddd
If (DayOfWeek = "Saturday") || (DayOfWeek = "Sunday") {
   MsgBox, 48, , WARNING:`n`nThat falls on a weekend.`n`n     %DayOfWeek%
}
else
   ToolTip, %DayOfWeek%%MySuffix%, %myToolTipX%, %myToolTipY%
SetTimer, RemoveToolTip, 2000
OffSet =		; Reset to nothing.
nOffset =
return
RemoveToolTip:
SetTimer, RemoveToolTip, Off
ToolTip 
return  ;========== End of Tooltip section ================
   
:?*:;dp:: ;=========== Popup calendar ===============
	Gui, dp:Add, MonthCal, vOffSet
	Gui, dp:Add, Button, Default, Submit 
	Gui, dp:Show ,,Date Picker
Return
dpButtonSubmit:
	Gui, dp:Submit 
SetTitleMatchMode, 2
IfWinActive, GoalView 
{
   FormatTime, MyDate, %OffSet%, M/d/yyyy
}
else 
{
	IniRead, dateFormat, DateFormat.txt, Date tool, Format, M/d/yyyy
    FormatTime, MyDate, %OffSet%, %dateFormat%
 }
SendInput {U+1F4C5}%MyDate%   ; This types out the date.
;Esc::
;dpGuiClose:
	Gui, dp:Destroy 
Return