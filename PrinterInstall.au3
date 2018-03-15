#comments-start
@Auth: Gabriel Rodriguez
Program: Printer Install Automate
Year: 2017

#comments-end

#include <Array.au3>
#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiButton.au3>
#include <WindowsConstants.au3>
#include <GUIToolTip.au3>
#include <FontConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiMenu.au3>
#include <WinAPI.au3>
#RequireAdmin

#pragma compile(AutoItExecuteAllowed, True)

; Global variables for drivers and scripts as well as printer names, IP addresses and printer driver information from the .inf files
Global $drivers = "C:\DRIVERS"
Global $scripts = "C:\SCRIPTS"
Global $nscInf = "CNLB0UA64.inf"
Global $driver_7270 = "C:\LaptopBuild\C7270"
Global $driver_3325 = "C:\LaptopBuild\IRC3325"
Global $loc[] = ["Sherwood_Main","Sherwood_House","Vancouver","Vancouver_North","Moses_Lake","Kent_Warehouse2","Kent_Warehouse","Kent_Main"]
Global $ip[] = ['10.100.148.254','10.100.148.251','10.100.150.254','10.100.150.29','10.73.8.10','10.73.19.195','10.73.19.194','10.73.4.25']
Global $printer[] = ["Canon iR-ADV C7260/7270 UFR II","Canon iR-ADV C5045/5051 UFR II","Canon iR-ADV C7260/7270 UFR II","Canon iR-ADV 4245/4251 UFR II","Canon iR-ADV C5030/5035 UFR II","Canon iR-ADV C7260/7270 UFR II","Canon iR-ADV C5250/5255 UFR II","Canon iR-ADV C7055/7065 UFR II"]
Global $loc2[] = ["Tacoma_LE","Tacoma_BSW","Bremerton","Molalla","Sherwood_Shop","Tumwater"]
Global $ip2[] = ['10.73.15.194','10.73.15.195','10.73.11.194','10.73.7.194','10.100.148.250','10.73.13.194']
Global $printer2 = "Canon iR-ADV C3325/3330 UFR II"
Global $x
Global $y
Global $i
Global $k
Global $j
Global $install = 0
Global $end
Global $end2 = 6
Global $next = False
Global $prog

; Global variables for drivers and scripts as well as printer names, IP addresses and printer driver information from the .inf files GSU
Global $drivers = "C:\DRIVERS"
Global $scripts = "C:\SCRIPTS"
Global $gsuInf = "eSf6u.inf"
Global $driver_3505 = "C:\LaptopBuild\a3505"
Global $loc3[] = ["Fremont", "Sacamento", "Sacramento Trailor", "One Fiber Trailer","Selma"]
Global $ip3[] = ['10.54.116.42','10.54.11.194', '10.54.11.196', '10.54.121.194', '10.54.118.5']
Global $printer3 = "TOSHIBA Universal Printer 2"
Global $bArray[2] = []
Global $uninstall

;open GUI
LapTopBuild()

; GUI to display options to user
Func LapTopBuild()
   Opt("GUIOnEventMode", 1)
   Opt("GUIResizeMode", 1)

   Local $hGUI = GUICreate("Laptop Build", 300, 200)

   Local $label = GUICtrlCreateLabel("Select an option from below", 30, 12,280)
   GUICtrlSetFont($label, 11, $FW_Bold,$GUI_FONTUNDER)

   GUISetOnEvent($GUI_EVENT_CLOSE, "SpecialEvents")
   GUISetOnEvent($GUI_EVENT_MINIMIZE, "")
   GUISetOnEvent($GUI_EVENT_RESTORE, "")

   Local $nscP = GUICtrlCreateButton("NSC Printers", 80, 50, 200)
   $bArray[0] = GUICtrlGetHandle($nscP)
   _GUICtrlButton_SetStyle($nscP, $BS_AUTORADIOBUTTON)
   GUICtrlSetFont($nscP, 9, $FW_Bold)

   Local $gsuP = GUICtrlCreateButton("GSU Printers", 80, 90, 200)
   $bArray[1] = GUICtrlGetHandle($gsuP)
   _GUICtrlButton_SetStyle($gsuP, $BS_AUTORADIOBUTTON)
   GUICtrlSetFont($gsuP, 9, $FW_Bold)

   Local $install = GUICtrlCreateButton("Install", 180 , 140, 100)
   GUICtrlSetOnEvent(-1, "CheckChoice")
   GUICtrlSetFont($install, 9, $FW_Bold)

   Local $cancel = GUICtrlCreateButton("Cancel", 20 , 140, 100)
   GUICtrlSetOnEvent(-1, "Cancel")
   GUICtrlSetFont($cancel, 9, $FW_Bold)

   Local $buttonToolTip = _GUIToolTip_Create(0)
   _GUIToolTip_AddTool($buttonToolTip, 0, "NSC printer install", $bArray[0])
   _GUIToolTip_AddTool($buttonToolTip, 0, "GSU printer install", $bArray[1])

   GUISetState(@SW_SHOW)

    ; Loop until the user exits.

   While 1
      sleep(100)
   WEnd
EndFunc

;Iterate over array to find selected radio
Func CheckChoice()
   FOR $j = 0 to 1 step 1
	  Local $iState = _GUICtrlButton_GetCheck($bArray[$j])
	  Switch $iState
	  Case $BST_CHECKED
		 $y = $j
		 $j = 5
	  EndSwitch
   Next
   StartProcess($y)
EndFunc

;receives a number from CheckChoice() and runs selected process
Func StartProcess($start)
   Switch $start
   case 0
	  NPrinters()
   case 1
	  GPrinters()
   EndSwitch
EndFunc

;Closes program and GUI when canceled
Func Cancel()
   GUIDelete()
   Exit
EndFunc

;closes program and gui when the user uses the X to close window
Func SpecialEvents()
   Select
	  Case @GUI_CtrlId = $GUI_EVENT_CLOSE
         GUIDelete()
         Exit
    EndSelect
EndFunc

;NSC printer install only
Func NPrinters()
   GUIDelete()
   $end = 8
   ProgressOn("Installing", "NSC Printers", "0%", 2, 16)
   Init()
   ClearPnt(0)
   sleep(2000)
   AddPnt(0)
   AddPnt(1)
   sleep(3000)
   ProgressSet(100, "Done", "Complete")
   Delete()
   ProgressOff()
   RemoveApp()
   Exit
EndFunc

;GSU printer install only
Func GPrinters()
   GUIDelete()
   $end = 5
   ProgressOn("Installing", "GSU Printers", "0%", 2, 16)
   Init()
   ClearPnt(1)
   sleep(2000)
   AddPnt(2)
   sleep(3000)
   ProgressSet(100, "Done", "Complete")
   Delete()
   ProgressOff()
   RemoveApp()
   Exit
EndFunc

;initialize events by creating directories to hold scripts and drivers
func Init()
   RunWait(@ComSpec & " /c " & "md C:\DRIVERS", "",@SW_HIDE); make directory for printer driver files;2.5
   Progress(10)
EndFunc

; add printer drivers with a parameter for the driver
func AddDriver($driver)
   RunWait(@ComSpec & " /c " & "ROBOCOPY " & $driver & " C:\DRIVERS /e", "",@SW_HIDE); populate driver directory
EndFunc

;Clear any printers listed in the arrays
func ClearPnt($company)
   $k = 0
   $j = 0
   if $company == 0 Then
      while $k < $end
	     RunWait(@ComSpec & " /c " & "cscript C:\Windows\System32\Printing_Admin_Scripts\en-US\Prnmngr.vbs -d -p " & '"' & $loc[$k] & '"', "",@SW_HIDE);1(8)
	     $k += 1
		 Progress(1)
      WEnd
      while $j < $end2
	     RunWait(@ComSpec & " /c " & "cscript C:\Windows\System32\Printing_Admin_Scripts\en-US\Prnmngr.vbs -d -p " & '"' & $loc2[$j] & '"', "",@SW_HIDE);1(6)
	     $j += 1
		 Progress(1)
      WEnd
   EndIf
   If $company == 1 Then
	  while $k < $end
	     RunWait(@ComSpec & " /c " & "cscript C:\Windows\System32\Printing_Admin_Scripts\en-US\Prnmngr.vbs -d -p " & '"' & $loc3[$k] & '"', "",@SW_HIDE);10
	     $k += 1
		 Progress(10)
      WEnd
   EndIf
EndFunc

;Add printers using arrays and drivers
Func AddPnt($print)
   $x = 0
   $i = 0

   if $print == 0 Then
	  AddDriver($driver_7270)
      While $x < $end
         RunScript($ip[$x], $printer[$x], $loc[$x], $nscInf);5.71
	     $x += 1
            Progress(5.42857142857142)
      WEnd
	  delete()
   EndIf
   if $print == 1 Then
	  AddDriver($driver_3325)
      While $i < $end2
		 RunScript($ip2[$i], $printer2, $loc2[$i], $nscInf);5.71
	     $i += 1
		 Progress(5.42857142857142)
      WEnd
   EndIf
   if $print == 2 Then
	  AddDriver($driver_3505)
      While $i < $end
		 RunScript($ip3[$i], $printer3, $loc3[$i], $gsuInf);5
	     $i += 1
		 Progress(8)
      WEnd
   EndIf
EndFunc

;Built in windows scripts to add printers
Func RunScript($ip, $printer, $loc, $printDriver)
   RunWait(@ComSpec & " /c " & "Cscript C:\Windows\System32\Printing_Admin_Scripts\en-US\Prnport.vbs -a -r IP_" & $ip & " -h " & $ip & " -o raw -n 9100", "",@SW_HIDE)
   RunWait(@ComSpec & " /c " & "Cscript C:\Windows\System32\Printing_Admin_Scripts\en-US\Prndrvr.vbs -a -m " & '"' & $printer & '"' & " -i " & "C:\DRIVERS\" & $printdriver & " -h " & $drivers, "",@SW_HIDE)
   RunWait(@ComSpec & " /c " & "Cscript C:\Windows\System32\Printing_Admin_Scripts\en-US\Prnmngr.vbs -a -p " & '"' & $loc & '"' & " -m " & '"' & $printer & '"' & " -r IP_" & $ip, "",@SW_HIDE)
EndFunc

;Delete directory for drivers
func Delete()
   RunWait(@ComSpec & " /c " & "rd C:\DRIVERS /s /q", "",@SW_HIDE); clean up folders
EndFunc

;Display progress to user
Func Progress($pr)
   $prog += $pr
   ProgressSet($prog, "Installing " & Round($prog, 1) & "%")
EndFunc

;self deleting
Func RemoveApp()
   FileDelete("C:\Users\Public\Desktop\PrinterInstall.lnk")
   FileDelete("C:\LaptopBuild\a3505")
   DirRemove("C:\LaptopBuild\a3505", 1)
   FileDelete("C:\LaptopBuild\C7270")
   DirRemove("C:\LaptopBuild\C7270", 1)
   FileDelete("C:\LaptopBuild\IRC3325")
   DirRemove("C:\LaptopBuild\IRC3325", 1)
   FileDelete("C:\LaptopBuild\Scripts")
   Run(@ComSpec & " /c " & "C:\del.bat", "")
   Exit
EndFunc