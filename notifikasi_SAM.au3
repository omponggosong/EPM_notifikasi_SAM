;mustaqim indrawan
;notifikasi SO di SAM

#include <IE.au3>
#include <Array.au3>
#include <Crypt.au3>
#include <GUIConstantsEx.au3>
#include <File.au3>

Const $Key = "terong"


While 1
Global $filepath = (@ScriptDir & "\config.txt")
Global $bacafile = FileOpen($filepath, $FO_READ)
Global $txtusername = FileReadLine($bacafile, 1)
Global $txtpassword = FileReadLine($bacafile, 2)
Global $second = FileReadLine($bacafile, 3)
Global $delayrefresh = $second * 1000
Global $timeoutmsgbox = 5
Global $aktiftitleie = WinGetTitle(":: Sales Approval Management [SAM] - Oracle :: - Internet Explorer","")
Global $hndIE = IsObj(WinGetHandle( ":: Sales Approval Management [SAM] - Oracle :: - Internet Explorer",":: Sales Approval Management [SAM] - Oracle :: - Internet Explorer"))
Global $ambildata = _IEAttach($hndIE,"html")
Global $statuslogin = IsObj(_IEGetObjById($ambildata, "lblStatusApproval"))
Global $evenvalidasi = _IEGetObjById($ambildata, "__EVENTVALIDATION")
Global $samexist = (WinExists(":: Sales Approval Management [SAM] - Oracle :: - Internet Explorer"))
Global $sudahlogin = ($samexist And $statuslogin)
Global $user = BinaryToString(_Crypt_DecryptData($txtusername, $Key, $CALG_3DES))
Global $pass = BinaryToString(_Crypt_DecryptData($txtpassword, $Key, $CALG_3DES))
Global $second = FileReadLine($bacafile, 3)
Global $delayrefresh = $second * 1000
Global $timeoutmsgbox = 5


If $sudahlogin = 1 Then
   refreshlistSO()
Else
   If $samexist = 1 Then
	  loginSAM()
   Else
	  GUI()
	  bukaIE()
   EndIf
EndIf
Sleep($delayrefresh)
WEnd

Func bukaIE()
   $url = _IECreate("https://sam.enseval.com/")
   $txtusernamex = _IEGetObjById($url, "TxtUser")
   _IEFormElementSetValue($txtusernamex, $user)
   $txtpasswordx = _IEGetObjById($url, "TxtPassword")
   _IEFormElementSetValue($txtpasswordx, $pass)
   $btnlogon = _IEGetObjById($url, "BtnLogon")
   _IEAction($btnlogon, "click")
EndFunc

Func tungguIE()
   WinWait(":: Sales Approval Management [SAM] - Oracle :: - Internet Explorer")
   WinActivate(":: Sales Approval Management [SAM] - Oracle :: - Internet Explorer")
EndFunc

Func loginSAM()
   $txtusernamex = _IEGetObjById($ambildata, "TxtUser")
   _IEFormElementSetValue($txtusernamex, $user)
   $txtpasswordx = _IEGetObjById($ambildata, "TxtPassword")
   _IEFormElementSetValue($txtpasswordx, $pass)
   $btnlogon = _IEGetObjById($ambildata, "BtnLogon")
   _IEAction($btnlogon, "click")
EndFunc

Func refreshlistSO()
   $findcheckbox = _IEGetObjById($ambildata, "RadGrid1_ctl00_ctl04_SelectColumnSelectCheckBox")
   $idtable = _IEGetObjById($ambildata, "RadGrid1_ctl00")
   $ambiltable = _IETableGetCollection($idtable, 2)
   $tablearray = _IETableWriteToArray($ambiltable, True)
   If WinExists(":: Sales Approval Management [SAM] - Oracle :: - Internet Explorer") And $statuslogin <> 0 Then
	  $btnrefresh = _IEGetObjById($ambildata, "btnRefresh")
	  _IEAction($btnrefresh, "click")
	  Sleep(3000)
	  If $findcheckbox <> 0 Then
		 MsgBox(0, "", "ada SO" & @CRLF & "OTA :" & $tablearray[2][3] & @CRLF & "Cust.Name :" & $tablearray[2][6], $timeoutmsgbox)
		 tungguIE()
	  EndIf
   EndIf
EndFunc

Func klikproses()
   $cariproses = _IEFormElementGetValue(_IEGetObjById($ambildata, "btnSubmit"))
   _IEAction($cariproses, "click")
EndFunc

Func GUI()
    ; Create a GUI with various controls.
    Local $hGUI = GUICreate("Login Disini!", 170, 100)
	GUICtrlCreateLabel("Username", 5, 8)
	GUICtrlCreateLabel("Password", 5, 33)
	Global $txtusername1 = GUICtrlCreateInput("", 60, 5, 100)
	Global $txtpassword1 = GUICtrlCreateInput("", 60, 30, 100)
	Global $OK = GUICtrlCreateButton("OK", 120, 55, 40)
    ; Display the GUI.
    GUISetState(@SW_SHOW, $hGUI)
While 1
        Switch GUIGetMsg()
		Case $GUI_EVENT_CLOSE, $OK
			$fileopen = FileOpen($filepath, $FO_OVERWRITE)
			$dapatuser = GUICtrlRead($txtusername1)
			$enkripuser = _Crypt_EncryptData($dapatuser, $Key, $CALG_3DES)
			FileWriteLine($fileopen, $enkripuser)
			FileWriteLine($fileopen, "")
			$dapatpass = GUICtrlRead($txtpassword1)
			$enkrippass = _Crypt_EncryptData($dapatpass, $Key, $CALG_3DES)
			FileWriteLine($fileopen, $enkrippass)
			FileWriteLine($fileopen, "")
			FileWriteLine($fileopen, "30")
			ExitLoop

        EndSwitch
    WEnd
 GUIDelete($hGUI)
EndFunc