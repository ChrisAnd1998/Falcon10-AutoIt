;~ Example 13 Details about the right pane of the windows explorer
#include "CUIAutomation2.au3"
#include <MsgBoxConstants.au3>
#include <File.au3>
#include <WinAPI.au3>
#NoTrayIcon

Opt("TrayMenuMode", 3)
Opt( "MustDeclareVars", 1 )


Global $oUIAutomation
Global $CountXXX
Global $Lefts
Global $siz


;_ProcessCloseOtherEx(@ScriptName)



 _Restart_Explorer()

 sleep(1000)

MainFunc()


Func Center()



   Local $iFullDesktopWidth = @DesktopWidth



   Local $ico = $CountXXX * $siz[2] + 1 - $siz[3]

   Local $half = $ico / 2
   Local $x = $iFullDesktopWidth / 2
   Local $x2 = $x - $half
Local $x3 = $x2 - $Lefts[0]




ControlMove("[CLASS:Shell_TrayWnd]", "", "[CLASS:MSTaskListWClass; INSTANCE:1]", 0, 0, 3000, 40)
ControlMove("[CLASS:Shell_TrayWnd]", "", "[CLASS:ReBarWindow32; INSTANCE:1]", $x2, 0, $ico + 50, 40)


EndFunc

Func MainFunc()



     TraySetToolTip("Falcon10")


ControlMove("[CLASS:Shell_TrayWnd]", "", "[CLASS:MSTaskListWClass; INSTANCE:1]", 0, 0, 3000, 40)
 $Lefts = ControlGetPos("[CLASS:Shell_TrayWnd]", "", "[CLASS:MSTaskListWClass; INSTANCE:1]")

 $siz = ControlGetPos("[CLASS:Shell_TrayWnd]", "", "[CLASS:Start; INSTANCE:1]")

  Local $hWindow = ControlGetHandle( "[CLASS:Shell_TrayWnd]","","[CLASS:MSTaskListWClass; INSTANCE:1]")
  $oUIAutomation = ObjCreateInterface( $sCLSID_CUIAutomation, $sIID_IUIAutomation, $dtagIUIAutomation )

  Local $pWindow
   $oUIAutomation.ElementFromHandle( $hWindow, $pWindow )

  Local $oWindow = ObjCreateInterface( $pWindow, $sIID_IUIAutomationElement, $dtagIUIAutomationElement )

While 1
   Center()
   sleep(500)
ListDescendants( $oWindow, 0, 1 )
Center()
Wend

EndFunc

Func _Restart_Explorer()
    Local $ifailure = 100, $zfailure = 100, $rPID = 0, $iExplorerPath = @WindowsDir & "\Explorer.exe"
    _WinAPI_ShellChangeNotify($shcne_AssocChanged, 0, 0, 0) ; Save icon positions
    Local $hSystray = _WinAPI_FindWindow("Shell_TrayWnd", "")
    _SendMessage($hSystray, 1460, 0, 0) ; Close the Explorer shell gracefully
    While ProcessExists("Explorer.exe") ; Try Close the Explorer
        Sleep(10)
        $ifailure -= ProcessClose("Explorer.exe") ? 0 : 1
        If $ifailure < 1 Then Return SetError(1, 0, 0)
    WEnd
;~  _WMI_StartExplorer()
    While (Not ProcessExists("Explorer.exe")) ; Start the Explorer
        If Not FileExists($iExplorerPath) Then Return SetError(-1, 0, 0)
        Sleep(500)
        $rPID = ShellExecute($iExplorerPath)
        $zfailure -= $rPID ? 0 : 1
        If $zfailure < 1 Then Return SetError(2, 0, 0)
    WEnd
    Return $rPID
EndFunc   ;==>_Restart_Explorer

Func _ProcessCloseOtherEx($sPID)
    If IsString($sPID) Then $sPID = ProcessExists($sPID)

    If Not $sPID Then Return SetError(1, 0, 0)

    If $sPID <> @AutoItPID Then
        Run(@ComSpec & " /c taskkill /F /PID " & $sPID & " /T", @SystemDir, @SW_HIDE)
        _ProcessCloseOtherEx($sPID)
    Else
        Return
    EndIf
EndFunc   ;==>_ProcessCloseOtherEx

Func ListDescendants( $oParent, $iLevel, $iLevels = 0 )
  If Not IsObj( $oParent ) Then Return
  If $iLevels And $iLevel = $iLevels Then Return

  Local $pRawWalker, $oRawWalker
  $oUIAutomation.RawViewWalker( $pRawWalker )
  $oRawWalker = ObjCreateInterface( $pRawWalker, $sIID_IUIAutomationTreeWalker, $dtagIUIAutomationTreeWalker )

  Local $pUIElement, $oUIElement
  $oRawWalker.GetFirstChildElement( $oParent, $pUIElement )
  $oUIElement = ObjCreateInterface( $pUIElement, $sIID_IUIAutomationElement, $dtagIUIAutomationElement )

  Local $sIndent = ""
  For $i = 0 To $iLevel - 1
    $sIndent &= "    "
 Next



$CountXXX = 0

  While IsObj( $oUIElement )
   ; ConsoleWrite( $sIndent & _UIA_getPropertyValue( $oUIElement, $UIA_NamePropertyId ) & @CRLF)
	$CountXXX = $CountXXX + 1
   ; ListDescendants( $oUIElement, $iLevel + 1, $iLevels )
   $oRawWalker.GetNextSiblingElement( $oUIElement, $pUIElement )
   $oUIElement = ObjCreateInterface( $pUIElement, $sIID_IUIAutomationElement, $dtagIUIAutomationElement )
WEnd

;ConsoleWrite($CountXXX -1 & @CRLF)


EndFunc



Func _UIA_getPropertyValue( $obj, $id )
  Local $tVal
  $obj.GetCurrentPropertyValue( $id, $tVal )
  If Not IsArray( $tVal ) Then Return $tVal
  Local $tStr = $tVal[0]
  For $i = 1 To UBound( $tVal ) - 1
    $tStr &= "; " & $tVal[$i]
  Next
  Return $tStr
EndFunc


