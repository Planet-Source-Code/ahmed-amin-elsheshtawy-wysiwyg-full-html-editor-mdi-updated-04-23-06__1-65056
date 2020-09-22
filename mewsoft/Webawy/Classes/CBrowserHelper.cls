VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBrowserHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' interfaces
Implements IObjectWithSite

'apis
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

' private variables
Private moUnknown As IUnknownVB
Private WithEvents moBrowser As InternetExplorer
Attribute moBrowser.VB_VarHelpID = -1

' called when the parent wants
' to know what we're aiming at
Private Sub IObjectWithSite_GetSite(ByVal priid As VBShellLib.REFIID, ppvObj As VBShellLib.VOID)

   ' return the interface we've got
   If Not (moUnknown Is Nothing) Then
      moUnknown.QueryInterface priid, ppvObj
   End If

End Sub

' called on init to give us a
' pointer to the parent browser
Private Sub IObjectWithSite_SetSite(ByVal pSite As VBShellLib.IUnknownVB)

   ' store the unknown
   Set moUnknown = pSite
   
   ' since this method is called again
   ' when the BHO is unloaded, we have to
   ' release our pointer
   If ObjPtr(pSite) = 0 Then
      CopyMemory moBrowser, 0&, 4
   Else
      ' get the IWebBrowser2 interface
      Set moBrowser = moUnknown
   End If

End Sub

' browser events below :)
Private Sub moBrowser_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)

   ' just for testing..
   MsgBox CStr(URL)

End Sub

Private Sub IOleCommandTarget_Exec(ByVal pguidCmdGroup As VBShellLib.LPGUID, ByVal nCmdID As VBShellLib.DWORD, ByVal nCmdExecOpt As VBShellLib.DWORD, pvaIn As Variant, pvaOut As Variant)
    MsgBox nCmdID
End Sub

Private Sub IOleCommandTarget_QueryStatus(ByVal pguidCmdGroup As VBShellLib.LPGUID, ByVal cCmds As VBShellLib.ULONG, ByVal prgCmds As VBShellLib.LPOLECMD, ByVal pCmdText As VBShellLib.LPOLECMDTEXT)

End Sub

