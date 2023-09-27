Attribute VB_Name = "Sheet2"
Option Explicit

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

' ==========================================================================================================
' Procedure Name : Worksheet_SelectionChange
' Procedure Info :
'
' Procedure Access: Private
' Parameter Target (Range)
' Author: Luc Van Overveldt  Create date: 29/03/2023
' ==========================================================================================================
  Const DBG_FNCNAME = "Worksheet_SelectionChange": Const DBG_MODNAME = "Sheet2": Const C_ERR_VERBOSE = False: Dim errMess As String
  On Error GoTo exit_with_error
  Dim retVal As Integer
  Dim newRng As Range


   If (Target.Address = "$B$3") Then
      'Target.Value = "Running..."
      'Call LibPlan_ButtonRefresh_Exec(Target)
      Set newRng = Target.Offset(0, 1)
      newRng.Activate
      Target.Value = "Refresh"
    End If
    

    
end_of_function:
   ' Worksheet_SelectionChange  = retVal
 

 Exit Sub


exit_with_error:
  If (errMess = "") Then errMess = Err.Description Else errMess = errMess + vbCrLf + Err.Description
  Dim ErrContext As String: ErrContext = "Error in " + DBG_FNCNAME + ", line " & Erl & "."
  'Call Lib_WTPlan_DebugWrite_Error(errMess, DBG_MODNAME,DBG_FNCNAME , C_ERR_VERBOSE,errContext)
  If (C_ERR_VERBOSE = True) Then Debug.Print "== Error == " + DBG_MODNAME + "." + DBG_FNCNAME + "  : " + errMess
  
  GoTo end_of_function

 


End Sub


