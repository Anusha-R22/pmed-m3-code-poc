Attribute VB_Name = "modAzRbMain"

'----------------------------------------------------------------------------------------
'   File:       modAzRbMain.bas
'   Copyright:  InferMed Ltd. 2006. All Rights Reserved
'   Author:     Nicky Johns, May 2006
'   Purpose:    To be used in the AREZZO Rebuild DLL
'       Contains a selection of routines COPIED and adapted from MACRO SD
'----------------------------------------------------------------------------------------
'Revisions:
'   NCJ 8 May 06 - Initial development
'
'----------------------------------------------------------------------------------------

Option Explicit

Public Const valAlpha                   As Integer = 1
Public Const valNumeric                 As Integer = 2
Public Const valSpace                   As Integer = 4
Public Const valOnlySingleQuotes        As Integer = 8
Public Const valComma                   As Integer = 16
Public Const valUnderscore              As Integer = 32
Public Const valDateSeperators          As Integer = 64
' PN 02/09/99
' allow includusion of mathematical operators in the check string
Public Const valMathsOperators          As Integer = 128
'Mo Morris 15/4/2002 valDecimalPoint added
Public Const valDecimalPoint            As Integer = 256

Private moMacroADODBConnection As ADODB.Connection
Private msDBCon As String

'--------------------------------------------------------------------------------
Public Property Get MacroConnectionString() As String
'--------------------------------------------------------------------------------
' Connection string - assume already set up
'--------------------------------------------------------------------------------

    MacroConnectionString = msDBCon

End Property

'--------------------------------------------------------------------------------
Public Property Get MacroADODBConnection() As ADODB.Connection
'--------------------------------------------------------------------------------

    If moMacroADODBConnection.State = adStateClosed Then
         InitializeMacroADODBConnection (msDBCon)
    End If
    
    Set MacroADODBConnection = moMacroADODBConnection

End Property

'--------------------------------------------------------------------------------
Public Sub InitializeMacroADODBConnection(sCon As String)
'--------------------------------------------------------------------------------
'This will initialize the ado connection to the database
'--------------------------------------------------------------------------------
   
    msDBCon = sCon
    ' Terminate first if necessary
    If Not moMacroADODBConnection Is Nothing Then
        moMacroADODBConnection.Close
        Set moMacroADODBConnection = Nothing
    End If
    
    Set moMacroADODBConnection = New ADODB.Connection
    moMacroADODBConnection.Open sCon
    moMacroADODBConnection.CursorLocation = adUseClient

End Sub

'--------------------------------------------------------------------------------
Public Sub TerminateMacroADODBConnection()
'--------------------------------------------------------------------------------
' Close the macro connection if it is open
'--------------------------------------------------------------------------------

    If Not moMacroADODBConnection Is Nothing Then
        moMacroADODBConnection.Close
        Set moMacroADODBConnection = Nothing
    End If

End Sub

'---------------------------------------------------------------------
Public Function MACROCodeErrorHandler(nTrappedErrNum As Long, sTrappedErrDesc As String, _
                            sProcName As String, sModuleName As String) As OnErrorAction
'---------------------------------------------------------------------
' Just raise the error.
'---------------------------------------------------------------------
    
    Err.Raise nTrappedErrNum, sTrappedErrDesc
    
End Function

'---------------------------------------------------------------------
Public Sub ExitMACRO()
'---------------------------------------------------------------------
' NCJ May 06
' Dummy routine for this DLL
'---------------------------------------------------------------------

End Sub

'---------------------------------------------------------------------
Public Sub MACROEnd()
'---------------------------------------------------------------------
' NCJ May 06
' Dummy routine for this DLL
'---------------------------------------------------------------------

End Sub

