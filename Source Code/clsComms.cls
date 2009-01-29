VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsComms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   File:       clsComms.cls
'   Author:     Nicky Johns, March 2000
'   Purpose:    Class for a collection of clsCommunication objects
'               This class is used by the frmCommunicationConfiguration form.
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'   NCJ 7/3/00 - SRs2780,3110 Initial creation of module
'   NCJ 13/6/00 - SR 3591
'   TA 26/9/01: Changes following db audit
'   NCJ 18 Jun 03 - Bug 912 - Include check for same dates & TrialOffice (primary key violation)
'----------------------------------------------------------------------------------------'

Option Explicit

' The collection of clsCommunication objects
Private mcolComms As Collection

' For temporary use
Private moComm As clsCommunication

'----------------------------------------------------------------------------------------'
Public Sub AddItem(oComm As clsCommunication)
'----------------------------------------------------------------------------------------'
' Add an item to our collection
'----------------------------------------------------------------------------------------'

    mcolComms.Add oComm
    
End Sub

'----------------------------------------------------------------------------------------'
Public Function Item(sKey As String) As clsCommunication
'----------------------------------------------------------------------------------------'
' Get item corresponding to given Key
' Assume key is valid
'----------------------------------------------------------------------------------------'

    Set Item = mcolComms.Item(sKey)

End Function

'----------------------------------------------------------------------------------------'
Public Function Count() As Integer
'----------------------------------------------------------------------------------------'

    Count = mcolComms.Count

End Function

'----------------------------------------------------------------------------------------'
Public Function GetCurrentRecord() As clsCommunication
'----------------------------------------------------------------------------------------'
' Get the record that is "current", i.e. the one whose dates include current date
' NCJ 13/6/00 SR 3591 Ensure that tha TransferData flag is 1
'----------------------------------------------------------------------------------------'
Dim dblNow As Double

    ' Get today's date without the time
    dblNow = CDbl(CLng(CDbl(Now)))
    
    ' Initialise result
    Set GetCurrentRecord = Nothing
    
    For Each moComm In mcolComms
        ' Check effective date interval
        ' NCJ 13/6/00 - And the TransferData flag
        If moComm.DblEffectiveFrom <= dblNow And moComm.DblEffectiveTo >= dblNow _
         And moComm.TransferData = 1 Then
            Set GetCurrentRecord = moComm
            Exit For
        End If
    Next

End Function

'----------------------------------------------------------------------------------------'
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
'----------------------------------------------------------------------------------------'
' The enumerator to allow For...Each iterations
'----------------------------------------------------------------------------------------'
    
    Set NewEnum = mcolComms.[_NewEnum]

End Function

'----------------------------------------------------------------------------------------'
Private Sub Class_Initialize()
'----------------------------------------------------------------------------------------'

    Set mcolComms = Nothing
    Set moComm = Nothing
    Set mcolComms = New Collection
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub Class_Terminate()
'----------------------------------------------------------------------------------------'

    Set mcolComms = Nothing
    Set moComm = Nothing

End Sub

'----------------------------------------------------------------------------------------'
Public Function ValidateCommRecord(oComm As clsCommunication) As String
'----------------------------------------------------------------------------------------'
' Returns error message if oComm conflicts with others in collection
' i.e. if date interval overlaps another one
' Returns empty string if OK
' NB Only checks if key fields have changed
' NCJ 18 Jun 03 - Check for an existing exact duplicate (active or inactive)
'----------------------------------------------------------------------------------------'
Dim sMsg As String
Dim dblFrom As Double
Dim dblTo As Double
Dim dblNow As Double

    sMsg = ""
    ' Check if anything's changed
    If oComm.HasKeyChanged Then
    
        dblFrom = oComm.DblEffectiveFrom
        dblTo = oComm.DblEffectiveTo
        
        ' Exclude time info
        dblNow = CDbl(CLng(CDbl(Now)))
        
        ' Check that To is after From
        If dblFrom > dblTo Then
            sMsg = "The Effective To date must be after the Effective From date."
        
        ' NCJ 13/6/00 Only check "active" records
        ElseIf oComm.TransferData = 1 Then
            For Each moComm In mcolComms
                ' Look for different key (i.e. don't compare with ourselves)
                ' NCJ 13/6/00 Only check "active" records (TransferData = 1)
                If moComm.CommKey <> oComm.CommKey And moComm.TransferData = 1 Then
                    If ValueWithin(moComm.DblEffectiveFrom, dblFrom, dblTo) _
                     Or ValueWithin(moComm.DblEffectiveTo, dblFrom, dblTo) _
                     Or (moComm.DblEffectiveFrom <= dblFrom And moComm.DblEffectiveTo >= dblTo) Then
                        sMsg = "This setting is not valid because its dates overlap " & vbNewLine
                        sMsg = sMsg & "with an existing active communication setting." & vbNewLine & vbNewLine
                        sMsg = sMsg & "(Set 'No data transfer' to deactivate a communication setting.)"
                        Exit For
                    End If
                End If
            Next
            
        End If
        
        ' NCJ 18 Jun 03 - Include check for exact duplicate records
        If sMsg = "" Then
            For Each moComm In mcolComms
                ' Look for different key (i.e. don't compare with ourselves)
                If moComm.CommKey <> oComm.CommKey Then
                    If moComm.DblEffectiveTo = dblTo _
                     And moComm.DblEffectiveFrom = dblFrom _
                     And moComm.TrialOffice = oComm.TrialOffice Then
                        sMsg = "This setting is not valid because its dates are identical to" & vbNewLine
                        sMsg = sMsg & "an existing communication setting for study office '" _
                                & oComm.TrialOffice & "'." & vbNewLine & vbNewLine
                        sMsg = sMsg & "Please change at least one of the dates."
                        Exit For
                    End If
                End If
            Next
        End If
        
    End If
    ValidateCommRecord = sMsg

End Function

'----------------------------------------------------------------------------------------'
Private Function ValueWithin(dblDX As Double, dblD1 As Double, dblD2 As Double) As Boolean
'----------------------------------------------------------------------------------------'
' Return TRUE if dblDX lies within the interval dblD1 to dblD2 (inclusive)
' Assume dblD1 =< dblD2
'----------------------------------------------------------------------------------------'

    If dblDX >= dblD1 And dblDX <= dblD2 Then
        ValueWithin = True
    Else
        ValueWithin = False
    End If

End Function

'---------------------------------------------------------------------
Public Sub Load()
'---------------------------------------------------------------------
' Create the collection of all existing Communication records
' from the TrialOffice table in the MACRO DB
'---------------------------------------------------------------------
Dim rsSettings As ADODB.Recordset
Dim sSQL As String
Dim sKey As String

    On Error GoTo ErrHandler
   
    ' Read the set of available config settings
    sSQL = "SELECT * FROM TrialOffice"
    Set rsSettings = New ADODB.Recordset
    rsSettings.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
     
    With rsSettings
        Do While Not .EOF
            Set moComm = New clsCommunication
            
            ' NCJ 8/3/00 - Allow previously stored data to be accepted
            ' (these properties now have more stringent error checks)
            On Error Resume Next
            moComm.TrialOffice = .Fields("TrialOffice")
            moComm.StoredTrialOffice = .Fields("TrialOffice")
            
            moComm.Site = RemoveNull(.Fields("Site"))
            moComm.StoredSite = RemoveNull(.Fields("Site"))
            
            On Error GoTo ErrHandler
            
            moComm.DblEffectiveFrom = .Fields("EffectiveFrom")
            moComm.StoredEffectiveFrom = .Fields("EffectiveFrom")
            
            moComm.DblEffectiveTo = .Fields("EffectiveTo")
            moComm.StoredEffectiveTo = .Fields("EffectiveTo")
            
            moComm.HTTPAddress = RemoveNull(.Fields("HTTPAddress"))
           
            moComm.User = RemoveNull(.Fields("UserName"))
            moComm.Password = RemoveNull(.Fields("UserPassword"))
            moComm.ProxyServer = RemoveNull(.Fields("ProxyServer"))
            
            moComm.TransferData = RemoveNull(.Fields("TransferData"))
            ' NCJ 13/6/00
            moComm.StoredTransferData = RemoveNull(.Fields("TransferData"))
            If Not IsNull(.Fields("PortNumber")) Then
                moComm.PortNumber = .Fields("PortNumber")
            End If
            
            ' Add it to our collection
            moComm.CommKey = moComm.TrialOffice & moComm.DblEffectiveFrom & moComm.DblEffectiveTo
            mcolComms.Add moComm, moComm.CommKey
            
            ' Flag that it's new
            moComm.HasChanges = False
            
            .MoveNext
        Loop
    End With
    
    rsSettings.Close
    Set rsSettings = Nothing
        
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "Load", "clsComms")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   

End Sub

