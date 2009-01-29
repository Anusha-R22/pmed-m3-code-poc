VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmActiveDirectoryServers 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Active Directory Servers"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8475
   Icon            =   "frmActiveDirectoryServers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   8475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   4500
      TabIndex        =   1
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5820
      TabIndex        =   2
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   7140
      TabIndex        =   3
      Top             =   5040
      Width           =   1215
   End
   Begin MSComctlLib.ListView ADListview 
      Height          =   4875
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   8599
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmActiveDirectoryServers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
'   File:       frmActiveDirectoryServers.frm
'   Copyright:  InferMed Ltd. 2002. All Rights Reserved
'   Author:     I Curtis, December 2005
'   Purpose:    Active Directory servers management form
'------------------------------------------------------------------------------
' REVISIONS
'------------------------------------------------------------------------------

Private Enum eADLoginType
    Default = 0
    DefaultWithPath = 1
    PathUserNamePasswordStored = 2
    PathUserNamePasswordEntered = 3
End Enum

Option Explicit

'---------------------------------------------------------------------
Private Sub cmdADD_Click()
'---------------------------------------------------------------------
' ic 06/12/2005
' add a server to the list
'---------------------------------------------------------------------
    Call frmActiveDirectoryNewServer.Show(vbModal)
    Call RefreshADServerList
End Sub

'---------------------------------------------------------------------
Private Sub cmdClose_Click()
'---------------------------------------------------------------------
' ic 06/12/2005
' close the form
'---------------------------------------------------------------------
    Unload Me
End Sub

'---------------------------------------------------------------------
Private Sub cmdRemove_Click()
'---------------------------------------------------------------------
' ic 06/12/2005
' remove a server from the list
'---------------------------------------------------------------------
Dim liADServer As ListItem
Dim n As Integer
Dim oCon As ADODB.Connection
Dim sSQL As String


    On Error GoTo ErrLabel
    
    HourglassOn
    
    Set liADServer = ADListview.SelectedItem
    If (Not IsNull(liADServer)) Then
        
        If (DialogQuestion("Are you sure you want to delete this Active Directory Server?") = vbYes) Then
        
            'open db connection
            Set oCon = New ADODB.Connection
            oCon.Open (SecurityADODBConnection)
            
            'delete the server
            sSQL = "DELETE FROM ACTIVEDIRECTORYSERVERS WHERE CONNECTORDER = " & liADServer.Tag
            oCon.Execute sSQL
            ADListview.ListItems.Remove (liADServer.Index)
            Set liADServer = Nothing
            
            'shift all the other server orders about
            For n = 1 To ADListview.ListItems.Count
                If (ADListview.ListItems(n).Tag <> n) Then
                    sSQL = "UPDATE ACTIVEDIRECTORYSERVERS SET CONNECTORDER = " & n & " WHERE CONNECTORDER = " & ADListview.ListItems(n).Tag
                    oCon.Execute sSQL
                    ADListview.ListItems(n).Tag = n
                    ADListview.ListItems(n).Text = n
                End If
            Next
            
            'close db connection
            oCon.Close
            Set oCon = Nothing
            
            If (ADListview.ListItems.Count = 0) Then
                cmdRemove.Enabled = False
            End If
        End If
    End If
    
    HourglassOff
    Exit Sub
    
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "frmActiveDirectoryServers.cmdRemove_Click"
End Sub

'---------------------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------------------
' ic 06/12/2005
' form load
'---------------------------------------------------------------------
    
    On Error GoTo ErrLabel
    
    'add column headers
    ADListview.ColumnHeaders.Add 1, "co", "Connect Order", 600
    ADListview.ColumnHeaders.Add 2, "name", "Name", 1500
    ADListview.ColumnHeaders.Add 3, "path", "Path", 4000
    ADListview.ColumnHeaders.Add 4, "lt", "Login Type", 2000
    Call RefreshADServerList
    Exit Sub
    
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "frmActiveDirectoryServers.Form_Load"
End Sub

'---------------------------------------------------------------------
Private Sub RefreshADServerList()
'---------------------------------------------------------------------
' ic 06/12/2005
' refresh list of ad servers
'---------------------------------------------------------------------
Dim vADServers As Variant
Dim sPath As String
Dim n As Integer

    On Error GoTo ErrLabel
    
    'get server list from database
    vADServers = ADServers()
    
    'clear the list
    ADListview.ListItems.Clear
    
    'add server list to listview
    If (Not IsNull(vADServers)) Then
        For n = LBound(vADServers, 2) To UBound(vADServers, 2)
            ADListview.ListItems.Add n + 1, "k_" & vADServers(0, n), vADServers(0, n)
            If (IsNull(vADServers(1, n))) Then
                sPath = ""
            Else
                sPath = DecryptString(vADServers(1, n))
            End If
            ADListview.ListItems(n + 1).ListSubItems.Add 1, , vADServers(3, n)
            ADListview.ListItems(n + 1).ListSubItems.Add 2, , sPath
            ADListview.ListItems(n + 1).ListSubItems.Add 3, , RtnADLoginTypeString(CInt(vADServers(2, n)))
            ADListview.ListItems(n + 1).Tag = vADServers(0, n)
        Next
        
        cmdRemove.Enabled = True
    End If
    Exit Sub
    
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "frmActiveDirectoryServers.RefreshADServerList"
End Sub

'---------------------------------------------------------------------
Private Function RtnADLoginTypeString(nADLoginType As Integer)
'---------------------------------------------------------------------
' ic 07/12/2005
' return active directory login type string
'---------------------------------------------------------------------
    Select Case nADLoginType
    Case eADLoginType.Default:
        RtnADLoginTypeString = "Current domain with logged in users credentials"
    Case eADLoginType.DefaultWithPath:
        RtnADLoginTypeString = "Given domain with logged in users credentials"
    Case eADLoginType.PathUserNamePasswordStored:
        RtnADLoginTypeString = "Given domain with stored username and password"
    Case eADLoginType.PathUserNamePasswordEntered:
        RtnADLoginTypeString = "Given domain with entered username and password"
    End Select
End Function


'---------------------------------------------------------------------
Private Function ADServers() As Variant
'---------------------------------------------------------------------
' ic 06/12/2005
' returns an array of all the active directory servers
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsADServers As ADODB.Recordset
Dim vADServers As Variant

    On Error GoTo ErrLabel

    'build sql
    sSQL = "SELECT CONNECTORDER, PATH, LOGINTYPE, NAME FROM ACTIVEDIRECTORYSERVERS ORDER BY CONNECTORDER"
    Set rsADServers = New ADODB.Recordset
    rsADServers.Open sSQL, SecurityADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsADServers.RecordCount > 0 Then
        vADServers = rsADServers.GetRows
    Else
        vADServers = Null
    End If
    
    ADServers = vADServers
    
    rsADServers.Close
    Set rsADServers = Nothing
    Exit Function

ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "frmActiveDirectoryServers.ADServers"
End Function

