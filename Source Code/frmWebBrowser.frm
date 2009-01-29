VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmWebBrowser 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   2925
   ClientLeft      =   8445
   ClientTop       =   5205
   ClientWidth     =   3285
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   2925
   ScaleWidth      =   3285
   ShowInTaskbar   =   0   'False
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   2805
      Left            =   15
      TabIndex        =   0
      Top             =   -15
      Width           =   3195
      ExtentX         =   5636
      ExtentY         =   4948
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmWebBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2002. All Rights Reserved
'   File:       frmWebBrowser.frm
'   Author:     Toby Aldridge Sept 2002
'   Purpose:    Display a whole web page as specified in load
'                    and call corresponding function in JSFunctiosn on navigate
'----------------------------------------------------------------------------------------'
' Revisions:
' MLM 21/01/03: Decode %20s into spaces, particularly for subject labels.
' TA 4 Mar 03 - Removed WebBrowser_DownloadComplete as it crashes in a compiled executable
' TA 03/04/2003 - document.close now called after document.write - No need to unload webbrowser forms any more before refreshing them
' RS 11/06/2003 - BUG1852 Catch error 438 when viewing excel report (ASP Reports) in WriteHTML
' ic 20/11/2006 added check for ie7 format url
'----------------------------------------------------------------------------------------'

Option Explicit


'enum for whether we are loading a page or setting the page's text
Public Enum eWebDisplayType
    wdtUrl = 0
    wdtHTML = 1
End Enum


'remember scrolloption
Private msScrollOption As String

'----------------------------------------------------------------------------------------'
Private Sub SetURL(sURL As String)
'----------------------------------------------------------------------------------------'
' navigate to specifed URL
'----------------------------------------------------------------------------------------'
    
    If sURL = "" Then
        WebBrowser.Navigate "about:blank"
    Else
        WebBrowser.Navigate sURL
    End If

End Sub

Public Sub WriteHTML(sHTML As String)
'----------------------------------------------------------------------------------------'
' write specifed HTML
'----------------------------------------------------------------------------------------'

    ' RS 11/06/2003: BUG1852 Allow for error 438 when showing excel report in browser (ASP Reports)
    On Error GoTo ErrHandler
    
    WebBrowser.Document.Write sHTML
    'TA 03/04/2003: close after writing
    WebBrowser.Document.Close

'   WebBrowser.Document.Script.Document.Clear
'   WebBrowser.Document.Script.Document.Write sHTML
'   WebBrowser.Document.Script.Document.Close

Exit Sub
ErrHandler:
    If Err.Number = 438 Then
        ' Ignore this error
        On Error Resume Next
        'WebBrowser.Document.Close
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub Form_Resize()
'----------------------------------------------------------------------------------------'

    If Me.BorderStyle = vbBSNone Then
        WebBrowser.Left = -30
        WebBrowser.Top = -30
        WebBrowser.Height = Me.ScaleHeight + 60
        WebBrowser.Width = Me.ScaleWidth + 60
    Else
        WebBrowser.Left = 0
        WebBrowser.Top = 0
        WebBrowser.Height = Me.ScaleHeight
        WebBrowser.Width = Me.ScaleWidth
    End If
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub WebBrowser_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
'----------------------------------------------------------------------------------------'
' intercept navigations beginning with "\\Xfn" and call coressonding JSFunction
' revisions
' ic 20/11/2006 added check for ie7 format url
'----------------------------------------------------------------------------------------'
Dim vFunctionDetails As Variant
Dim oJs As JSFunctions
Dim sFN As String
Static bCurrentlyProcessing
    
On Error GoTo Errorlabel
    
        'ic 20/11/2006 added check for ie7 format url
        'we are already doing something
        If Left(URL, 15) = "about:blankVBfn" _
        Or Left(URL, 10) = "about:VBfn" Then
            'stop navigation
            Cancel = True
            If Not bCurrentlyProcessing Then
                'show we are doing something
                bCurrentlyProcessing = True
                'MLM 21/01/03: Decode %20s into spaces, particularly for subject labels.
                vFunctionDetails = Split(Replace(URL, "%20", " "), "|")
                
                'ic 20/11/2006 added check for ie7 format url
                If Left(URL, 15) = "about:blankVBfn" Then
                    sFN = Mid(vFunctionDetails(eFunctionParams.fpFunctionName), 14)
                Else
                    sFN = Mid(vFunctionDetails(eFunctionParams.fpFunctionName), 9)
                End If
                
                If sFN = "fnClose" Then
                    'fnClose is a special one that always cloes the window
                    Unload Me
                    Exit Sub
                End If
                
                Set oJs = New JSFunctions
                oJs.Init Me
                Call CallByName(oJs, sFN, VbMethod, vFunctionDetails)
                Set oJs = Nothing
                bCurrentlyProcessing = False
            End If
        End If
    
Exit Sub

Errorlabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "WebBrowser_BeforeNavigate2", Err.Source) = Retry Then
        Resume
    End If
    
End Sub


'----------------------------------------------------------------------------------------'
Public Sub Display(enDisplayType As eWebDisplayType, sHTMLURL As String, _
                        Optional sScrollOption As String = "", Optional bModal As Boolean = False, Optional sTitle As String = "", Optional sPostData As String = "")
'----------------------------------------------------------------------------------------'
' show the form and load the given URL if not ""
'scrolloption is "yes","no","auto" or "" for no change
'----------------------------------------------------------------------------------------'

    Load Me
    
'moved modeless show to end of routine - move back if problems
'    If Not bModal Then
'        Me.Show vbModeless
'    End If
'
    msScrollOption = sScrollOption
    Me.Caption = sTitle

    'stop it tryuing to navigate somewhere
    Me.WebBrowser.Stop
    'blank out the navigation cancelled message
    WriteHTML ""

    If enDisplayType = wdtUrl Then
        'webbrowser.document.all.tags("base")(0).href = "file:///C:/VSS/MACRO 3.0/www/asp/general/AppFrm.htm"
        'WriteHTML "<base href='file:///C:/VSS/MACRO 3.0/www/asp/general/AppMenuLh.htm'>" & StringFromFile(sURL)
        If sPostData = "" Then
            SetURL sHTMLURL
        Else
        
'----------------------------------------------------------------------------------------'
            'little section copied from MS help to do post data
             
             Dim Flags As Long
             Dim TargetFrame As String
             Dim PostData() As Byte
             Dim Headers As String
             Flags = 0
             TargetFrame = ""
             PostData = sPostData
    
             ' VB creates a Unicode string by default so we need to
             ' convert it back to Single byte character set.
             PostData = StrConv(PostData, vbFromUnicode)
             Headers = "Content-Type: application/x-www-form-urlencoded" & vbCrLf
             WebBrowser.Navigate sHTMLURL, Flags, TargetFrame, PostData, Headers
             
'----------------------------------------------------------------------------------------'
        
        End If
        
    Else
        
        
        WriteHTML sHTMLURL
        
        WebBrowser.Document.body.Scroll = sScrollOption
    End If

    If bModal Then
        'ensure we appear completely on the screen
        If Me.Top < 100 Then
            Me.Top = 100
        End If
        If Me.Top + Me.Height > Screen.Height - 500 Then
            Me.Top = Screen.Height - Me.Height - 400
        End If
        If Me.Left < 100 Then
            Me.Left = 100
        End If
        If Me.Left + Me.Width > Screen.Width Then
            Me.Left = Screen.Width - Me.Width
        End If
        Me.Show vbModal
    Else
        Me.Show vbModeless
        Me.ZOrder
    End If

End Sub

'----------------------------------------------------------------------------------------'
Public Function ExecuteJavaScript() As Object
'----------------------------------------------------------------------------------------'

    'example of calling js function from vb
    
    Set ExecuteJavaScript = WebBrowser.Document.parentWindow '.fnLogOutUrl
 '   WebBrowser.Document.parentWindow.fnLogOutUrl

End Function

' TA 4 Mar 03 - This FAILS (crashes) in a compiled executable
' Not currently needed as F5 (refresh) is now disabled
'Private Sub WebBrowser_DownloadComplete()
'    WebBrowser.Document.body.Scroll = msScrollOption
'End Sub
