VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmAbout"
   ClientHeight    =   4560
   ClientLeft      =   1635
   ClientTop       =   1755
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   7095
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   2595
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   3975
      ExtentX         =   7011
      ExtentY         =   4577
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
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
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'THIS IS A COPY OF FRMWEBBROWSER WITH MDI CHILD SET TO FALSE
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2002. All Rights Reserved
'   File:       frmWebNonMDI.frm
'   Author:     Toby Aldridge Sept 2002
'   Purpose:    Display a whole web page as specified in load
'                    and call corresponding function in JSFunctiosn on navigate
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
' ic 20/11/2006 added check for ie7 format url

Option Explicit


   
   


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

    WebBrowser.Document.Write sHTML


End Sub

'----------------------------------------------------------------------------------------'
Private Sub Form_Resize()
'----------------------------------------------------------------------------------------'

        WebBrowser.Left = 0
        WebBrowser.Top = 0
        WebBrowser.Height = Me.ScaleHeight + 40
        WebBrowser.Width = Me.ScaleWidth + 40
    
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub WebBrowser_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
'----------------------------------------------------------------------------------------'
' intercept navigations beginning with "\\Xfn" and call coressonding JSFunction
' revisions
' ic 20/11/2006 added check for ie7 format url
'----------------------------------------------------------------------------------------'
Dim vFunctionDetails As Variant
Dim sFN As String
Static bCurrentlyProcessing
    
On Error GoTo Errorlabel

        'ic 20/11/2006 added check for ie7 format url
        'we are already doing something
        If Left(URL, 15) = "about:blankVBfn" _
        Or Left(URL, 10) = "about:VBfn" Then
            'stop navigation
            Cancel = True
                'MLM 21/01/03: Decode %20s into spaces, particularly for subject labels.
                vFunctionDetails = Split(Replace(URL, "%20", " "), "|")
                'ic 20/11/2006 added check for ie7 format url
                If Left(URL, 15) = "about:blankVBfn" Then
                    sFN = Mid(vFunctionDetails(0), 14)
                Else
                    sFN = Mid(vFunctionDetails(0), 9)
                End If
                If sFN = "fnClose" Then
                    'fnClose is a special one that always cloes the window
                    Unload Me
                    Exit Sub
                End If
                

            End If
        
    
Exit Sub

Errorlabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "WebBrowser_BeforeNavigate2", Err.Source) = Retry Then
        Resume
    End If
    
End Sub


'----------------------------------------------------------------------------------------'
Public Sub Display()
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    Load Me
    Height = 5220
    Width = 7860
    
    Icon = frmMenu.Icon
    Caption = "About " & GetApplicationTitle
    FormCentre Me, frmMenu
    'stop it tryuing to navigate somewhere
    Me.WebBrowser.Stop

    WriteHTML GetAboutHTML
    WebBrowser.Document.body.Scroll = "no"

    Me.Show vbModal


End Sub

'----------------------------------------------------------------------------------------'
Public Function ExecuteJavaScript() As Object
'----------------------------------------------------------------------------------------'

    'example of calling js function from vb
    
    Set ExecuteJavaScript = WebBrowser.Document.parentWindow '.fnLogOutUrl
 '   WebBrowser.Document.parentWindow.fnLogOutUrl

End Function


