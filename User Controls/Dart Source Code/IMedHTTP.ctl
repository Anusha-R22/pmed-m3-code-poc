VERSION 5.00
Object = "{0E56FD71-943D-11D2-BA66-0040053687FE}#1.0#0"; "DartWeb.dll"
Begin VB.UserControl IMedHTTP 
   ClientHeight    =   960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1050
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   960
   ScaleWidth      =   1050
   Begin VB.Image Image1 
      Height          =   690
      Left            =   120
      Picture         =   "IMedHTTP.ctx":0000
      Top             =   120
      Width           =   750
   End
   Begin DartWebCtl.Http Http1 
      Left            =   600
      OleObjectBlob   =   "IMedHTTP.ctx":0607
      Top             =   480
   End
End
Attribute VB_Name = "IMedHTTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------------------
'   File:       IMedHTTP.ctl
'   Copyright:  InferMed Ltd., 2007. All Rights Reserved
'   Author:     Nicky Johns, February 2007
'   Purpose:    Simple wrapper for DART HTTP control for MACRO data transfer
'--------------------------------------------------------------------------------
' REVISIONS:
'   NCJ 20 Feb 07 - Initial development
'
'--------------------------------------------------------------------------------

Option Explicit

'--------------------------------------------------------------------------------
Public Property Get Version() As String
'--------------------------------------------------------------------------------

    Version = Http1.Version

End Property

'--------------------------------------------------------------------------------
Public Property Let Version(sVer As String)
'--------------------------------------------------------------------------------

    Http1.Version = sVer

End Property

'--------------------------------------------------------------------------------
Public Property Get Cache() As Boolean
'--------------------------------------------------------------------------------

    Cache = Http1.Cache

End Property

'--------------------------------------------------------------------------------
Public Property Let Cache(bCache As Boolean)
'--------------------------------------------------------------------------------

    Http1.Cache = bCache
    
End Property

'--------------------------------------------------------------------------------
Public Property Get Timeout() As Long
'--------------------------------------------------------------------------------

    Timeout = Http1.Timeout

End Property

'--------------------------------------------------------------------------------
Public Property Let Timeout(lMillisecs As Long)
'--------------------------------------------------------------------------------

    Http1.Timeout = lMillisecs
    
End Property

'--------------------------------------------------------------------------------
Public Property Get Proxy() As String
'--------------------------------------------------------------------------------

    Proxy = Http1.Proxy

End Property

'--------------------------------------------------------------------------------
Public Property Let Proxy(sProxy As String)
'--------------------------------------------------------------------------------

    Http1.Proxy = sProxy
    
End Property

'--------------------------------------------------------------------------------
Public Property Get URL() As String
'--------------------------------------------------------------------------------

    URL = Http1.URL

End Property

'--------------------------------------------------------------------------------
Public Property Let URL(sURL As String)
'--------------------------------------------------------------------------------

    Http1.URL = sURL

End Property

'--------------------------------------
Public Sub SetSecurity()
'--------------------------------------
' Set security to what MACRO likes
'--------------------------------------

    Http1.Security = httpAllowRedirectToHTTP + httpAllowRedirectToHTTPS

End Sub

'--------------------------------------------------------------------------
Public Sub Post(sURLParamData As String, sRetData As String, sUser As String, sPwd As String)
'--------------------------------------------------------------------------
' Post data using HTTP as done for MACRO data transfer
'--------------------------------------------------------------------------

    Http1.Post sURLParamData, , sRetData, , sUser, sPwd
    
End Sub

'--------------------------------------------------------------------------
Public Sub GetData(vData() As Byte, sUser As String, sPwd As String)
'--------------------------------------------------------------------------
' Get data using HTTP as done for MACRO data transfer
'--------------------------------------------------------------------------

    Http1.Get vData, , sUser, sPwd
    
End Sub


