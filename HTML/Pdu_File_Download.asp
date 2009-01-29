<!--#include file=LocalSettings.txt-->
<%
'-----------------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2005 All Rights Reserved
'   File:       pdu_file_download.asp
'   Author:     David Hook, 2005
'   Purpose:    Used by TrialOffice for the purpose of downloading PDU entries identified in 
'		the Message table on a Server to a calling Client site.
'		This script is called for each PDU message to be downloaded.
'		The file is read using 'filename' parameter from the filesystem and a section of up to 'maxchunk' size is read.
'		The 'lastfilepos' parameter determines the reading start point within the file.
'		The variable 'gsSecurePdu' (LocalSettings.txt) determines the pdu file storage folder.
'-----------------------------------------------------------------------------------------------'
'
'-----------------------------------------------------------------------------------------------'
'   Revisions:
'-----------------------------------------------------------------------------------------------'
	' get filename
	sFileName=Request.QueryString("filename")
	' get last file position
	sFilePos=Request.QueryString("lastfilepos")
	' get maximum chunk size
	sMaxChunk=Request.QueryString("maxchunk")
	on error resume next
	lFilePos=clng(sFilePos)
	lMaxChunk=clng(sMaxChunk)
	' put in for debugging only
	on error goto 0
	
	' force to binary
	sContentType = "application/octet-stream"

    ' stream to browser using ADODB Stream object
    Set oStream = server.CreateObject("ADODB.Stream")
    oStream.Open
	oStream.Type = 1 ' 1 = adTypeBinary

	' check for existence of pdu setting
	if gsSecurePdu <> "" then
		' get pdu folder
		sPduDirectory = gsSecurePdu
	else
		' default to published html folder
		sPduDirectory = gsAppPath
	end if
	
    oStream.LoadFromFile sPduDirectory & sFileName 
    ' position stream
    oStream.Position = lFilePos
    ' calculate next chunk size
    ' bytes left 
    lChunkSize = oStream.Size - lFilePos 
    ' if bytes left greater than maximum chunk size
    if lChunkSize > lMaxChunk then
		' restrict to max chunk size
		lChunkSize = lMaxChunk
    end if
    ' Forces users to save a file to a location on their hard drive
    ' default as attachment (will open if can else will offer download dialog)
    Response.AddHeader "content-disposition", "attachment; filename=" & sFileName
	' set content-type
    Response.ContentType = sContentType 
    ' transmit chunk of file
    Response.AddHeader "Content-Length", lChunkSize + 3
    ' read stream
    If lChunkSize > 0 Then
        Response.BinaryWrite oStream.Read(lChunkSize)
        ' add footer success bytes
        Response.BinaryWrite StringToBinary("SUC")
    Else
        ' set empty footer
        Response.BinaryWrite StringToBinary("EMP")
    End If
    ' if has been an error then set error footer marking
    ' current 'chunk' invalid
    If Err.Number <> 0 Then
        ' set error footer
        Response.BinaryWrite StringToBinary("ERR")
    End If
    oStream.Close
    Set oStream = Nothing
    Response.End 
    
Function GetMIMEType(sFileName)
    ' get correct MIME type for file!
    Select Case Right(sFileName, 4)
        Case ".jpg", ".jpe", ".jpeg"
            GetMIMEType = "image/jpeg"
        Case ".gif"
            GetMIMEType = "image/gif"
        Case ".bmp"
            GetMIMEType = "image/bmp"
        Case ".png"
            GetMIMEType = "image/png"
        Case ".mpeg", ".mpg", ".mpe"
            GetMIMEType = "video/mpeg"
        Case ".mp3"
            GetMIMEType = "audio/x-mpeg"
        Case ".wav"
            GetMIMEType = "audio/x-wav"
        Case ".avi"
            GetMIMEType = "video/msvideo"
        Case ".mov"
            GetMIMEType = "video/quicktime"
        Case ".swf"
            GetMIMEType = "application/x-shockwave-flash"
        Case ".txt"
            GetMIMEType = "text/plain"
        Case ".htm", ".html"
            GetMIMEType = "text/html"
        Case ".ram"
            GetMIMEType = "audio/x-pn-realaudio-plugin"
        Case ".doc"
            GetMIMEType = "application/msword"
        Case ".pdf"
            GetMIMEType = "application/pdf"
        Case ".rtf"
            GetMIMEType = "application/rtf"
        Case ".zip"
            GetMIMEType = "application/zip"
        Case ".xls", ".xlw", ".xla", ".xlc", ".xlm", ".xlt"
            GetMIMEType = "application/vnd.ms-excel"
        Case ".ppt", ".pps", ".pot"
            GetMIMEType = "application/vnd.ms-powerpoint"
        Case ".mdb"
            GetMIMEType = "application/x-msaccess"
        Case ".hlp"
            GetMIMEType = "application/winhlp"
        Case Else
            GetMIMEType = "application/octet-stream"
    End Select
End Function

' convert a string to a binary array
Function StringToBinary(S)
  Dim i, ByteArray
  For i=1 To Len(S)
    ByteArray = ByteArray & ChrB(Asc(Mid(S,i,1)))
  Next
  StringToBinary = ByteArray
End Function

' convert a binary array to a string
Function BinaryToString(Binary)
  Dim I, S
  For I = 1 To LenB(Binary)
    S = S & Chr(AscB(MidB(Binary, I, 1)))
  Next
  BinaryToString = S
End Function
%>