<%
	' get filename
	sFileName=Request.QueryString("filename")
	' set up document path
	sDocumentPath =  server.mappath("../../documents/") & "\" '& rsResult1("documentpath")
	
	' work out content-type
	sContentType = GetMIMEType(sFileName)
	
	' Forces users to save a file to a location on their hard drive
    ' default as attachment (will open if can else will offer download dialog)
    Response.AddHeader "content-disposition", "attachment; filename=" & sFileName
	' set content-type
    Response.ContentType = sContentType 

    ' stream to browser using ADODB Stream object
    Set oStream = server.CreateObject("ADODB.Stream")
    oStream.Open
	oStream.Type = 1 ' 1 = adTypeBinary
    oStream.LoadFromFile sDocumentPath & sFileName 
    Response.AddHeader "Content-Length", CStr(oStream.Size)
    Response.BinaryWrite oStream.Read
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
%>