cscript.exe adsutil.vbs set W3Svc/Filters/Compression/GZIP/HcFileExtensions "htm" "html" "js" "css"
cscript.exe adsutil.vbs set W3Svc/Filters/Compression/DEFLATE/HcFileExtensions "htm" "html" "js" "css"
cscript.exe adsutil.vbs set W3Svc/Filters/Compression/GZIP/HcScriptFileExtensions "asp"
cscript.exe adsutil.vbs set W3Svc/Filters/Compression/DEFLATE/HcScriptFileExtensions "asp" 
IISreset.exe /restart