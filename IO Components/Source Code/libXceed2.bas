Attribute VB_Name = "libXceed2"
Option Explicit
' Store the licence keys that are required for runtime dlls
Public Const msEXCEED_ZIP_LICENCE = "SFX45-XHUMC-NTHKJ-G45A"
Public Const msEXCEED_BIN_ENCODE_LICENCE = "BEN10-5RUZC-TTYTN-AAXA"
Public Const msEXCEED_ENCRYPTION_LICENCE = "CRY10-ARUXC-ATY7N-B4JA"

Private Const msSECRETKEY = "Zorba the Greek"

'Chars that need to be encoded in string to be sent via http
Public Const gsCHARS_TO_ENCODE = " []/^{}#&+,\:<>=?@~" & vbCrLf

'--------------------------------------------------------------------------------------------------------
Public Function HexEncodeChars(ByVal sString As String, ByVal sCharsToEncode As String) As String
'--------------------------------------------------------------------------------------------------------
' MLM 21/06/02: Returns sString but with any characters from sCharsToEncode replaced by a % and a
' 2-character ASCII code in hex.
' sCharsToEncode must NOT contain % or 0 - F; these would make the output undecodable.
' Example: HexEncodeChars("abc%","b") = "a%62c%25"
'--------------------------------------------------------------------------------------------------------

Dim nCount As Integer
Dim sChar As String * 1
    
    On Error GoTo ErrHandler
    
    '%s must always be encoded, and done 1st
    sCharsToEncode = "%" & sCharsToEncode
    
    For nCount = 1 To Len(sCharsToEncode)
        sChar = Mid(sCharsToEncode, nCount, 1)
        'we expect most input strings to only contain few characters to encode, so performance will be
        'improved by doing the Replace() conditionally
        If InStr(sString, sChar) > 0 Then
            sString = Replace(sString, sChar, "%" & IIf(Asc(sChar) < 16, "0", "") & Hex(Asc(sChar)))
        End If
    Next nCount
    
    HexEncodeChars = sString
    
    Exit Function

ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|libXCeed2.HexEncodeChars"

End Function

'--------------------------------------------------------------------------------------------------------
Public Function HexDecodeChars(ByRef sString As String) As String
'--------------------------------------------------------------------------------------------------------
' MLM 21/06/02: Do the opposite of HexEncodeChars ;)
' Don't call this with strings that aren't encoded properly, or it might fall over.
'--------------------------------------------------------------------------------------------------------

Dim lPosition As Long
    
    On Error GoTo ErrHandler
    
    HexDecodeChars = sString
    lPosition = InStrRev(sString, "%")
    Do While lPosition > 1
        HexDecodeChars = Mid(HexDecodeChars, 1, lPosition - 1) & _
            Chr(CInt("&h" & Mid(HexDecodeChars, lPosition + 1, 2))) & Mid(HexDecodeChars, lPosition + 3)
            'NB: This takes the 2 characters after the % and converts them from hex to a single character.
        lPosition = InStrRev(sString, "%", lPosition - 1)
    Loop
    If lPosition = 1 Then
        HexDecodeChars = Chr(CInt("&h" & Mid(HexDecodeChars, 2, 2))) & Mid(HexDecodeChars, 4)
    End If
    
    Exit Function

ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|libXCeed2.HexDecodeChars"
    
End Function

'---------------------------------------------------------------------
Public Function EncryptString(ByVal sString As String) As String
'---------------------------------------------------------------------
'REM 12/09/02
'Compresses and Encrypts a string
'---------------------------------------------------------------------
Dim xEncryptor As XceedEncryption
Dim xRijndael As XceedRijndaelEncryptionMethod
Dim vaDecryptedText As Variant
Dim sEncrypted As String
Dim xCompress As XceedCompression
Dim vResult As Variant
Dim vCompressedData As Variant
    

    On Error GoTo Errorlabel

    Set xEncryptor = New XceedEncryption
    xEncryptor.License msEXCEED_ENCRYPTION_LICENCE
    
    Set xCompress = New XceedCompression
    xCompress.License msEXCEED_ZIP_LICENCE
    
    Set xRijndael = New XceedRijndaelEncryptionMethod
    
    xRijndael.EncryptionMode = 0
    xRijndael.PaddingMethod = 0
    
    'set the secret key
    Call xRijndael.SetSecretKeyFromPassPhrase(msSECRETKEY, 256)
    
    'set the encryption method
    Set xEncryptor.EncryptionMethod = xRijndael
    
    Set xRijndael = Nothing

    vaDecryptedText = StrConv(sString, vbFromUnicode)
    
    'compress the data before encryption
    vResult = xCompress.Compress(vaDecryptedText, vCompressedData, True)
    
    If vResult = 0 Then
        'encrypt the string and convert it from binary to hex
        sEncrypted = BinaryToHex(xEncryptor.Encrypt(vCompressedData, True))
    Else
        'raise error
        Err.Raise vResult, "EncryptString", xCompress.GetErrorDescription(vResult)
    End If
    
    'return the encrypted string
    EncryptString = sEncrypted
    
    Set xEncryptor = Nothing
    Set xCompress = Nothing

Exit Function
Errorlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "libXCeed2.EncryptString"
End Function

'---------------------------------------------------------------------
Public Function DecryptString(ByVal sString As String) As String
'---------------------------------------------------------------------
'REM 12/09/02
'Decrypts then decompresses a string that was encrypted using the EncryptString() Function
'---------------------------------------------------------------------
Dim xEncryptor As XceedEncryption
Dim xCompress As XceedCompression
Dim xRijndael As XceedRijndaelEncryptionMethod
Dim vaDecrypted As Variant
Dim sDecrypted As String
Dim vResult As Variant
Dim vUncompressedData As Variant

    On Error GoTo Errorlabel

    Set xEncryptor = New XceedEncryption
    xEncryptor.License msEXCEED_ENCRYPTION_LICENCE
    
    Set xCompress = New XceedCompression
    xCompress.License msEXCEED_ZIP_LICENCE
    
    'set the encryption method
    Set xRijndael = New XceedRijndaelEncryptionMethod
    
    'set the encryption parameters
    xRijndael.EncryptionMode = 0
    xRijndael.PaddingMethod = 0
    
    'set the secret key
    Call xRijndael.SetSecretKeyFromPassPhrase(msSECRETKEY, 256)
    
    'set the encryption method
    Set xEncryptor.EncryptionMethod = xRijndael
    
    Set xRijndael = Nothing
    
    'decrypt the string
    vaDecrypted = xEncryptor.Decrypt(HexToBinary(sString), True)
    
    'uncompresses the decrypted data
    vResult = xCompress.Uncompress(vaDecrypted, vUncompressedData, True)
    
    If vResult = 0 Then
        'convert the decrypted variant into a string
        sDecrypted = StrConv(vUncompressedData, vbUnicode)
    Else
        'raise error
        Err.Raise vResult, "EncryptString", xCompress.GetErrorDescription(vResult)
    End If
    
    'return the decrypted string
    DecryptString = sDecrypted
    
    Set xEncryptor = Nothing
    Set xCompress = Nothing
    
Exit Function
Errorlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "libXCeed2.DecryptString"
End Function

'---------------------------------------------------------------------
Public Function HashHexEncodeString(ByVal sPasswordTimeStamp As String) As String
'---------------------------------------------------------------------
'REM 13/09/02
'Function takes a string inserts a secret key then Xors and hashes the string twice
'and then returns a Hexed string
'---------------------------------------------------------------------
Dim xHasher As XceedHashing
Dim xSHA As XceedSHAHashingMethod
Dim vaMessageToHash1 As Variant
Dim vaMessageToHash2 As Variant
Dim sHashValue1 As String
Dim sHashFinal As String
Dim i As Long
Dim sKey As String
Dim sIPad As String
Dim sOPad As String
Dim sResult1 As String
Dim sResult2 As String
Dim sConcat1 As String
Dim sConcat2 As String

    On Error GoTo Errorlabel

    'set the Hashing object
    Set xHasher = New XceedHashing
    
    'set the mashing method
    Set xSHA = New XceedSHAHashingMethod
    
    'licence the component
    xHasher.License msEXCEED_ENCRYPTION_LICENCE
    
    'set the hash size
    xSHA.HashSize = 256
    
    Set xHasher.HashingMethod = xSHA
    
    Set xSHA = Nothing
    
    'create a secret key
    sKey = Left(msSECRETKEY & String(64, Chr(0)), 64)
    sIPad = String(64, Chr(54))
    sOPad = String(64, Chr(94))

    sResult1 = ""
    'create result1 by Xoring the key with a string of characters
    For i = 1 To Len(sKey)
        sResult1 = sResult1 & Chr(Asc(Mid(sKey, i, 1)) Xor Asc(Mid(sIPad, i, 1)))
    Next

    'Concatenate with password and timestamp
    sConcat1 = sResult1 & sPasswordTimeStamp
    
    'convert the concatanated string into a varient
    vaMessageToHash1 = StrConv(sConcat1, vbFromUnicode)
    
    'hash the first concatanated string
    Call xHasher.Hash(vaMessageToHash1, True)
    
    'convert the result into a hexed string
    sHashValue1 = BinaryToHex(xHasher.HashingMethod.HashValue)
    
    sResult2 = ""

    'create result2 string
    For i = 1 To Len(sKey)
        sResult2 = sResult2 & Chr(Asc(Mid(sKey, i, 1)) Xor Asc(Mid(sOPad, i, 1)))
    Next

    'Concatenate the result2 string with the hash string
    sConcat2 = sResult2 & sHashValue1
    
    'convert the concatanated string to a varient
    vaMessageToHash2 = StrConv(sConcat2, vbFromUnicode)
    
    'hash the seconded concatanated string
    Call xHasher.Hash(vaMessageToHash2, True)
    
    'convert the result into a hexed string
    sHashFinal = BinaryToHex(xHasher.HashingMethod.HashValue)
    
    'return the hexed string
    HashHexEncodeString = sHashFinal

    Set xHasher = Nothing

Exit Function
Errorlabel:
    Err.Raise Err.Number, , "XCeed Error Number" & "libXCeed2.XceedHashing"
End Function

'---------------------------------------------------------------------
Private Function BinaryToHex(ByRef vaBinaryValue As Variant) As String
'---------------------------------------------------------------------
'REM 12/09/02
'Convert binary to hex
'---------------------------------------------------------------------
Dim i As Long
Dim sNewValue As String
    
    sNewValue = ""
    
    If (VarType(vaBinaryValue) And vbArray) = vbArray Then
        For i = 0 To LenB(vaBinaryValue) - 1
            sNewValue = sNewValue & Right("0" & Hex(CLng(vaBinaryValue(i))), 2)
        Next i
    End If
    
    BinaryToHex = sNewValue

End Function

'---------------------------------------------------------------------
Private Function HexToBinary(ByRef sHexValue As String) As Variant
'---------------------------------------------------------------------
'REM 12/09/02
'Convert the hexadecimal string to a byte array, return as a variant
'---------------------------------------------------------------------
Dim i As Long
Dim NewValue() As Byte
    
    i = 0
    While i < Len(sHexValue)
        ReDim Preserve NewValue(i / 2)
        NewValue(i / 2) = Val("&H" & Mid$(sHexValue, i + 1, 2))
        i = i + 2
    Wend
    
    HexToBinary = NewValue

End Function

Public Function HexToBinaryString(sHexValue As String) As String
Dim b() As Byte

    b = HexToBinary(sHexValue)
    HexToBinaryString = b
    
End Function

Public Function BinaryStringToHex(sBinaryValue As String) As String
Dim b() As Byte

    b = sBinaryValue
    BinaryStringToHex = BinaryToHex(b)
    
End Function

'--------------------------------------------------------------------
Public Sub UnZipFiles(sUnZipToFolder As String, sZipFileName As String)
'--------------------------------------------------------------------
'REM 05/04/02
'Unzip .zip files into a predefined folder
' DPH 24/01/2003 - Added to libXceed2 for use in DLL
'--------------------------------------------------------------------
Dim oZip As XceedZip
Dim ResultCode As xcdError

    On Error GoTo Errorlabel

    Set oZip = New XceedZip
    
    ' Required for runtime
    oZip.License msEXCEED_ZIP_LICENCE
    
    oZip.FilesToProcess = "" ' The file to unzip, if left blank will unzip all files in the zip file
    oZip.PreservePaths = False
    
    oZip.UnzipToFolder = sUnZipToFolder
    oZip.ZipFilename = sZipFileName

    'Unzip
    ResultCode = oZip.Unzip
    
    If ResultCode <> 0 Then
        GoTo Errorlabel
    End If

    Set oZip = Nothing
    
Exit Sub
Errorlabel:
    Err.Raise Err.Number, , "XCeed Error Number " & ResultCode & "|" & "libXCeed2.UnZipFiles"
End Sub

'--------------------------------------------------------------------
Public Sub ZipFiles(sFileName() As String, sZipFileName As String)
'--------------------------------------------------------------------
'REM 05/04/02
'Compresses files into .zip format
' DPH 10/04/2002 - Takes an array of filenames rather than one
' DPH 24/01/2003 - Added to libXceed2 for use in DLL
'--------------------------------------------------------------------
Dim oZip As XceedZip
Dim ResultCode As xcdError
Dim i As Integer
Dim nFiles As Integer

    On Error GoTo Errorlabel

    Set oZip = New XceedZip
    
    ' Required for runtime
    oZip.License msEXCEED_ZIP_LICENCE
    
    ' The files to be compressed, sFileName can be set as the path to a specific file or all files in the folder
    ' by using eg. "c:\Data\*"
    
    ' Add files 'to process' to zip
    nFiles = UBound(sFileName) + 1
    
    For i = 0 To nFiles - 1
        oZip.AddFilesToProcess sFileName(i)
    Next
    
'    oZip.FilesToProcess = sFileName

    'The name and path of the compressed file to be created, must have a .zip file extension
    oZip.ZipFilename = sZipFileName
    
    'Zip file
    ResultCode = oZip.Zip
    
    If ResultCode <> 0 Then
        GoTo Errorlabel
    End If
    
    Set oZip = Nothing
    
Exit Sub
Errorlabel:
    Err.Raise Err.Number, , "XCeed Error Number " & ResultCode & "|" & "libXCeed2.ZipFiles"

End Sub


''---------------------------------------------------------------------
'Public Function BinHexString(sSource As String) As String
''---------------------------------------------------------------------
'Dim xBinEncode As New XceedBinaryEncoding
'Dim vEncoded As Variant
'Dim sEncoded As String
'Dim sHex As String
'Dim stext As String
'Dim vDecoded As Variant
'Dim DecodedDataArray() As Byte
'
'
'    'Set xBinEncode.EncodingFormat = New XceedBase64EncodingFormat
'    Set xBinEncode.EncodingFormat = New XceedHexaEncodingFormat
'    xBinEncode.EncodingFormat.EndOfLineType = bltNone
'
'    On Error Resume Next
'    vEncoded = xBinEncode.Encode(sSource, True)
'    DecodedDataArray = vEncoded
'    sHex = StringFromXCeedArray(vEncoded)
'
'    vDecoded = xBinEncode.Decode(vEncoded, True)
'
'    stext = StringFromXCeedArray(vDecoded)
'
'    Set xBinEncode = Nothing
'
'End Function
'
''----------------------------------------------------------------------------------------'
'Private Function StringFromXCeedArray(vArray As Variant) As String
''----------------------------------------------------------------------------------------'
''REM 05/04/02
''Returns a string from a varient
''----------------------------------------------------------------------------------------'
'Dim stext As String
'Dim i As Long
'
'    On Error GoTo Errorlabel
'
'    'build a string of spaces to the full length we need to avoid concatenation
'    stext = String$(UBound(vArray) + 1, " ")
'
'    For i = 0 To UBound(vArray)
'        'replace characters one by one
'        Mid$(stext, i + 1) = Chr(vArray(i))
'    Next
'
'    StringFromXCeedArray = stext
'
'Exit Function
'Errorlabel:
'    Err.Raise Err.Number, , Err.Description & "|" & "libXCeed2.StringFromXCeedArray"
'End Function
'
