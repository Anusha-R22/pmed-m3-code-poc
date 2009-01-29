Attribute VB_Name = "libXCeed"
'----------------------------------------------------------------------------------------'
'   File:       libXCeed.bas
'   Copyright:  InferMed Ltd. 2001. All Rights Reserved
'   Author:     Toby Aldridge, June 2001
'   Purpose:    General XCeed functions
'----------------------------------------------------------------------------------------'
' Revisions:
' DPH 10/04/2002 - ZipFiles function changed to accept file array
' MLM 22/04/2002: Moved hex en/decoding functions to this module from
'                 basCommon, since they rely on functions from here.
' DPH 24/04/2002 - Delete output file if exists in HEXEncodeFileXCeed
'                   Added licence information for runtime use
'----------------------------------------------------------------------------------------'

Option Explicit

' Store the licence keys that are required for runtime dlls
Public Const msEXCEED_ZIP_LICENCE = "SFX45-XHUMC-NTHKJ-G45A"
Public Const msEXCEED_BIN_ENCODE_LICENCE = "BEN10-5RUZC-TTYTN-AAXA"
Public Const msEXCEED_ENCRYPTION_LICENCE = "CRY10-ARUXC-ATY7N-B4JA"

Private Const sSECRETKEY = "Zorba the Greek"

'--------------------------------------------------------------------
Public Sub ZipFiles(sFileName() As String, sZipFileName As String)
'--------------------------------------------------------------------
'REM 05/04/02
'Compresses files into .zip format
' DPH 10/04/2002 - Takes an array of filenames rather than one
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
    Err.Raise Err.Number, , "XCeed Error Number " & ResultCode & "|" & "libXCeed.ZipFiles"

End Sub

'--------------------------------------------------------------------
Public Function ZipFilesToEncodedString(sZipFileName As String) As String
'--------------------------------------------------------------------
'REM 05/04/02
'Returns a compresses encoded string from a zipped file
'--------------------------------------------------------------------
Dim bCompData() As Byte
Dim vEncoded As Variant
Dim sCompEncoded As String

    On Error GoTo Errorlabel

    'convert zipped file into byte array
    bCompData = ByteArrayFromFile(sZipFileName)
    
    'Encode byte array
    vEncoded = EncodeVariant(bCompData)
    
    'Converts the compressed encoded variant to a string
    sCompEncoded = StringFromXCeedArray(vEncoded)

    'Return string
    ZipFilesToEncodedString = sCompEncoded
    
Exit Function
Errorlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "libXCeed.ZipFilesToEncodedString"
End Function

'--------------------------------------------------------------------
Public Sub DecodeAndWriteStringToFile(sEncodedString As String, sZipFileName As String)
'--------------------------------------------------------------------
'REM 05/04/02
'Routine decodes a string and writes the it to a .zip file
'--------------------------------------------------------------------
Dim xBinEncode As New XceedBinaryEncoding
Dim baEncoded() As Byte

    On Error GoTo Errorlabel

    ' required for runtime
    xBinEncode.License msEXCEED_BIN_ENCODE_LICENCE
    
    'set the decoding format
    Set xBinEncode.EncodingFormat = New XceedHexaEncodingFormat
    
    baEncoded = StrConv(sEncodedString, vbFromUnicode)
    
    'Decode compressed string and write it to a file (with file extension .zip)
    Call xBinEncode.WriteFile(baEncoded, bfpDecode, True, sZipFileName, False)
    
Exit Sub
Errorlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "libXCeed.DecodeAndWriteStringToFile"
End Sub

'--------------------------------------------------------------------
Public Sub UnZipFiles(sUnZipToFolder As String, sZipFileName As String)
'--------------------------------------------------------------------
'REM 05/04/02
'Unzip .zip files into a predefined folder
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
    Err.Raise Err.Number, , "XCeed Error Number " & ResultCode & "|" & "libXCeed.UnZipFiles"
End Sub

'--------------------------------------------------------------------
Private Function EncodeVariant(vCompData As Variant) As Variant
'--------------------------------------------------------------------
'REM 05/04/02
'Hexadecimalises a compressed string
' DPH 24/04/2002 - Added EndOfLineType to avois checksum problems
'--------------------------------------------------------------------
Dim xBinEncode As New XceedBinaryEncoding
Dim b() As Byte
Dim vEncoded As Variant

    On Error GoTo Errorlabel

    ' required for runtime
    xBinEncode.License msEXCEED_BIN_ENCODE_LICENCE
    
    'set the encoding format
    Set xBinEncode.EncodingFormat = New XceedHexaEncodingFormat
    xBinEncode.EncodingFormat.EndOfLineType = bltNone
    
    'encode compressed string
    vEncoded = xBinEncode.Encode(vCompData, True)

    EncodeVariant = vEncoded
    
    Set xBinEncode = Nothing

Exit Function
Errorlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "libXCeed.EncodeVariant"
End Function

'----------------------------------------------------------------------------------------'
Private Function StringFromXCeedArray(vArray As Variant) As String
'----------------------------------------------------------------------------------------'
'REM 05/04/02
'Returns a string from a varient
'----------------------------------------------------------------------------------------'
Dim sText As String
Dim i As Long

    On Error GoTo Errorlabel

    'build a string of spaces to the full length we need to avoid concatenation
    sText = String$(UBound(vArray) + 1, " ")
    
    For i = 0 To UBound(vArray)
        'replace characters one by one
        Mid$(sText, i + 1) = Chr(vArray(i))
    Next

    StringFromXCeedArray = sText
    
Exit Function
Errorlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "libXCeed.StringFromXCeedArray"
End Function

'---------------------------------------------------------------------
Public Function HEXEncodeFileXCeed(ByVal vInputFile As String, _
                          ByVal vOutputFile As String) As Boolean
'---------------------------------------------------------------------
' HEX Encode using XCeed Component
'---------------------------------------------------------------------
' REVISIONS
' DPH 24/04/2002 - Delete output file if exists
'---------------------------------------------------------------------
Dim sHex As String
Dim nFile As Integer

    On Error GoTo ErrHandler
    
    sHex = ZipFilesToEncodedString(vInputFile)
    
    If Len(sHex) > 0 Then
    
        ' If file exists
        If FolderExistence(vOutputFile, True) Then
            Kill vOutputFile
        End If
        
        nFile = FreeFile
        Open vOutputFile For Binary Access Write As #nFile
        Put #nFile, , sHex
        Close #nFile
    
        HEXEncodeFileXCeed = True
    Else
        HEXEncodeFileXCeed = False
    End If
    
    Exit Function
 
ErrHandler:
    ' Function Failed
    gLog gsHEX_ENCODE, Err.Description
    HEXEncodeFileXCeed = False
End Function

'---------------------------------------------------------------------
Public Function HEXDecodeFileXCeed(ByVal vInputFile As String, _
                          ByVal vOutputFile As String) As Boolean
'---------------------------------------------------------------------
' Hex Decoding file using XCeed component
'---------------------------------------------------------------------
Dim sHex As String
Dim nFile As Integer
Dim bHexOutput As Boolean
    
    nFile = FreeFile
    Open vInputFile For Binary Access Read As #nFile
    sHex = Input(LOF(nFile), #nFile)
    Close #nFile
    
    If sHex <> "" Then
        Call DecodeAndWriteStringToFile(sHex, vOutputFile)
    Else
        HEXDecodeFileXCeed = False
        Exit Function
    End If
    
    If GetFileLength(vOutputFile) > 0 Then
        ' Set function value
        HEXDecodeFileXCeed = True
    Else
        HEXDecodeFileXCeed = False
    End If
    
    Exit Function
 
ErrHandler:
    ' Function Failed
    gLog gsHEX_DECODE, Err.Description
    HEXDecodeFileXCeed = False
End Function

'---------------------------------------------------------------------
Public Sub HEXEncodeFile(ByVal vInputFile As String, _
                          ByVal vOutputFile As String)
'---------------------------------------------------------------------
Dim mnFreeFile1 As Integer
Dim mnFreeFile2 As Integer

    On Error GoTo ErrHandler
    
    mnFreeFile1 = FreeFile
     
    Open vInputFile For Binary Access Read As #mnFreeFile1
    
    mnFreeFile2 = FreeFile
    
    Open vOutputFile For Binary Access Write As #mnFreeFile2
    
    Do While Not EOF(mnFreeFile1)
    
        Put #mnFreeFile2, , Right$("0" & Hex$(AscB(InputB(1, #mnFreeFile1))), 2)
    
    Loop
    
    Close #mnFreeFile1
    Close #mnFreeFile2
    
    Exit Sub
 
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                        "HexEncodeFile", "basCommon")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'---------------------------------------------------------------------
Public Function HEXDecodeFile(ByVal vInputFile As String, _
                          ByVal vOutputFile As String) As Boolean
'---------------------------------------------------------------------
' REVISIONS
' DPH 08/04/2002 - Turned into function & added Check to see if HexDecode Produces any output
' DPH 11/04/2002 - Check if CAB or ZIP dealing with & act appropriately
'---------------------------------------------------------------------
Dim mnFreeFile1 As Integer
Dim mnFreeFile2 As Integer
Dim msHexInput As String
Dim msHexInputTemp As String
Dim mnLong As Long
Dim msByte() As Byte
Dim bHexOutput As Boolean
    
    On Error GoTo ErrHandler

    ' DPH 11/04/2002 - Check if CAB or ZIP dealing with & act appropriately
    If UCase(Right(vInputFile, 3)) = "ZIP" Then
        HEXDecodeFile = HEXDecodeFileXCeed(vInputFile, vOutputFile)
        Exit Function
    End If
    
    mnFreeFile1 = FreeFile
     
    Open vInputFile For Binary Access Read As #mnFreeFile1
    
    mnFreeFile2 = FreeFile
    
    Open vOutputFile For Binary Access Write As #mnFreeFile2
    
    ' DPH 08/04/2002 - Set Output As False
    bHexOutput = False
    
    Do While Not EOF(mnFreeFile1)
        msHexInputTemp = Input(2, #mnFreeFile1)
        msHexInput = "&H"
        If msHexInputTemp <> vbCrLf Then
            msHexInput = msHexInput & msHexInputTemp
        End If
    
        If msHexInput <> "&H" Then
            mnLong = CLng(msHexInput)
            Put #mnFreeFile2, , CByte(mnLong)
             ' DPH 08/04/2002 - Set Output as True
            bHexOutput = True
        End If
    Loop
    
    Close #mnFreeFile1
    Close #mnFreeFile2

    ' DPH 08/04/2002 - Set function value
    HEXDecodeFile = bHexOutput
    
    Exit Function
 
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                        "HexDecodeFile", "basCommon")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
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
    Call xRijndael.SetSecretKeyFromPassPhrase(sSECRETKEY, 256)
    
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
    Err.Raise Err.Number, , Err.Description & "|" & "libXCeed.EncryptString"
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
    Call xRijndael.SetSecretKeyFromPassPhrase(sSECRETKEY, 256)
    
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
    Err.Raise Err.Number, , Err.Description & "|" & "libXCeed.DecryptString"
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
    sKey = Left(sSECRETKEY & String(64, Chr(0)), 64)
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
    Err.Raise Err.Number, , "XCeed Error Number" & "libXCeed.XceedHashing"
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

