Attribute VB_Name = "modBlobs"
'----------------------------------------------------------------------------------------'
'   File:       modBlobs.bas
'   Copyright:  InferMed Ltd. 1999. All Rights Reserved
'   Author:
'   Purpose:
'----------------------------------------------------------------------------------------'
' Revisions:
'   SDM 10/11/99
'       Copied the error handler routines in.
' TA 04/10/2001 -  adLongVarWChar checked when adLongVarChar checked
'----------------------------------------------------------------------------------------'
Option Explicit
Const BLOCK_SIZE = 16384

'---------------------------------------------------------------------
Sub BlobToFile(fld As ADODB.Field, ByVal FName As String, _
                    Optional FieldSize As Long = -1, _
                    Optional Threshold As Long = 1048576)
'---------------------------------------------------------------------
' Assumes file does not exist
' Data cannot exceed approx. 2Gb in size
'---------------------------------------------------------------------
Dim F As Long
Dim bData() As Byte
Dim sData As String
    On Error GoTo ErrHandler
    
    F = FreeFile
    Open FName For Binary As #F
    Select Case fld.Type
        Case adLongVarBinary
            If FieldSize = -1 Then   ' blob field is of unknown size
                WriteFromUnsizedBinary F, fld
            Else                     ' blob field is of known size
                If FieldSize > Threshold Then   ' very large actual data
                    WriteFromBinary F, fld, FieldSize
                Else                            ' smallish actual data
                    bData = fld.Value
                    Put #F, , bData  ' PUT tacks on overhead if use fld.Value
                End If
            End If
        Case adLongVarChar, adLongVarWChar ' TA 04/10/2001 -  adLongVarWChar checked when adLongVarChar checked
            If FieldSize = -1 Then
                WriteFromUnsizedText F, fld
            Else
                If FieldSize > Threshold Then
                    WriteFromText F, fld, FieldSize
                Else
                    sData = fld.Value
                    Put #F, , sData  ' PUT tacks on overhead if use fld.Value
                End If
            End If
    End Select
    Close #F

Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "BlobToFile", "modBlobs.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Sub
     
'---------------------------------------------------------------------
Sub WriteFromBinary(ByVal F As Long, fld As ADODB.Field, _
                    ByVal FieldSize As Long)
'---------------------------------------------------------------------
Dim Data() As Byte
Dim BytesRead As Long
    On Error GoTo ErrHandler
    
    Do While FieldSize <> BytesRead
        If FieldSize - BytesRead < BLOCK_SIZE Then
            Data = fld.GetChunk(FieldSize - BLOCK_SIZE)
            BytesRead = FieldSize
        Else
            Data = fld.GetChunk(BLOCK_SIZE)
            BytesRead = BytesRead + BLOCK_SIZE
        End If
        Put #F, , Data
    Loop

Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "WriteFromBinary", "modBlobs.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Sub
      
'---------------------------------------------------------------------
Sub WriteFromUnsizedBinary(ByVal F As Long, fld As ADODB.Field)
'---------------------------------------------------------------------
Dim Data() As Byte
Dim Temp As Variant
    On Error GoTo ErrHandler
    
    Do
        Temp = fld.GetChunk(BLOCK_SIZE)
        If IsNull(Temp) Then Exit Do
        Data = Temp
        Put #F, , Data
    Loop While LenB(Temp) = BLOCK_SIZE

Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "WriteFromUnsizedBinary", "modBlobs.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Sub
      
'---------------------------------------------------------------------
Sub WriteFromText(ByVal F As Long, fld As ADODB.Field, _
                    ByVal FieldSize As Long)
'---------------------------------------------------------------------
Dim Data As String
Dim CharsRead As Long
    On Error GoTo ErrHandler
    
    Do While FieldSize <> CharsRead
        If FieldSize - CharsRead < BLOCK_SIZE Then
            Data = fld.GetChunk(FieldSize - BLOCK_SIZE)
            CharsRead = FieldSize
        Else
            Data = fld.GetChunk(BLOCK_SIZE)
            CharsRead = CharsRead + BLOCK_SIZE
        End If
        Put #F, , Data
    Loop

Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "WriteFromText", "modBlobs.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Sub
          
'---------------------------------------------------------------------
Sub WriteFromUnsizedText(ByVal F As Long, fld As ADODB.Field)
'---------------------------------------------------------------------
Dim Data As String
Dim Temp As Variant
    On Error GoTo ErrHandler
    
    Do
        Temp = fld.GetChunk(BLOCK_SIZE)
        If IsNull(Temp) Then Exit Do
        Data = Temp
        Put #F, , Data
    Loop While Len(Temp) = BLOCK_SIZE

Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "WriteFromUnsizedText", "modBlobs.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Sub
       
'---------------------------------------------------------------------
Sub FileToBlob(ByVal FName As String, fld As ADODB.Field, _
                Optional Threshold As Long = 1048576)
'---------------------------------------------------------------------
' Assumes file exists
' Assumes calling routine does the UPDATE
' File cannot exceed approx. 2Gb in size
'---------------------------------------------------------------------
Dim F As Long
Dim Data() As Byte
Dim FileSize As Long
    On Error GoTo ErrHandler
    
    F = FreeFile
    Open FName For Binary As #F
    FileSize = LOF(F)
    Select Case fld.Type
        Case adLongVarBinary
            If FileSize > Threshold Then
                ReadToBinary F, fld, FileSize
            Else
                Data = InputB(FileSize, F)
                fld.Value = Data
            End If
        Case adLongVarChar, adLongVarWChar ' TA 04/10/2001 -  adLongVarWChar checked when adLongVarChar checked
            If FileSize > Threshold Then
                ReadToText F, fld, FileSize
            Else
                fld.Value = Input(FileSize, F)
            End If
        End Select
        Close #F

Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "FileToBlob", "modBlobs.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Sub
        
'---------------------------------------------------------------------
Sub ReadToBinary(ByVal F As Long, fld As ADODB.Field, _
                ByVal FileSize As Long)
'---------------------------------------------------------------------
Dim Data() As Byte
Dim BytesRead As Long
    On Error GoTo ErrHandler
    
    Do While FileSize <> BytesRead
        If FileSize - BytesRead < BLOCK_SIZE Then
            Data = InputB(FileSize - BytesRead, F)
            BytesRead = FileSize
        Else
            Data = InputB(BLOCK_SIZE, F)
            BytesRead = BytesRead + BLOCK_SIZE
        End If
        fld.AppendChunk Data
    Loop

Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ReadToBinary", "modBlobs.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Sub
          
'---------------------------------------------------------------------
Sub ReadToText(ByVal F As Long, fld As ADODB.Field, _
                ByVal FileSize As Long)
'---------------------------------------------------------------------
Dim Data As String
Dim CharsRead As Long
    On Error GoTo ErrHandler
    
    Do While FileSize <> CharsRead
        If FileSize - CharsRead < BLOCK_SIZE Then
            Data = Input(FileSize - CharsRead, F)
            CharsRead = FileSize
        Else
            Data = Input(BLOCK_SIZE, F)
            CharsRead = CharsRead + BLOCK_SIZE
        End If
        fld.AppendChunk Data
    Loop

Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ReadToText", "modBlobs.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Sub


'---------------------------------------------------------------------
Public Function BlobToVariant(fld As ADODB.Field, _
                    Optional FieldSize As Long = -1, _
                    Optional Threshold As Long = 1048576) As Variant
'---------------------------------------------------------------------
Dim bData() As Byte
Dim vData As Variant
Dim VarName As Variant
    On Error GoTo ErrHandler
    Select Case fld.Type
        Case adLongVarBinary
            'Binary data not tested yet
            If FieldSize > Threshold Then   ' very large actual data
                'WriteFromBinary F, fld, FieldSize
            Else                            ' smallish actual data
                bData = fld.Value
                'Put #F, , bData  ' PUT tacks on overhead if use fld.Value
            End If
        Case adLongVarChar, adLongVarWChar ' TA 04/10/2001 -  adLongVarWChar checked when adLongVarChar checked
            If FieldSize > Threshold Then
                vData = WriteFromTextVar(fld, FieldSize)
            Else
                vData = fld.Value
            End If
    End Select
    BlobToVariant = vData

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "BlobToVariant", "modBlobs.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
    
End Function

'---------------------------------------------------------------------
Private Function WriteFromTextVar(fld As ADODB.Field, _
                    ByVal FieldSize As Long) As Variant
'---------------------------------------------------------------------
Dim sData As String
Dim lCharsRead As Long
Dim vData As Variant
    On Error GoTo ErrHandler
    
    Do While FieldSize <> lCharsRead
        If FieldSize - lCharsRead < BLOCK_SIZE Then
            sData = fld.GetChunk(FieldSize - BLOCK_SIZE)
            lCharsRead = FieldSize
        Else
            sData = fld.GetChunk(BLOCK_SIZE)
            lCharsRead = lCharsRead + BLOCK_SIZE
        End If
        vData = vData & sData
    Loop
    WriteFromTextVar = vData

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "WriteFromTextVar", "modBlobs.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Function


'---------------------------------------------------------------------
Public Function ReadFromTextVar(fld As ADODB.Field, _
                                 ByVal vData As Variant)
'---------------------------------------------------------------------
Dim sData As String
Dim lCharsRead As Long
Dim oData As Variant

    On Error GoTo ErrHandler
    
    lCharsRead = 0
    
    Do While Len(vData) <> lCharsRead
        If Len(vData) - lCharsRead < BLOCK_SIZE Then
            fld.AppendChunk vData
        Else
            oData = Mid(vData, lCharsRead + 1, BLOCK_SIZE)
            fld.AppendChunk oData
            lCharsRead = lCharsRead + BLOCK_SIZE
        End If
    Loop

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ReadFromTextVar", "modBlobs.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Function



