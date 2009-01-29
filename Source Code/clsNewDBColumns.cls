VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNewDBColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   File:       clsNewDbColumns.cls
'   Author:     Toby Aldridge, October 2000
'   Purpose:    Collection Class for clsNewDBColumn objects with routines to convert
'               SDD and PRD files from previous versions
'--------------------------------------------------------------------------------
'   Revisions:
'   TA 19/10/2000: Code to handle the version being in diferent columns for different types of import
'   Mo Morris 25/9/01   Changes around  expanding the clsNewDBColumn functionality
'                       to handle the dropping of columns.
'                       Unused function FieldRead removed.
'                       Functions Convert and ConvertString changed.
'                       New Function SplitCheckingForCommas added.
'--------------------------------------------------------------------------------
Option Explicit

'added to filename to create new filename
Private Const mFILE_NAME_CHANGE = "_1"

'internal collection of clsNewDbColumn objects
Private mcolNC As Collection

'STYDEF, PATRSP, or LABDEF type of import
Private msImportType As String

'--------------------------------------------------------------------------------
Public Function Convert(ByVal sFileName As String, ByRef sMessage As String) As Boolean
'--------------------------------------------------------------------------------
'Convert an Import file from a previous version to one for the current version.
'Mo Morris 25/9/01, With the additional functionality of dropping numbered fields this
'function now reads the import file a field at a time instead of a line at a time.
'--------------------------------------------------------------------------------
Dim sConvertedName As String
Dim lDotPos As Long
Dim sVersion As String
Dim bChanged As Boolean
Dim sLine As String
Dim nFile As Integer
Dim nConvertedFile As Integer
'column to find version in header line of text file ('numbering from 0)
Dim nVersionColNo As Integer

    On Error GoTo ErrHandler
    
    HourglassOn

    sMessage = ""
    bChanged = False

    'determine study type
    Select Case UCase(Right(sFileName, 3))
    Case "SDD"
        msImportType = "STYDEF"
        nVersionColNo = 5 '6th col
    Case "PRD"
        msImportType = "PATRSP"
        nVersionColNo = 6 '7th col
    Case "LDD"
        msImportType = "LABDEF"
        nVersionColNo = 5 '6th col
    End Select
    
    lDotPos = InStrRev(sFileName, ".")
    sConvertedName = Left(sFileName, lDotPos - 1) & mFILE_NAME_CHANGE & Mid(sFileName, lDotPos)
    
    nFile = FreeFile
    Open sFileName For Input As nFile
    Line Input #nFile, sLine
    
    'get version number
    sVersion = Replace(Split(sLine, ",")(nVersionColNo), """", "")
    
    'call Load which places into mcolNC only the database column changes that have happened
    'since the version in which the import file was created.
    Load Val(Split(sVersion, ".")(0)), Val(Split(sVersion, ".")(1)), Val(Split(sVersion, ".")(2))
            
    If mcolNC.Count > 0 Then
        'there are potential changes to be made
        nConvertedFile = FreeFile
        Open sConvertedName For Output As #nConvertedFile
        'print header line
        Print #nConvertedFile, sLine
        Do While Not EOF(nFile)
            Line Input #nFile, sLine
            'convert the string and update bChanged to true if changed
            bChanged = bChanged Or ConvertString(sLine)
            'print new line
            Print #nConvertedFile, sLine
        Loop
        Close #nFile
        Close #nConvertedFile
    Else
        'no changes between these two version
        Close #nFile
        bChanged = False
        
    End If
    
    If bChanged Then
        'kill old import file
        On Error Resume Next
        Kill sFileName
        On Error GoTo ErrHandler
        
        'rename file
        'copy the converted file to a file with the original name
        FileCopy sConvertedName, sFileName
        'kill Converted name file
        On Error Resume Next
        Kill sConvertedName
        On Error GoTo ErrHandler
        sMessage = "The import file " & sFileName & " was converted from version " & sVersion & "during import"
    Else
        'hasn't changed
        sMessage = "The import file " & sFileName & " was not changed"
    End If
    
    HourglassOff
    
    Convert = bChanged
    
Exit Function

ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "Convert", "clsNewDbColumns")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select


End Function

'--------------------------------------------------------------------------------
Private Sub Load(lMajor As Long, lMinor As Long, lRevision As Long)
'--------------------------------------------------------------------------------
'load the collection with all added columns since version passed in
'--------------------------------------------------------------------------------
Dim rsNC As ADODB.Recordset
Dim lImportVersion As Long
Dim lVersion As Long
Dim sSQL As String
Dim oNewColumn As clsNewDBColumn

    On Error GoTo ErrHandler

    lImportVersion = VersionNumber(lMajor, lMinor, lRevision)
    
    Set mcolNC = New Collection
    
    sSQL = "SELECT NewDbColumn.*, SegmentId FROM NewDBColumn, MACROTable WHERE "
    If goUser.Database.DatabaseType = MACRODatabaseType.oracle80 Then
        sSQL = sSQL & "NLS_UPPER(NewDBColumn.TableName) = NLS_UPPER(MACROTable.TableName) "
    Else
        sSQL = sSQL & "NewDBColumn.TableName = MACROTable.TableName"
    End If
    'the import type is the column to filter on
    sSQL = sSQL & " AND " & msImportType & "=1"
    sSQL = sSQL & " ORDER BY VersionMajor, VersionMinor, VersionRevision, ColumnOrder"
    Set rsNC = New ADODB.Recordset
    rsNC.Open sSQL, MacroADODBConnection
    Do While Not rsNC.EOF
        lVersion = VersionNumber(rsNC!VersionMajor, rsNC!VersionMinor, rsNC!VersionRevision)
        If lVersion > lImportVersion Then
            'record is a database change since lImportVersion
            Set oNewColumn = New clsNewDBColumn
            With oNewColumn
                .Version = lVersion
                .TableName = rsNC!TableName
                .ColumnName = rsNC!ColumnName
                .ColumnOrder = rsNC!ColumnOrder
                .DefaultValue = rsNC!DefaultValue
                .Segment = Val(rsNC!SegmentId)
                'Mo Morris 25/9/01, new class elements ChangeType and ColumnNumber added
                .ChangeType = rsNC!ChangeType
                .ColumnNumber = rsNC!ColumnNumber
                mcolNC.Add oNewColumn, GetKey(.Version, .TableName, .ColumnName)
            End With
        End If
        rsNC.MoveNext
    Loop
    
    Set oNewColumn = Nothing
    rsNC.Close
    Set rsNC = Nothing
    
Exit Sub

ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "Load", "clsNewDBColumns")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

'--------------------------------------------------------------------------------
Private Function GetKey(lVersion As Long, sTableName As String, sColumnName As String) As String
'--------------------------------------------------------------------------------
'genreate unique key for a NewDbColumn object
'--------------------------------------------------------------------------------

    GetKey = "K" & Format(lVersion) & "|" & sTableName & "|" & sColumnName

End Function

'--------------------------------------------------------------------------------
Private Function VersionNumber(lMajor As Long, lMinor As Long, lRevison As Long) As Long
'--------------------------------------------------------------------------------
'create a single version number from the Major, Minor and Revision numbers
'--------------------------------------------------------------------------------

    VersionNumber = Val(Format(lMajor, "00") & Format(lMinor, "00") & Format(lRevison, "0000"))

End Function

'--------------------------------------------------------------------------------
Private Function ConvertString(sLine As String) As Boolean
'--------------------------------------------------------------------------------
' convert an SDD,PRD file line from a previous version into one forthe current version
'--------------------------------------------------------------------------------
'Mo Morris 25/9/01, This sub used to handle NEWCOLUMN activities on a single line
'   in an import file. It now handles DROPColumn activities.
'--------------------------------------------------------------------------------
Dim oNewColumn As clsNewDBColumn
Dim lSegment As Long
Dim lPos As Long
Dim bChanged As Boolean
Dim asFields() As String
Dim i As Integer
    
    On Error GoTo ErrHandler
    
    'get segment - always the 2nd, 3rd and 4th characters on a line
    lSegment = Val(Mid(sLine, 2, 3))
    
    For Each oNewColumn In mcolNC
        'check for changes in order
        If lSegment = oNewColumn.Segment Then
            'this line needs changing
            bChanged = True
            'Check ChangeType for "NEWCOLUMN" or "DROPCOLUMN"
            If oNewColumn.ChangeType = "NEWCOLUMN" Then
                lPos = InStrRev(sLine, ",")
                'insert the contents of the NewColumn's DefaultValue as the last field but one
                'The last field needs to remain the line counter
                sLine = Left(sLine, lPos) & oNewColumn.DefaultValue & "," & Mid(sLine, lPos + 1)
            Else
                'Its a "DROPCOLUMN". Remove column field number oNewColumn.ColumnNumber
                'Note that the format of a line is
                'SegmentId,Field1,Field2,Field3,Field4, ... Field26,LineCounter
                'place the contents of the comma separated fields into array asFields
                'note that Split cannot be used because the text within some of the fields might contain commas
                asFields = SplitCheckingForCommas(sLine)
                sLine = ""
                For i = LBound(asFields) To UBound(asFields)
                    'copy all fields back into sLine apart from field to be dropped
                    If i <> oNewColumn.ColumnNumber Then
                        sLine = sLine & asFields(i)
                        If i <> UBound(asFields) Then
                            sLine = sLine & ","
                        End If
                    End If
                Next i
            End If
        End If
    Next
    
    ConvertString = bChanged

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ConvertString", "clsNewDBColumns")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select

End Function

'--------------------------------------------------------------------------------
Public Function Count() As Long
'--------------------------------------------------------------------------------
'return the count of the internal collection
'--------------------------------------------------------------------------------

    Count = mcolNC.Count

End Function

''--------------------------------------------------------------------------------
'Public Function FieldRead(lField As Long, sFieldString As String, sSep As String, Optional sDelim As String = "") As String
''--------------------------------------------------------------------------------
''TA 11/10/2000: return a string from a field string specifying separator and delimiter
''--------------------------------------------------------------------------------
'Dim lFields As Long
'Dim lStart As Long
'Dim lEnd As Long
'Dim i As Long
'
'    If lField <> 0 Then
'        If sDelim = "" Then
'            FieldRead = Split(sFieldString, sSep)(lField - 1)
'        Else
'            'Debug.Print Len(sFieldString)
'            lStart = 1
'            lEnd = InStr(2, sFieldString, sDelim)
'            For i = 2 To lField
'                lStart = lEnd + 2
'                lEnd = InStr(lStart + 1, sFieldString, sDelim)
'                If lEnd >= Len(sFieldString) Or lEnd = 0 Then
'                    'end of string - lField th field doesn not exist
'                    FieldRead = ""
'                    Exit Function
'                End If
'                'Debug.Print i, Mid(sFieldString, lStart, lEnd - lStart), lStart, lEnd
'            Next
'            FieldRead = Mid(sFieldString, lStart + 1, lEnd - lStart - 1)
'        End If
'    End If
'
'End Function

'--------------------------------------------------------------------------------
Private Sub Class_Terminate()
'--------------------------------------------------------------------------------

    Set mcolNC = Nothing

End Sub

'--------------------------------------------------------------------------------
Private Function SplitCheckingForCommas(ByVal sLine As String) As String()
'--------------------------------------------------------------------------------
Dim asTemp() As String
Dim nFieldNumber As Integer
Dim sField As String
Dim sChar As String
Dim i As Integer
Dim bNotWithinQuotes As Boolean

    'To make this function work correctly a comma is added at the end of the input string
    sLine = sLine & ","

    'initialise the boolean that controls if a comma occurs between double quotes
    bNotWithinQuotes = True
    'initialise the field number as used to index the array
    nFieldNumber = 0
    'initialise the variable into which a field is built up
    sField = ""
    
    For i = 1 To Len(sLine)
        sChar = Mid(sLine, i, 1)
        If sChar = """" Then
            bNotWithinQuotes = Not bNotWithinQuotes
        End If
        If sChar = "," And bNotWithinQuotes Then
            'place comma separated field into the array of strings
            ReDim Preserve asTemp(nFieldNumber)
            asTemp(nFieldNumber) = sField
            nFieldNumber = nFieldNumber + 1
            'initialise sField ready for the next field
            sField = ""
        Else
            'add current character to current field
            sField = sField & sChar
        End If
    Next i
        
    SplitCheckingForCommas = asTemp

End Function
