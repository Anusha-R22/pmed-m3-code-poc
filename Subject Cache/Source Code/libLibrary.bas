Attribute VB_Name = "libLibrary"
'----------------------------------------------------------------------------------------'
'   File:       libLibrary.bas
'   Copyright:  InferMed Ltd. 2001. All Rights Reserved
'   Author:     Toby Aldridge, April 2001
'   Purpose:    General Library Functions
'----------------------------------------------------------------------------------------'
' Revisions:
'   NCJ 22 May 01 - Need object keys in CollectionDeSerialise
'   NCJ 25 May 01 - CollectionMember and CollectionAdd
'   TA 5/7/02 - CollectionToArray added
'   NCJ 2 Oct 02 - Added IMedNow function to provide an accurate timestamp
'   TA 10/10/2002: Added double quotes constant and function
'   NCJ 31 Oct 02 - Added CollectionRemoveAnyway
'   ic 05/11/2002   moved FileExists() from basCommon to libLibrary
'   NCJ 12 May 03 - Added new optional params to LocalNumToStandard and StandardNumToLocal
'   ic 23/03/2004 added function CollectionRemoveAll() to fix memory leak problems
'----------------------------------------------------------------------------------------'

Option Explicit

Public Const NULL_INTEGER = -32768
Public Const NULL_LONG = -2147483648#
Public Const NULL_DOUBLE = -2147483648#
Public Const NULL_STRING = ""
Public Const NULL_DATE = 0

Public Const DATATYPE_DATE = -1

' following two property bag utils constants are as short as possible to save space
Public Const PB_ITEM = "I"
Public Const PB_COUNT = "C"

' "Standard" dot and comma
Private Const msSTANDARD_DOT = "."
Private Const msSTANDARD_COMMA = ","

Public Const DOUBLE_QUOTE = """"

'enum for ServerExists Function
Public Enum eServerExistsResult
    serYes = 0
    serNo = 1
    serMaybe = 2
End Enum

'---------------------------------------------------------------------
Public Sub CollectionRemoveAll(ByRef colCollection As Collection)
'---------------------------------------------------------------------
'   ic/dph 23/03/2004
'   function removes all items from a passed collection
'---------------------------------------------------------------------
Dim n As Integer
    
    If (Not colCollection Is Nothing) Then
        For n = colCollection.Count To 1 Step -1
            Call colCollection.Remove(n)
        Next
    End If
End Sub

'TA 18/9/01 Uncomment this once we have can remove other version from basCommon
''----------------------------------------------------------------------
'Public Function RemoveNull(vVariable As Variant) As String
''----------------------------------------------------------------------
'
'    RemoveNull = ConvertFromNull(vVariable, vbString)
'
'End Function

'---------------------------------------------------------------------
Function FileExists(ByVal strPathName As String) As Integer
'---------------------------------------------------------------------
' FUNCTION: FileExists
' Determines whether the specified file exists
'
' IN: [strPathName] - file to check for
'
' Returns: True if file exists, False otherwise
'---------------------------------------------------------------------
Dim intFileNum As Integer

    On Error Resume Next

    '
    'Remove any trailing directory separator character
    '
    If Right$(strPathName, 1) = "\" Then
        strPathName = Left$(strPathName, Len(strPathName) - 1)
    End If

    '
    'Attempt to open the file, return value of this function is False
    'if an error occurs on open, True otherwise
    '
    intFileNum = FreeFile
    Open strPathName For Input As intFileNum

    FileExists = IIf(Err, False, True)

    Close intFileNum

    Err = 0

End Function

'---------------------------------------------------------
Public Function ReplaceQuotes(sStr As String) As String
'---------------------------------------------------------
' Reaplace all single quotes in sStr with two single quotes
' To be used for strings in SQL statements
'---------------------------------------------------------

    ReplaceQuotes = Replace(sStr, "'", "''")
    
End Function

'----------------------------------------------------------------------
Public Function ReplaceDoubleQuotes(sText As String) As String
'----------------------------------------------------------------------
'function to double up double quotes so that the literal string can be placed inside quotes
'basically it doubles them up
'----------------------------------------------------------------------

    ReplaceDoubleQuotes = Replace(sText, DOUBLE_QUOTE, DOUBLE_QUOTE & DOUBLE_QUOTE)

End Function

'----------------------------------------------------------------------------------------'
Public Function TextWrapLines(ByVal sText As String, lCharLength As Long) As Long
'----------------------------------------------------------------------------------------'
' return number of lines if text is wrapped after lCharLength characters
' (does not split words over a line)
'----------------------------------------------------------------------------------------'
Dim lMarker As Long
Dim sChar As String
Dim sPortion As String
Dim lLines As Long

    sPortion = sText

    Do While sPortion <> ""

        For lMarker = lCharLength To 1 Step -1
            sChar = Mid(sPortion, lMarker, 1)
            If (sChar = " ") Or (sChar = vbCrLf) Then
                Exit For
            End If
        Next

        If lMarker = 0 Then
            lMarker = InStr(sPortion, vbCrLf)
            If lMarker = 0 Then
                lMarker = InStr(sPortion, " ")
                If lMarker = 0 Then
                    lMarker = lCharLength
                End If
            End If
        End If
        sPortion = Mid(sPortion, lMarker + 1)
        lLines = lLines + 1
    Loop

    TextWrapLines = lLines

End Function

'----------------------------------------------------------------------------------------'
Public Function StringFromFile(sFileName As String) As String
'----------------------------------------------------------------------------------------'
' Return contents of a file as a string
'----------------------------------------------------------------------------------------'
Dim n As Integer

    n = FreeFile
    Open sFileName For Input As n
    
    StringFromFile = Input(LOF(n), n)
    
    Close n

End Function

'----------------------------------------------------------------------------------------'
Public Sub StringToFile(sFileName As String, sText As String)
'----------------------------------------------------------------------------------------'
' Write string to given file
'----------------------------------------------------------------------------------------'
Dim n As Integer

    n = FreeFile
    Open sFileName For Output As n
    
    Print #n, sText
    
    Close n

End Sub

'----------------------------------------------------------------------------------------'
Public Function ConvertFromNull(vVariable As Variant, Optional nType As Integer = NULL_INTEGER) As Variant
'----------------------------------------------------------------------------------------'
' Convert to predefined values if NULL
'----------------------------------------------------------------------------------------'

    If nType = NULL_INTEGER Then
        nType = VarType(vVariable)
    End If

    If IsNull(vVariable) Then
 
        Select Case nType
        Case vbInteger: ConvertFromNull = NULL_INTEGER
        Case vbLong: ConvertFromNull = NULL_LONG
        Case vbDouble: ConvertFromNull = NULL_DOUBLE
        Case vbString: ConvertFromNull = NULL_STRING
        Case DATATYPE_DATE: ConvertFromNull = NULL_DATE
        Case Else: ConvertFromNull = vVariable
        End Select
    
    Else
        ConvertFromNull = vVariable
    End If
    
End Function

'----------------------------------------------------------------------------------------'
Public Function ConvertToNull(ByRef vVariable As Variant, Optional nType As Integer = NULL_INTEGER) As Variant
'----------------------------------------------------------------------------------------'

    ConvertToNull = vVariable

    If nType = NULL_INTEGER Then
        nType = VarType(vVariable)
    End If
    
    Select Case nType
    Case vbInteger: If vVariable = NULL_INTEGER Then ConvertToNull = Null
    Case vbLong: If vVariable = NULL_LONG Then ConvertToNull = Null
    Case vbDouble: If vVariable = NULL_DOUBLE Then ConvertToNull = Null
    Case vbString: If vVariable = NULL_STRING Then ConvertToNull = Null
    Case DATATYPE_DATE: If vVariable = NULL_DATE Then ConvertToNull = Null
    End Select

End Function

'----------------------------------------------------------------------------------------'
Public Function DataTypeDefault(ByVal nType As Integer) As String
'----------------------------------------------------------------------------------------'

        Select Case nType
        Case vbInteger: DataTypeDefault = NULL_INTEGER
        Case vbLong: DataTypeDefault = NULL_LONG
        Case vbSingle: DataTypeDefault = NULL_LONG
        Case vbDouble: DataTypeDefault = NULL_DOUBLE
        Case vbString: DataTypeDefault = NULL_STRING
        Case DATATYPE_DATE: DataTypeDefault = NULL_DATE
        Case Else: DataTypeDefault = "NULL"
        End Select

End Function

'----------------------------------------------------------------------------------------'
Public Function DataTypeDeclaration(ByVal nType As Integer) As String
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'
    Select Case nType
    Case vbEmpty: DataTypeDeclaration = "Empty"
    Case vbString: DataTypeDeclaration = "String"
    Case vbVariant: DataTypeDeclaration = "Variant"
    Case vbBoolean: DataTypeDeclaration = "Boolean"
    Case vbInteger: DataTypeDeclaration = "Integer"
    Case vbLong: DataTypeDeclaration = "Long"
    Case vbDouble: DataTypeDeclaration = "Double"
    Case Else: DataTypeDeclaration = "VarType " & Format(nType)
    End Select
End Function

'----------------------------------------------------------------------------------------'
Public Function CollectionSerialise(ByVal colCollection As Collection) As String
'----------------------------------------------------------------------------------------'
' Serialise a collection (all objects in the collection must be serialisable)
'----------------------------------------------------------------------------------------'
Dim i As Long
Dim pb As PropertyBag

    Set pb = New PropertyBag
    pb.WriteProperty PB_COUNT, colCollection.Count
    
    For i = 1 To colCollection.Count
        pb.WriteProperty PB_ITEM & i, colCollection(i)
    Next
    
    CollectionSerialise = pb.Contents
    
    Set pb = Nothing
    
End Function

'----------------------------------------------------------------------------------------'
Public Function CollectionDeSerialise(ByVal sByteArray As String, _
                    Optional bUseKey As Boolean = False) As Collection
'----------------------------------------------------------------------------------------'
' DeSerialise a previously serialised collection
' If bUseKey = true, assume all items in collection are objects
' and that they have a Key property
' revisions
' ic 23/03/2004 set colCollection = nothing
'----------------------------------------------------------------------------------------'
Dim pb As PropertyBag
Dim colCollection As Collection
Dim i As Long
Dim ByteArray() As Byte
Dim vObj As Object

    Set pb = New PropertyBag
    ByteArray = sByteArray
    pb.Contents = ByteArray
    
    Set colCollection = New Collection
    
    For i = 1 To pb.ReadProperty(PB_COUNT)
        If bUseKey Then
            ' Use the key stored in the object
            Set vObj = pb.ReadProperty(PB_ITEM & i)
            colCollection.Add vObj, vObj.Key
        Else
            colCollection.Add pb.ReadProperty(PB_ITEM & i)
        End If
    Next
    
    Set CollectionDeSerialise = colCollection
    
    Set pb = Nothing
    Set colCollection = Nothing
    
End Function

'----------------------------------------------------------------------------------------'
Public Function CollectionMember(colCollection As Collection, _
                        sKey As String, Optional bObject As Boolean = True) As Boolean
'----------------------------------------------------------------------------------------'
' NCJ 25/5/01
' Returns TRUE if item with given Key exists in Collection,
' or FALSE otherwise
' bObject says whether the collection contains objects or "atomic" items
'----------------------------------------------------------------------------------------'
Dim v As Variant
   
    On Error GoTo ErrHandler
    If bObject Then
        ' Use Set for objects
        Set v = colCollection.Item(sKey)
    Else
        ' Use =
        v = colCollection.Item(sKey)
    End If
    
    CollectionMember = True
    
    Exit Function
    
ErrHandler:
    
    If Err.Number = 5 Then  'Invalid procedure call or argument - this occurs when item doesn't exist in collection
        CollectionMember = False
    Else
        Err.Raise Err.Number
    End If
    
    
End Function

'----------------------------------------------------------------------------------------'
Public Function CollectionToArray(colCollection As Collection) As Variant
'----------------------------------------------------------------------------------------'
' Converts a collection to an array
' Returns null if empty collection
'----------------------------------------------------------------------------------------'
Dim vArray As Variant
Dim i As Long
    If colCollection.Count = 0 Then
        CollectionToArray = Null
    Else
        ReDim vArray(colCollection.Count - 1)
        For i = 0 To UBound(vArray)
            vArray(i) = colCollection(i + 1)
        Next
        CollectionToArray = vArray
    End If
            

End Function

'----------------------------------------------------------------------------------------'
Public Sub CollectionAddAnyway(colCollection As Collection, _
                        vItem As Variant, Optional sKey As String = "")
'----------------------------------------------------------------------------------------'
' NCJ 25/5/01
' DO NOT USE THIS SUB UNLESS YOU ARE SURE YOU WANT TO IGNORE ERRORS!!!
' Adds item to Collection and ignores errors if it already exists
'----------------------------------------------------------------------------------------'

    On Error GoTo Ignore
    If sKey > "" Then
        colCollection.Add vItem, sKey
    Else
        colCollection.Add vItem
    End If
    
Ignore:

End Sub

'----------------------------------------------------------------------------------------'
Public Sub CollectionRemoveAnyway(colCollection As Collection, sKey As String)
'----------------------------------------------------------------------------------------'
' NCJ 31 Oct 02
' DO NOT USE THIS SUB UNLESS YOU ARE SURE YOU WANT TO IGNORE ERRORS!!!
' Removes item from Collection and ignores errors if it does not exist
'----------------------------------------------------------------------------------------'

    On Error GoTo Ignore

    colCollection.Remove sKey
    
Ignore:

End Sub

'-----------------------------------------------------
Public Function RegionalDecimalPointChar() As String
'-----------------------------------------------------
' Return local character used for decimal point
' (as in Windows Regional Settings)
' Do this by formatting 0.1 and looking at second char
'-----------------------------------------------------
Dim dblNum As Double
Dim sNum As String

    dblNum = 1 / 10     ' Set it to 0.1
    sNum = Format(CStr(dblNum), "0.0")
    RegionalDecimalPointChar = Mid(sNum, 2, 1)

End Function

'-----------------------------------------------------
Public Function RegionalThousandSeparatorChar() As String
'-----------------------------------------------------
' Return local character used for thousand separator
' (as in Windows Regional Settings)
' Do this by formatting 1,000 and looking at second char
'-----------------------------------------------------
Dim dblNum As Double
Dim sNum As String

    dblNum = 1000
    sNum = Format(CStr(dblNum), "0,000")
    RegionalThousandSeparatorChar = Mid(sNum, 2, 1)

End Function

'-------------------------------------------------------------------------
Public Function LocalNumToStandard(ByVal sLocalNumStr As String, _
                        Optional bIncludeThousands As Boolean = False, _
                        Optional sDecPt As String = "", _
                        Optional sThouSep As String = "") As String
'-------------------------------------------------------------------------
' Convert number string in "regional" format to "standard" format
' Remove all thousand separators
' and replace local decimal point with dot
' MLM 08/03/01: Always remove thousands separators
' NCJ 17 Jun 02 - CBB 2.2.14/19 Only remove thousands if bIncludeThousands = False
' NCJ 12 May 03 - Added optional Dec. Pt. and Thou. Sep. specifications
'-------------------------------------------------------------------------
Dim n As Integer
Dim sLocalDot As String
Dim sLocalComma As String
Dim sStdStr As String
Dim sCurChar As String

    sStdStr = sLocalNumStr  ' Initially assume current value
    If sLocalNumStr > "" Then
        ' Get the regional settings
        ' NCJ 12 May 03 - Use those passed in (if any)
        If sDecPt > "" Then
            sLocalDot = sDecPt
        Else
            ' Use machine's own
            sLocalDot = RegionalDecimalPointChar
        End If
        
        If sThouSep > "" Then
            sLocalComma = sThouSep
        Else
            ' Use machine's own
            sLocalComma = RegionalThousandSeparatorChar
        End If
        
        ' Do we need to convert?
        If sLocalDot <> msSTANDARD_DOT Or sLocalComma <> msSTANDARD_COMMA Then
            ' Regional setting is different from standard
            sStdStr = ""
            For n = 1 To Len(sLocalNumStr)
                ' Pick up next character
                sCurChar = Mid(sLocalNumStr, n, 1)
                ' Convert if necessary
                Select Case sCurChar
                Case sLocalDot
                    ' Replace with standard dot
                    sStdStr = sStdStr & msSTANDARD_DOT
                Case sLocalComma
                    'TA 01/03/2001: do not put in thousand separators
                    ' NCJ 17 Jun 02 - Do put them in if bIncludeThousands = True
                    ' Replace with standard comma
                    If bIncludeThousands Then
                        sStdStr = sStdStr & msSTANDARD_COMMA
                    End If
                Case Else
                    ' Leave it unchanged
                    sStdStr = sStdStr & sCurChar
                End Select
            Next n
        Else
            'MLM 08/03/01: If regional setting is same as standard, only remove thousands separators
            ' NCJ 17 Jun 02 - Only remove them if bIncludeThousands = False
            If Not bIncludeThousands Then
                sStdStr = Replace(sLocalNumStr, msSTANDARD_COMMA, "")
            End If
        End If
    End If
    LocalNumToStandard = sStdStr

End Function

'-------------------------------------------------------------------------
Public Function StandardNumToLocal(ByVal sStdNumStr As String, _
                        Optional sDecPt As String = "", _
                        Optional sThouSep As String = "") As String
'-------------------------------------------------------------------------
' Replace all dots and commas in sStdNumStr
' with local decimal points and thousand separators
' Assume sStdNumStr contains a number with dots and commas
' NCJ 12 May 03 - Added optional Dec. Pt. and Thou. Sep. specifications
'-------------------------------------------------------------------------
Dim n As Integer
Dim sLocalDot As String
Dim sLocalComma As String
Dim sLocalStr As String
Dim sCurChar As String

    sLocalStr = sStdNumStr  ' Initially assume current value
    If sStdNumStr > "" Then
        ' Get the regional settings
        ' NCJ 12 May 03 - Use those passed in (if any)
        If sDecPt > "" Then
            sLocalDot = sDecPt
        Else
            ' Use machine's own
            sLocalDot = RegionalDecimalPointChar
        End If
        
        If sThouSep > "" Then
            sLocalComma = sThouSep
        Else
            ' Use machine's own
            sLocalComma = RegionalThousandSeparatorChar
        End If
        
        ' Do we need to convert?
        If sLocalDot <> msSTANDARD_DOT Or sLocalComma <> msSTANDARD_COMMA Then
            ' Regional setting is different from standard
            sLocalStr = ""
            For n = 1 To Len(sStdNumStr)
                ' Pick up next character
                sCurChar = Mid(sStdNumStr, n, 1)
                ' Convert if necessary
                Select Case sCurChar
                Case msSTANDARD_DOT
                    ' Replace with local dot
                    sLocalStr = sLocalStr & sLocalDot
                Case msSTANDARD_COMMA
                    ' Replace with local comma
                    sLocalStr = sLocalStr & sLocalComma
                Case Else
                    ' Leave it unchanged
                    sLocalStr = sLocalStr & sCurChar
                End Select
            Next n
        End If
    End If
    StandardNumToLocal = sLocalStr

End Function

'--------------------------------------------------------------------------
Public Function RndLong(ByVal lMax As Long, Optional bFromZero As Boolean = False) As Long
'-------------------------------------------------------------------------
' Retuns a random long >= 1 (>=0 if bFromZero) and <= lMax
'-------------------------------------------------------------------------

    If bFromZero Then
        RndLong = Int((lMax + 1) * Rnd)
    Else
        RndLong = Int(lMax * Rnd) + 1
    End If

End Function

'------------------------------------------------
Public Function Max(ByVal l1 As Long, ByVal l2 As Long) As Long
'------------------------------------------------
' Return the maximum of two longs
'------------------------------------------------

    If l1 >= l2 Then
        Max = l1
    Else
        Max = l2
    End If
    
End Function

'------------------------------------------------
Public Function Min(ByVal l1 As Long, ByVal l2 As Long) As Long
'------------------------------------------------
' Return the minimum of two longs
'------------------------------------------------

    If l1 < l2 Then
        Min = l1
    Else
        Min = l2
    End If
    
End Function

'----------------------------------------------------------------------------------------'
Public Function RangeIncludesValue(Optional vValue As Variant = Null, Optional vMin As Variant = Null, Optional vMax As Variant = Null, _
                                Optional bInclusive As Boolean = True, Optional bAllowNullValue As Boolean = False) As Boolean
'----------------------------------------------------------------------------------------'
' returns true if value is in range
' when Binclusive is true - if bAllowNullValue is true then null is allowed
'   if either end of range is null, otherwise both ends must be null
'----------------------------------------------------------------------------------------'
    
    If VarType(vValue) = vbNull Then
        If bAllowNullValue Then
            'min OR max must be null , inclusive
            RangeIncludesValue = (VarType(vMin) = vbNull) Or (VarType(vMax) = vbNull) And bInclusive
        Else
            'min AND max must be null, inclusive
            RangeIncludesValue = (VarType(vMin) = vbNull) And (VarType(vMax) = vbNull) And bInclusive
        End If
    Else
            'min and max are null                       OR
            'min is null and value less than max        OR
            'max is null and value greater than min     OR
            'value is between min and max
        If bInclusive Then
            'min and max inclusive
            RangeIncludesValue = ((VarType(vMin) = vbNull) And (VarType(vMax) = vbNull)) _
                Or ((VarType(vMin) = vbNull) And (vValue <= vMax)) _
                Or ((VarType(vMax) = vbNull) And (vValue >= vMin) _
                Or ((vValue >= vMin) And (vValue <= vMax)))
        Else
            'min and max exclusive
            RangeIncludesValue = ((VarType(vMin) = vbNull) And (VarType(vMax) = vbNull)) _
                Or ((VarType(vMin) = vbNull) And (vValue < vMax)) _
                Or ((VarType(vMax) = vbNull) And (vValue > vMin) _
                Or ((vValue > vMin) And (vValue < vMax)))
            
        End If
    End If


End Function

'----------------------------------------------------------------------------------------'
Public Function RangeInRange(Optional vRange1Min As Variant = Null, Optional vRange1Max As Variant = Null, _
                                Optional vRange2Min As Variant = Null, Optional vRange2Max As Variant = Null, Optional bInclusive As Boolean = True) As Boolean
'----------------------------------------------------------------------------------------'
'returns true if range 1 inside range 2
'----------------------------------------------------------------------------------------'
    
   'true if min1 in range2 or max1 in range2
    If (VarType(vRange1Min) = vbNull And VarType(vRange1Max) = vbNull) And (VarType(vRange2Min) = vbNull And VarType(vRange2Max) = vbNull) Then
        'range 1 or 2 isn't a range - don't bother checking
        RangeInRange = bInclusive
    Else
        If (VarType(vRange1Min) = vbNull And VarType(vRange2Min) <> vbNull) Or (VarType(vRange1Max) = vbNull And VarType(vRange2Max) <> vbNull) Then
            'range 1 not bound and range 2 is
            RangeInRange = False
        Else
            RangeInRange = RangeIncludesValue(vRange1Min, vRange2Min, vRange2Max, bInclusive, True) And RangeIncludesValue(vRange1Max, vRange2Min, vRange2Max, bInclusive, True)
        End If
    End If

End Function

'----------------------------------------------------------------------------------------'
Public Function RangeOverlap(Optional vRange1Min As Variant = Null, Optional vRange1Max As Variant = Null, _
                                Optional vRange2Min As Variant = Null, Optional vRange2Max As Variant = Null) As Boolean
'----------------------------------------------------------------------------------------'
'returns true if range 1 and range 2 overlap
'----------------------------------------------------------------------------------------'

'nb. min and max inclusive
   RangeOverlap = RangeIncludesValue(vRange1Min, vRange2Min, vRange2Max) Or RangeIncludesValue(vRange1Max, vRange2Min, vRange2Max) _
                    Or RangeIncludesValue(vRange2Min, vRange1Min, vRange1Max)

End Function

'----------------------------------------------------------------------------------------'
Public Function RangeValid(ByVal vMin As Variant, ByVal vMax As Variant) As Boolean
'----------------------------------------------------------------------------------------'
'returns true if range min is greater than max
'----------------------------------------------------------------------------------------'
Dim bRangeValid As Boolean
    
    bRangeValid = True
    If Not (VarType(vMin) = vbNull Or VarType(vMax) = vbNull) Then
        If vMin > vMax Then
            bRangeValid = False
        End If
    End If
    
    RangeValid = bRangeValid
            

End Function

'----------------------------------------------------------------------------------------'
Public Function ByteArrayFromFile(sFileName As String) As Byte()
'----------------------------------------------------------------------------------------'
' REM 05/04/02
' Return contents of a file as Byte Array
'----------------------------------------------------------------------------------------'
Dim n As Integer

    n = FreeFile
    Open sFileName For Binary As n
    
    ByteArrayFromFile = InputB(LOF(n), n)
    
    Close n

End Function

'----------------------------------------------------------------------------------------'
Public Function IMedNow() As Double
'----------------------------------------------------------------------------------------'
' NCJ 2 Oct 02
' Returns double representing current timestamp accurate to 1/100 sec.
' This should always be used instead of CDbl(Now)
'----------------------------------------------------------------------------------------'
Dim sglTimer As Single
Dim dblNow As Double
Dim dtNow As Date

    ' Grab the Timer and the Now datestamp
    sglTimer = Timer        ' accurate to 1/100 sec.
    dtNow = Now             ' accurate to 1 sec
    
    ' Convert datestamp to a double
    dblNow = CDbl(dtNow)

    ' Get the hundredths of a second from the Timer,
    ' convert to a fraction of a day and add to the datestamp
    IMedNow = dblNow + (((sglTimer - Int(sglTimer)) / 60) / 60) / 24
    
End Function

'----------------------------------------------------------------------------------------'
Public Function ServerExists(sServerName As String, sTestClass As String) As eServerExistsResult
'----------------------------------------------------------------------------------------'
'Returns whether a server exists on a network
'the classname passed in must be a class registered on the clinet machine eg MACROTimeZone30.Timezone
'----------------------------------------------------------------------------------------'
Dim o As Object

    On Error GoTo ErrLabel
    
    Set o = CreateObject(sTestClass, sServerName)
    
    ServerExists = serYes
    
    Exit Function
    
ErrLabel:
    Select Case Err.Number
    Case 462: ServerExists = serNo 'The remote server machine does not exist or is unavailable
    Case Else: ServerExists = serMaybe
    End Select
    
    
End Function
