Attribute VB_Name = "modHTML"
'-----------------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       modHTML.bas
'   Author:         Andrew Newbigging, June 1997
'   Purpose:    Generates HTML versions of MACRO forms
'
'-----------------------------------------------------------------------------------------------'
'
'-----------------------------------------------------------------------------------------------'
'   Revisions:
'   1   Andrew Newbigging   24/02/98
'   2   Joanne Lau          5/03/98
'   3   Joanne Lau          13/03/98
'   4   Joanne Lau          26/03/98
'   5   Andrew Newbigging   02/04/98
'   6   Andrew Newbigging   30/04/98
'   7   Andrew Newbigging   15/05/98
'   8   Andrew Newbigging   15/05/98
'   9   Joanne Lau          30/06/98
'   10  Joanne Lau          30/06/98
'   11  Joanne Lau          30/06/98
'   12  Joanne Lau          14/07/98
'   13  Andrew Newbigging   04/5/99
'   Modified CreateHTMLCOmponents,CreateHTMLFiles:
'   changed reference path to use the virtual directory 'WebRDEDocuments'
'   14  Joanne Lau          31/07/99
'   Added   CreateNetscapeFiles.
'   Added   JavascriptNNvalidation additional argument of ControlType.
'   ?    joanne Lau         18/08/99
'   Modified CreateHTMLFiles. Included z-index in style attribute for Controls. This
'   prevents the invisible layer containing comments to overlap a
'   text field making it inaccessable to the cursor. A z-index of 1 will be the highest in the stack.
'   14  NCJ 1 Sep 99
'       Need to deal with MandatoryValidation and WarningValidation
'       which are no longer in CRFElement recordset (not done yet!)
'       Mo Morris   1/11/99
'       DAO to ADO conversion
'   SDM 10/11/99    Copied in the error handler routines
'   WillC    Changed the Following where present from Integer to Long  ClinicalTrialId
'           CRFPageId,VisitId,CRFElementID
'   JL 06/02/00 Absolute positioning not working in rsValueData Recordset (ADODB recordeset?)
'       - removed and replaced with record counter variable.
'   JL 14/02/00 For each element ; whenever we access R("responseValue") recordset (defined in dataValues.asp), put in
'       a check that it has a Response value before accessing the recordset . This is because of changes that
'       only fields which has a response value has a record in the DataItemResponse Table,
'       whereas previously empty fields also had an entry in DataItemResponse.
'       Same changes to getDataValue.
'   JL Bug: 2810    28/02/00
'       added flags bNoOfelements & bNoOfpages to test for study with no eforms, and
'       eform with no elements/questions.
'   JL 24/05/00. Changes to CreateNetscapeFiles. (1) Added checks in the 1st rsCrfElement Loop that the control
'       is Not a radio. (2) In the 2nd Loop RsRadioElement which deals with Radios, added 'Order by FieldOrder' in sql
'       to ensure we print out the radios in decending order of fieldOrder. It worked previously because crfElementId
'       number just happened to be the same as fieldorder.
'   JL 26/05/00. Changed from R(""ResponseValue"") to R(""ValueCode""). So that the javascript Validation can compare
'       datavalues can't do this using the string value.
'-----------------------------------------------'
'   Macro Version 2.1                           '
'-----------------------------------------------'
'JL             Added option Explicit
'JL 13/07/00    Always call SetStatusImage, if status=Requested then the image is set to a dummy transparent image.
'JL  26/07/00   Replaced all references to control types with new Constants.
'JL  26/07/00   Bug. Most Captions and comments are suddenly not being displayed ?????
'               Wrongly printing a Caption for a <HR>. Code Attempting to read Caption field of a <HR> element.
'               It printed a strange tag which was Tag which was interferring with all pages containing Lines.
'JL  26/07/00   Added calls to CalulateStatusXPosition and SetStatusImage for each field.
'               Identical for IE sub except for position:absolute; in the style tag and Layer tag.
'JL  27/07/00   Removed References to R Recordset replaced with getResponseValue.
'               Changes made to GetDataValue as well.
'JL  27/07/00   Removed References to R2 Recordset for radios replaced with getResponseValue.
'JL  28/07/00   Changed subscript from asp to txt TypeChecking.txt and is now included in this mod instead
'               of in datavalues.asp and in datavaluesNN.asp
'JL  02/08/00   Added call to disableLinks() on the Body's onUnload event. This is to prevent the user from clicking
'               any of the buttons whilst a new page is loading.
'               Added call to load_Links() which enables the navigation buttons on the left frame of the browser.
'               Placed the function call in a New layer, which will be the last element loaded in the document body.
'               This ensures that the buttons are available only after the new page has loaded.
'JL  07/08/00   Updated Code in CreateHTMLFiles.
'               Removed all references to R recordset, replaced it with getResponseValue. This function returns
'               the value or an empty string if theres no response value. Previously very cumbersome using recordsets,
'               because of having to check if the item was contained in the recordset.
'JL  21/08/00   Added new include file LogoutUser.txt.
'JL  08/09/00   Added Function JavascriptDisableCaptions.
'JL  11/09/00   Added Function JavascriptUpdateRadioValues.
'JL  19/09/00   Added new Include file in ASP's JSObjects.htm.
'MLM 18/09/00   SR 3858: Check that user is logged in and has "View data" user role in CreateHTMLFiles and CreateNetscapeFiles
'MLM 19/09/00   Check for null captions in ElementWidth and CellContents
'JL  28/09/00   Added Include file Format.htm to beginning of the file.
'MLM 13/10/00   Modified CreateFormObject to make global JavaScript object
'               (instead of creating new properties and methods for existing form object)
'NCJ 24/10/00   Use new gsWEB_HTML_LOCATION for HTML forms
'JL 26/10/00    Added test for forms with no questions (check rsDataItem.eof) and with no form
'               elements (bNoOfelements = false).
'MLM 26/10/00   Tidied up comments
'               Added "Change Data" permission as urgument to GetFormArezzoDetails in CreateHTMLFiles
'               Removed []s from SLQ in CreateNetscapeFiles for Oracle/SQL Server compatibility
'JL  09/11/00   Added string to call javascript function AddRadioClearButton.
'MLM 09/11/00   Added check of user's site permissions to CreateHTMLFiles and CreateNetscapeFiles
'ic  10/11/00   added recordset declaration to CreateNetscapeFiles
'               added call to CreateFormObject in function CellContents
'Mo 14/11/00    Calls to internal function ReplaceCharacters replaced
'               with calls to VB function Replace.
'ic  14/11/00   added event handler to elements in CellContents
'ic  16/11/00   changed form name to myform in CreateNetscapeFiles
'ic  17/11/00   added case option for multimedia element in CellContents
'ic  19/11/00   added 'o' prefix to objects in CreateFormObject
'MLM 21/11/00   Added calls to new DLL function OpenSubject to implement subject locking
'ic  23/11/00   added variable declaration in CellContents
'ic  28/11/00   added condition for drawing image spacers to CreateNetscapeFiles
'ic  29/11/00   commented out javascript declaration in CreateNetscapeFiles
'               added 'round function call in CreateNetscapeFiles
'               added condition to limit image spacer size in CreateNetscapeFiles
'               added variable declaration in CellContents
'               added code to prevent wide textboxes in CellContents
'ic  30/11/00   added extra conditions to SQL query in CreateNetscapeFiles
'               added code to contain elements in table in CellContents
'ic  01/12/00   added condition so that if the element width or height is 0 (it cannot be found) it is omitted in CreateNetscapeFiles
'               added check to see if picture exists before loading it. this was causing a run time error (logged 53) in ElementHeight
'               added check to see if picture exists before loading it. this was causing a run time error (logged 53) in ElementWidth
'               added condition so that if the element  cannot be found it is omitted in CreateHTMLFiles
'ic  05/12/00   added declaration and copy image file code to CreateHTMLFiles
'ic  06/12/00   applied style to radio controls in CellContents
'ic  21/12/00   added onblur event to radios so that answers are evaluated when tabbing between fields as well as clicking in CreateHTMLFiles
'ic  22/12/00   added nElementHeight, nElementWidth to CellContents function call so that images can be dynamically sized in CreateNetscapeFiles
'               added declarations for X Y calculations in CreateNetscapeFiles
'               added code to change negative x y coordinates to 0 in CreateNetscapeFiles
'               added condition to eliminate hidden captions from recordset in CreateNetscapeFiles
'               added code so that if a picture has been positioned partly off the top of the page, height is reduced in ElementHeight
'               added code so that if a picture has been positioned partly off the side of the page, width is reduced in ElementWidth
'               changed code to write radiobuttons into asp rather than by calling 'Radio' function from asp in CellContents
'               removed code that writes elements in container tables as it prevents tabbing in CellContents
'               changed element width calculation for textboxes in CellContents
'               added height and width parameters to image code to resize images if x,y coordinates are negative in CellContents
'               added height and width arguements to CellContents function to size images in CellContents
'               added condition so that if a text element is hidden, a hidden field is written to CellContents
'ic  12/02/01   removed gsDOCUMENTS_PATH constant from picture code as it was being written in the html in CreateHTMLFiles
'jl   15/03/01  Changes made to JavascriptDisableCaptions. See notes.
'-----------------------------------------------'
'   Macro Version 2.2                           '
'-----------------------------------------------'
'ic  21/08/01   commented out CreateHTMLFiles sub, rewrote for new JSVE structure
'ic  23/08/01   SR4415 function FieldAttributes_Explorer - check if optional parameter is missing
'DPH 17/10/2001 Added calls to FolderExistence routine to create missing folders
'DPH 29/10/2001 Fixes to Creating HTML - Active categories & Question Numbering
'DPH 30/10/2001 Completing Form Definitions
'DPH 5/11/2001 Order Categories to ValueOrder when generating HTML
'ASH 11/01/2002 - Rewrote sql for use on oracle in sub formDefinition
'
'ic  02/04/2002 commented out CreateHTMLFiles() replaced with PublishEFormASPFiles()
'MLM 15/05/03: Added various CLngs in lTabIndex calculations to cope with long eForms.
'---------------------------------------------------------------------------------------------------------------------|
'------------------------------------------------
'   Macro Version 3.
'------------------------------------------------
'ZA 22/07/2002 - updated font/colour properties for caption and comments
' ic 05/06/2003 added return button in OutputEformFile()
' DPH 20/06/2003 - only use active categories / font style for js field template
' DPH 04/09/2003 - Create tab order for 'zero' order elements
' ic 17/09/2003 skip hidden controls in RtnEformCategoryTxtFunc(),RtnRadioCreateFuncs()
' ic 15/02/2005 bug 2529, only add active categories in RtnEformElement()
' ic 27/05/2005 issue 2579, dont add hidden fields for comments/lines/pictures in RtnEformElement()
' DPH 29/06/2005 - issue 2412 keep note/comment icon in text box view when display length shorter than
'                   actual length of field in RtnEformElement()
' ic 07/03/2006 moved GetQuestionDefinition(), GetRFC() and GetRFO() from IO to solver permissions problem
'               when generating html
' DPH 05/03/2007 - Bug 2545. Default display size to actual size up to max of 60 in SetElementWidth
' ic 16/04/2007 issue 2759, added the derivation expression
Option Explicit
Option Base 0
Option Compare Binary

'JL 16/06/00
Public Const gsTEXT_BOX = 1
Public Const gsOPTION_BUTTONS = 2
Public Const gsPOPUP_LIST = 4
Public Const gsCALENDAR = 8
Public Const gsATTACHMENT = 32
Public Const gsPUSH_BUTTONS = 256
Public Const gsVISUAL_ELEMENT = 16384
Public Const gsLINE = 16385
Public Const gsCOMMENT = 16386
Public Const gsPICTURE = 16388

'ic 22/12/00
'changed nXScale from 12 to 12.5 for improvement in page sizing
' DPH 13/05/2003 - changed X scale to make page width smaller
Const nXSCALE As Integer = 15 '12.5
Const nYSCALE As Integer = 15
Const nIeXSCALE = 15 '12.5
Const nIeYSCALE = 15 '15.7 '15.672

Const mnTABINDEXGAP As Integer = 200

'Enum ControlType
'    TextBox = 1
'    OptionButtons = 2
'    PopUp = 4
'    Calendar = 8
'    RichTextBox = 16
'    Attachment = 32
'    PushButtons = 258
'    Line = 16385
'    Comment = 16386
'    Picture = 16388
'End Enum

'--------------------------------------------------------------------------------------------------
Public Function GetQuestionDefinition(ByRef oEformElement As eFormElementRO, _
                             Optional ByVal enInterface As eInterface = iwww) As String
'--------------------------------------------------------------------------------------------------
'   ic 14/10/02
'   builds and returns an html table representing a question audit trail
'   revisions
'   ic 31/03/2003 adjusted validation table width
'   ic 16/04/2007 issue 2759, added the derivation expression
'--------------------------------------------------------------------------------------------------

Dim oValidation As Validation
Dim sHTML As String
    
    
    sHTML = sHTML & "<html>" & vbCrLf _
                  & "<head>" & vbCrLf _
                  & "<title>Question Definition</title>" & vbCrLf
                  
    If (enInterface = iwww) Then
        sHTML = sHTML & "<link rel='stylesheet' href='../../../style/MACRO1.css' type='text/css'>" & vbCrLf
    End If
    
    sHTML = sHTML & "</head>" & vbCrLf _
                  & "<body>"
    
    'outer table, code,age,help,format rows
    sHTML = sHTML & "<table align='center' border='0' width='95%'>" & vbCrLf _
                    & "<tr height='10'><td></td></tr>" & vbCrLf _
                    & "<tr height='15' class='clsLabelText'>" _
                      & "<td colspan='2'><b>Question: " & oEformElement.Code & "</b></td>" _
                      & "<td width='10%'><a style='cursor:hand;' "
                      
    If (enInterface = iwww) Then
        sHTML = sHTML & "onclick='javascript:window.close();'"
    Else
        sHTML = sHTML & "onclick='javascript:window.navigate(" & """" & "VBfnClose" & """" & ");'"
    End If
    
    sHTML = sHTML & "><u>Close</u></a></td>" _
                    & "</tr>" _
                    & "<tr height='15'><td></td></tr>" & vbCrLf _
                    & "<tr height='15' class='clsTableText'>" _
                      & "<td width='20%'>Code:</td><td width='70%' colspan='2'>" & oEformElement.Code & "</td>" _
                    & "</tr>" & vbCrLf _
                    & "<tr height='15' class='clsTableText'>" _
                      & "<td>Name:</td><td colspan='2'>" & oEformElement.Name & "</td>" _
                    & "</tr>" & vbCrLf _
                    & "<tr height='15' class='clsTableText'>" _
                      & "<td>User help text:</td><td colspan='2'>" & oEformElement.Helptext & "</td>" _
                    & "</tr>" & vbCrLf _
                    & "<tr height='15' class='clsTableText'>" _
                      & "<td>Format:</td><td colspan='2'>" & oEformElement.Format & "</td>" _
                    & "</tr>" & vbCrLf _
                    & "<tr height='15' class='clsTableText'>" _
                      & "<td>Collect data if:</td><td colspan='2'>" & oEformElement.CollectIfCond & "</td>" _
                    & "</tr>" & vbCrLf _
                    & "<tr height='15' class='clsTableText'>" _
                      & "<td>Derivation:</td><td colspan='2'>" & oEformElement.DerivationExpr & "</td>" _
                    & "</tr>" & vbCrLf
                   
    If (oEformElement.Validations.Count > 0) Then
        sHTML = sHTML & "<tr height='10'><td></td></tr>" _
                      & "<tr height='20' valign='middle'>" _
                        & "<td class='clsTableText' colspan='3'>Validation:</td>" _
                      & "</tr>" _
                      & "<tr>" _
                        & "<td colspan='3'>"
               
        'inner validation table
        sHTML = sHTML & "<table border='1' width='100%'>" _
                        & "<tr height='15' class='clsTableHeaderText'>" _
                          & "<td width='10%'>Type</td>" _
                          & "<td width='45%'>Validation</td>" _
                          & "<td width='45%'>Message</td>" _
                        & "</tr>"
    
        'validation loop
        For Each oValidation In oEformElement.Validations
            sHTML = sHTML & "<tr height='15' class='clsTableText'>" _
                            & "<td>" & GetValidationTypeString(oValidation.ValidationType) & "</td>" _
                            & "<td>" & oValidation.ValidationCond & "</td>" _
                            & "<td>" & oValidation.MessageExpr & "</td>" _
                          & "</tr>" & vbCrLf
    
        Next

        'end of validation table
        sHTML = sHTML & "</table>"
        sHTML = sHTML & "</td></tr>"
    End If
    
    'end of outer table
    sHTML = sHTML & "</table>"

    sHTML = sHTML & "</body>" & vbCrLf _
                  & "</html>"

    Set oValidation = Nothing
    GetQuestionDefinition = sHTML
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetRFC(ByRef oStudyDef As StudyDefRO, _
              Optional ByVal enInterface As eInterface = iwww) As String
'--------------------------------------------------------------------------------------------------
'   ic 30/10/02
'   builds and returns an html table representing a rfc dialog
'--------------------------------------------------------------------------------------------------
' REVISIONS
' DPH 21/05/2003 - OK & Cancel
'--------------------------------------------------------------------------------------------------
Dim sHTML As String
Dim nLoop As Integer
    
    
    sHTML = sHTML & "<html>" & vbCrLf _
                  & "<head>" & vbCrLf _
                  & "<title>Reason For Change</title>" & vbCrLf
    
    If (enInterface = iwww) Then
        sHTML = sHTML & "<link rel='stylesheet' href='../../../style/MACRO1.css' type='text/css'>" & vbCrLf _
                      & "<script language='javascript' src='../../../script/Dialog.js'></script>" & vbCrLf
                      
        sHTML = sHTML & "<script language='javascript'>" & vbCrLf _
                      & "function fnPageLoaded()" & vbCrLf _
                        & "{" & vbCrLf _
                          & "var aQS=fnSplitQS(window.location.search);" & vbCrLf _
                          & "document.all['divName'].innerHTML='<b>Question: '+aQS['name']+'</b>';" & vbCrLf _
                          & "document.all['divLabel'].innerHTML='Please enter or choose the reason for changing '+aQS['name'];" & vbCrLf _
                          & "txtInput.focus();" & vbCrLf _
                        & "}" & vbCrLf _
                      & "</script>" & vbCrLf
    End If

    sHTML = sHTML & "</head>" _
                  & "<body onload='fnPageLoaded();'>"

    ' DPH 21/05/2003 - OK & Cancel
    sHTML = sHTML & "<table align='center' width='95%' border='0'>" & vbCrLf _
                    & "<tr height='10'><td></td></tr>" & vbCrLf _
                    & "<tr height='15' class='clsLabelText'>" _
                      & "<td colspan='2'>" _
                        & "<div id='divName'></div>" _
                      & "</td>" _
                      & "<td><a style='cursor:hand;' onclick='javascript:fnReturn1(txtInput.value);'><u>OK</u></a></td>" _
                      & "<td><a style='cursor:hand;' onclick='javascript:fnReturn(" & Chr(34) & Chr(34) & ");'><u>Cancel</u></a></td>" _
                    & "</tr>" _
                    & "<tr height='15'><td></td></tr>" & vbCrLf
    
    sHTML = sHTML & "<tr height='5'>" _
                      & "<td></td>" _
                    & "</tr>" _
                    & "<tr height='30'>" _
                      & "<td colspan='4' class='clsLabelText'>" _
                        & "<div id='divLabel'></div>" _
                      & "</td>" _
                    & "</tr>" _
                    & "<tr>" _
                      & "<td width='100'></td>" _
                      & "<td><input style='width:180px;' class='clsTextbox' name='txtInput' type='text'></td>" _
                    & "</tr>" _
                    & "<tr>" _
                      & "<td width='100'></td>" _
                      & "<td><select style='width:180px;' name='selRFC' class='clsSelectList' size='3' onchange='javascript:SelectChange(" & Chr(34) & "selRFC" & Chr(34) & "," & Chr(34) & "txtInput" & Chr(34) & ")'>"
                      
    'add rfcs to select list
    With oStudyDef.RFCs
        For nLoop = 1 To .Count
            sHTML = sHTML & "<option value='" & .Item(nLoop) & "'>" _
                          & .Item(nLoop) _
                          & "</option>" & vbCrLf
        Next
    End With
                      
    
    sHTML = sHTML & "</select></td>" _
                  & "</tr>" _
                  & "</table>" _
                  & "</body>" _
                  & "</html>"
    
    
    GetRFC = sHTML
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetRFO(ByRef oStudyDef As StudyDefRO, _
              Optional ByVal enInterface As eInterface = iwww) As String
'--------------------------------------------------------------------------------------------------
'   ic 08/04/2003
'   builds and returns a js string representing rfo options
'--------------------------------------------------------------------------------------------------
' REVISIONS
' DPH 22/05/2003 - Changed to create complete Reject/Warning/Inform dialog
' ic 27/05/2005 changed warning to "W" and reject to "R"
'--------------------------------------------------------------------------------------------------
Dim vDialogHTML() As String
Dim sJS As String
Dim sOverrule As String
Dim sHTML As String
Dim nLoop As Integer

    ' initialise vDialogHTML
    ReDim vDialogHTML(0)
    
    'add rfcs to select list
    With oStudyDef.RFOs
        For nLoop = 1 To .Count
            sJS = sJS & .Item(nLoop) & gsDELIMITER1
        Next
    End With
    If (Len(sJS) > 0) Then sJS = Left(sJS, Len(sJS) - 1)
    sOverrule = "var sOverrule=" & Chr(34) & sJS & Chr(34) & ";"
    
    ' Create Form
    Call AddStringToVarArr(vDialogHTML, "<html>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "<head>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "<title>Reject/Warn/Inform</title>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "<link rel='stylesheet' href='../../../style/MACRO1.css' type='text/css'>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "<script language='javascript' src='../../../script/Dialog.js'></script>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "<script id='rfoScript' language='javascript' src=''></script>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "<script language='javascript'>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "var sDel1=""`"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "var sType;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "var bCollectRFO=true;" & vbCrLf)
    ' sOverrule
    Call AddStringToVarArr(vDialogHTML, sOverrule & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "function fnPageLoaded()" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "var olArg=window.dialogArguments;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "var sName=olArg[0];" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "sType=olArg[1];" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "var sMessage=olArg[2];" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "var sExpression=olArg[3];" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "var bMenu=olArg[4];" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "var sRFO=olArg[5];" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "var sDatabase=olArg[6];" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "var sStudy=olArg[7];" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "var bOverrule=olArg[8];" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "var sTitle;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "var sHeader;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "switch (sType)" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "case ""W"":" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "sTitle=""Warning"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "sHeader = ""The warning for question "" + sName + "" has been generated because:"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "fnWarning();" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "if (sRFO=="""")" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "fnHideRFO();" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "this.returnValue=""R"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "else" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "txtInput.value=sRFO;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "txtInput.focus();" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "this.returnValue=""O""+fnReplaceWithJSChars(sRFO);" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "if (!bOverrule) fnDisableOverrule();" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "break;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "case ""R"":" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "sTitle=""Reject"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "sHeader = ""The value for question "" + sName + "" has been rejected because:"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "this.returnValue=""R"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "break;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "case ""I"":" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "sTitle=""Inform"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "sHeader = ""The flag for question "" + sName + "" has been generated because:"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "this.returnValue=""I"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "break;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "fnSetName(sName);" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "fnSetTitle(sTitle);" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "fnSetIcon(sType,(sRFO!=""""));" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "fnSetHeader(sHeader);" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "fnSetMessage(sMessage,sExpression);" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "function fnSetName(sName)" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "document.all[""divName""].innerHTML=""Question: ""+sName;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "function fnSetTitle(sTitle)" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "document.all[""divTitle""].innerHTML=sTitle;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "function fnSetIcon(sType,bOverruled)" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "var sIcon="""";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "switch(sType)" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "case ""W"":sIcon=(bOverruled)?""ico_ok_warn.gif"":""ico_warn.gif"";break;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "case ""R"":sIcon=""ico_invalid.gif"";break;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "case ""I"":sIcon=""ico_inform.gif"";break;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "document.all[""divIcon""].innerHTML=""<img border='0' src='../../../img/""+sIcon+""'>"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "function fnSetHeader(sHeader)" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "document.all[""divHeader""].innerHTML=sHeader;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "function fnSetMessage(sMessage, sExpression)" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "document.all[""divMessage""].innerHTML=sMessage;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "function fnWarning()" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "var sHtm=""<table border='0'><tr height='15' class='clsLabelText'><td width='25%'>&nbsp;</td>"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "sHtm+=""<td width='75%'><div id='reasononoff'><a style='cursor:hand;' onclick='fnHideRFO();'><u>Remove overrule Reason</u></a></div></td></tr>"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "sHtm+=""<tr height='5'><td></td></tr>"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "sHtm+=""<tr id='RowRFOMess' height='30'><td colspan='4' class='clsLabelText'>""" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "sHtm+=""Please enter or choose the reason for overruling this warning"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "sHtm+=""</td></tr><tr id='RowRFO'>"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "sHtm+=""<td colspan='2'><input style='width:400px;' class='clsTextbox' name='txtInput' type='text'></td>"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "sHtm+=""</tr><tr id='RowRFOList'>"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "sHtm+=""<td colspan='2'><select style='width:400px;' name='selRFO' class='clsSelectList' size='3' onchange='javascript:SelectChange(\""selRFO\"",\""txtInput\"")'>"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "//add rfo" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "var aOverrule=sOverrule.split(sDel1);" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "if (aOverrule!=undefined)" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "for (var n=0;n<aOverrule.length;n++)" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "sHtm+=(aOverrule[n]!="""")?""<option value='""+aOverrule[n]+""'>""+aOverrule[n]+""</option>"":"""";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "sHtm+=""</select></td></tr></table>""" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "document.all[""divWarning""].innerHTML=sHtm;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "function fnClose()" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "switch (sType)" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "case ""W"":fnCheckRFO();break;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "case ""R"":fnReturn(""R"");break;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "case ""I"":fnReturn(""I"");break;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "function fnReturn(sVal)" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "this.returnValue=fnReplaceWithJSChars(sVal);" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "this.close();" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "function fnDisableOverrule()" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "document.all['reasononoff'].innerHTML='';" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "selRFO.disabled=true;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "txtInput.disabled=true;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "function fnHideRFO()" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "document.all['RowRFO'].style.visibility=""hidden"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "document.all['RowRFOList'].style.visibility=""hidden"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "document.all['RowRFOMess'].style.visibility=""hidden"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "document.all['reasononoff'].innerHTML=""<a style='cursor:hand;' onclick='fnShowRFO();'><u>Overrule this warning</u></a>"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "txtInput.value="""";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "bCollectRFO=false;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "function fnShowRFO()" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "document.all['RowRFO'].style.visibility=""visible"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "document.all['RowRFOList'].style.visibility=""visible"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "document.all['RowRFOMess'].style.visibility=""visible"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "document.all['reasononoff'].innerHTML=""<a style='cursor:hand;' onclick='fnHideRFO();'><u>Remove overrule Reason</u></a>"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "bCollectRFO=true;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "function fnCheckRFO()" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "if(bCollectRFO)" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "if(txtInput.value=="""")" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "alert('You must enter an overrule reason');" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "txtInput.focus();" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "else" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "fnReturn('O' + txtInput.value);" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "else" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "fnReturn('W');" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "function fnReplaceWithJSChars(sStr)" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "var sVal=sStr.replace(/\""/g, '\""');" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "sVal=sVal.replace(/\'/g, ""\'"");" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "sVal=sVal.replace(/\</g, ""&lt;"");" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "sVal=sVal.replace(/\>/g, ""&gt;"");" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "return sVal;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "</script>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "</head>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "<body onload='fnPageLoaded();'>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "<table align='center' width='95%' border='0'>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "<tr height='10'><td></td></tr>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "<tr height='15' class='clsLabelText'>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "<td colspan='2'><b><div id='divName'></div></b></div></td>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "<td><a style='cursor:hand;' onclick='fnClose();'><u>Close</u></a></td>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "</tr>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "<tr height='15' class='clsLabelText'><td bgcolor='#6699CC' colspan='4'><div id='divTitle'></div></td></tr>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "<tr height='15' class='clsLabelText'>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "<td width='20%'><div id='divIcon'></div></td>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "<td width='80%'><div id='divHeader'></div></td>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "</tr>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "<tr height='15' class='clsLabelText'>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "<td width='20%'>&nbsp;</td>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "<td width='80%'><div id='divMessage'></div></td>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "</tr>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "<tr><td colspan='2'>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "<div id='divWarning'></div>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "</td></tr>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "</table>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "</body>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "</html>" & vbCrLf)

    GetRFO = Join(vDialogHTML, "")
End Function

'-------------------------------------------------------------------------------------------'
Public Sub PublishStudy(ByVal lStudyCode As Long, _
                        ByVal nVersion As Integer)
'-------------------------------------------------------------------------------------------'
' ic 29/10/2002
' function creates non-dynamic study html files (question definitions, reason for change...)
' MLM 17/02/03: Generate eForm-specific files containing layout, category Q initialisation.
' ic 12/03/2003 added 'exit sub'
' ic 08/04/2003 uncommented rfo file create
' DPH 22/05/2003 - Reject Warn Inform dialog fully created for each study
'-------------------------------------------------------------------------------------------'
Dim oStudyDef As StudyDefRO
Dim oEf As eFormRO
Dim oCRFElement As eFormElementRO
'Dim oIo As MACROWWWIO30.clsWWW

Dim sDatabaseCnn As String
Dim sPath As String
Dim sFileName As String
Dim sDatabase As String
Dim nFile As Integer
Dim sQList As String
Dim sPList As String

    On Error GoTo ErrLabel

    'get user database
    sDatabase = goUser.Database.DatabaseCode

    'get database connection
    sDatabaseCnn = gsADOConnectString
    
    'Create a new StudyDef class and load study def
    Set oStudyDef = New StudyDefRO
    Call oStudyDef.Load(sDatabaseCnn, lStudyCode, nVersion)
    
    'get next free file number
    nFile = FreeFile
    
    'create directory path (asp\database\studyid\)
    sPath = gsWEB_HTML_LOCATION & "sites\" & sDatabase & "\" & oStudyDef.StudyId & "\"

    If (FolderExistence(sPath)) Then

        'Set oIo = New MACROWWWIO30.clsWWW

        'create rfc dialog
        sFileName = "rfc.html"
        'open the file for output
        Open sPath & sFileName For Output As #nFile
        'write to the file
        Print #nFile, GetRFC(oStudyDef)
        'close the file
        Close #nFile
        
        'ic 13/02/2003 rfo is now created 'on the fly'
        'ic 08/04/2003 not any more
        'create rfo dialog
        ' DPH 22/05/2003 - Reject Warn Inform dialog fully created for each study
        sFileName = "RejectWarnInform.htm"

        'open the file for output
        Open sPath & sFileName For Output As #nFile

        'write to the file
        Print #nFile, GetRFO(oStudyDef)

        'close the file
        Close #nFile

        'loop through each eform in the study
        For Each oEf In oStudyDef.eForms
            Call oStudyDef.LoadElements(oEf)
            
            'MLM 18/02/03: Output 2 files for each eform, for layout when used as main and visit eForm in web
            OutputEFormFile sPath, oStudyDef, oEf, eEFormUse.User
            OutputEFormFile sPath, oStudyDef, oEf, eEFormUse.VisitEForm
            
            'loop through each eform element on the eform
            For Each oCRFElement In oEf.EFormElements
                If (oCRFElement.QuestionId > 0) Then
                    'check 'already processed' list for questions to skip
                    If InStr(sQList, "|" & oCRFElement.QuestionId & "|") = 0 Then
                        sQList = sQList & "|" & oCRFElement.QuestionId & "|"
                    
                        sFileName = "qd" & oCRFElement.QuestionId & ".html"
                        
                        'open the file for output
                        Open sPath & sFileName For Output As #nFile
                        'write to the file
                        Print #nFile, GetQuestionDefinition(oCRFElement)
                        'close the file
                        Close #nFile
                        'Debug.Print "completed " & oCRFElement.Code & "(" & sFileName & ")"
                    Else
                        'Debug.Print "skipped " & oCRFElement.Code
                    End If
                Else
                    'copy images
                    If (oCRFElement.ControlType = ControlType.Picture) Then
                        If Dir(gsDOCUMENTS_PATH & oCRFElement.Caption) <> "" Then
                            'check we havent already copied this file by looking for the filename in our copied images variable
                            'which is just a comma delimited list of files we add to as we process this function
                            If InStr(sPList, "," & oCRFElement.Caption) = 0 Then
                
                                'if the file exists in the html folder already, delete it
                                If Dir(sPath & oCRFElement.Caption) <> "" Then Kill (sPath & oCRFElement.Caption)
                
                                'copy the file from source to destination
                                FileCopy gsDOCUMENTS_PATH & oCRFElement.Caption, sPath & oCRFElement.Caption
                
                                'add to our list of copied images
                                sPList = sPList & "," & oCRFElement.Caption
                                
                                Debug.Print "copied picture (" & oCRFElement.Caption & ")"
                            End If
                        Else
                            Debug.Print "picture not found (" & oCRFElement.Caption & ")"
                        End If
                    End If
                End If
            Next
            
            'remove elements from studydef to free memory
            Call oStudyDef.RemoveElements(oEf)
        Next
        
        
        'destroy object instances
        Set oEf = Nothing
    
    Else
        Call DialogError("Study files could not be created because " & sPath & " could not be created")
    End If
    
    Set oStudyDef = Nothing
    Exit Sub
    
ErrLabel:
    If MACROErrorHandler("modHTML", Err.Number, Err.Description, "PublishStudy", Err.Source) = Retry Then
        Resume
    End If
End Sub

'-------------------------------------------------------------------------------------------
Private Sub OutputEFormFile(ByRef sPath As String, oSD As StudyDefRO, _
                            oEform As eFormRO, enEFormUse As eEFormUse)
'-------------------------------------------------------------------------------------------
'
'-------------------------------------------------------------------------------------------
' REVISIONS
' DPH 02/05/2003 Give Button DIV an ID
' DPH 08/05/2003 Expand gap of next button
' DPH 12/05/2003 Place Next Button to full page width
' DPH 28/05/2003 Allow for Lab questions in header
' ic 05/06/2003 added return button
'-------------------------------------------------------------------------------------------

Dim lEFormWidth As Long
Dim nPageLength As Integer
Dim lHighestTabIndex As Long
Dim nFile As Integer
Dim oElement As eFormElementRO
Dim sURL As String
Dim lNextButtonLeft As Long
Dim bHasLabQuestions As Boolean

    On Error GoTo ErrLabel
    
    nFile = FreeFile
    
    lEFormWidth = oEform.eFormWidth
    If lEFormWidth = NULL_LONG Then
        lEFormWidth = 8515 ' Portrait width constant value
    End If
    
    ' DPH 28/05/2003 Allow for Lab questions in header
    bHasLabQuestions = EformHasLabQuestions(oEform)
    
    Open sPath & "form" & oEform.EFormId & "_" & enEFormUse & ".js" For Output As #nFile
    'Print #nFile, "document.writeln('";
    Print #nFile, "function DrawForm" & enEFormUse & "(){o3['Form" & enEFormUse & "'].outerHTML = '";
    
    'convert sPath into relative URL
    sURL = Replace(Replace(sPath, gsWEB_HTML_LOCATION, "../"), "\", "/")
    
    For Each oElement In oEform.EFormElements
        Print #nFile, JSStringLiteral(RtnEformElement(oSD.FontColour, _
            oSD.FontSize, _
            oSD.FontName, _
            oEform, _
            oElement, _
            sURL, _
            nPageLength, _
            lEFormWidth, _
            lHighestTabIndex, _
            enEFormUse, _
            bHasLabQuestions));
    Next oElement
    
    'if this is the main eform, include HTML for the Next button
    ' DPH 16/10/2002 - Changed to use oLastFocusID
    ' DPH 26/11/2002 - Tab No to last on page
    ' DPH 27/11/2002 - Corrected setFieldFocus call
    If enEFormUse = eEFormUse.User Then
        ' DPH 12/05/2003 Place Next button to left of edge of page
        lNextButtonLeft = (lEFormWidth / nIeXSCALE) - 32 ' 32 being pixel size of next button image
        If lNextButtonLeft < 600 Then
            lNextButtonLeft = 600
        End If
        'ic 05/06/2003 added return button
        ' DPH 02/05/2003 Give Button DIV an ID
        ' DPH 08/05/2003 Expand gap of next button
        Print #nFile, JSStringLiteral("<div style='position:absolute; left:" & lNextButtonLeft & "; top:" & nPageLength + 80 & ";z-index:4;' id='btnNextDIV'>" & _
            "<a onfocus='if(o1.oLastFocusID!=null){if(!o1.fnLostFocus(o1.oLastFocusID)){o1.setFieldFocus(o1.oLastFocusID.name,o1.DefaultRepeatNo(o1.oLastFocusID.idx));}else{o1.fnOnBtnNext();}}' tabindex='" & _
            (lHighestTabIndex + mnTABINDEXGAP) & "' name='btnNext' type='button' value=' > ' href='javascript:fnSave(""m2"")'>" & _
            "<img alt='Save and move to next eForm' border='0' src='../img/ico_nexteform.gif'></a><br>" & _
            "<a onfocus='if(o1.oLastFocusID!=null){if(!o1.fnLostFocus(o1.oLastFocusID)){o1.setFieldFocus(o1.oLastFocusID.name,o1.DefaultRepeatNo(o1.oLastFocusID.idx));}else{o1.fnOnBtnNext();}}' tabindex='" & _
            (lHighestTabIndex + (mnTABINDEXGAP * 2)) & "' name='btnBack' type='button' value=' > ' href='javascript:fnBack();'>" & _
            "<img alt='Return' border='0' src='../img/ico_previouson.gif'></a></div>");
    End If
    
    ' DPH 19/05/2003 - Set up an eForm DIV for the color
    If enEFormUse = eEFormUse.User Then
        If bHasLabQuestions Then
            Print #nFile, JSStringLiteral("<div id=""eFormColorDIV"" style=""top:" & (1100 / nIeYSCALE) & ";left:0;width:" & (lEFormWidth / nIeXSCALE) & ";height:" & (nPageLength + 115) & ";background-color:" & RtnHTMLCol(oEform.BackgroundColour) & ";z-index:0;""></div>");
        Else
            Print #nFile, JSStringLiteral("<div id=""eFormColorDIV"" style=""top:" & (750 / nIeYSCALE) & ";left:0;width:" & (lEFormWidth / nIeXSCALE) & ";height:" & (nPageLength + 115) & ";background-color:" & RtnHTMLCol(oEform.BackgroundColour) & ";z-index:0;""></div>");
        End If
    End If
    Print #nFile, "';"
    Print #nFile, RtnRadioCreateFuncs(oEform)
    Print #nFile, RtnEformCategoryTxtFunc(oEform)
    Print #nFile, "}"
    Close #nFile
    
    Exit Sub

ErrLabel:
    If MACROErrorHandler("modHTML", Err.Number, Err.Description, "OutputEFormFile", Err.Source) = Retry Then
        Resume
    End If

End Sub
'-------------------------------------------------------------------------------------------'
Private Function RtnEformElement(ByVal lStudyFontCol As Long, _
                                 ByVal nStudyFontSize As Integer, _
                                 ByVal sStudyFontName As String, _
                                 ByRef oEform As eFormRO, _
                                 ByRef oCRFElement As eFormElementRO, _
                                 ByVal sImagePath As String, _
                                 ByRef nPageLength As Integer, _
                                 ByVal lEFormWidth As Long, _
                                 ByRef lHighestTabIndex As Long, _
                                 Optional ByRef enEFormUse As eEFormUse = eEFormUse.User, _
                                 Optional ByVal bHasLabQuestions As Boolean = False) As String
'-------------------------------------------------------------------------------------------'
' function accepts a studydef object, eform element object, form bg colour
' function returns a string representing the passed element as vbscript and html: rewrite
' for v3.0 using business objects
'-------------------------------------------------------------------------------------------'
'Revisions:
'ic 28/02/2002 changed user eform picture path to /img/eForm/
'ZA 30/07/2002 added style sheet for font/colour for question/caption/comment
'ic 03/10/2002 moved from modhtml, arguments changed
'MLM 10/10/02: Added Optional enEFormUse argument, and used to hide elements on the visit form (except for the visit date)
'DPH 15/10/2002 Changed fnGotFocus, fnLostFocus calls to use 'this' in the parameter
'DPH 18/10/2002 Added RQG objects
'DPH 29/10/2002 - Added default Repeat No (zero) to fnPopMenu Call as RQG do not get drawn here
'DPH 07/11/2002 Added RQG Caption / changed RQG Position/Sizing
'ic  22/11/2002 added image path arguement
'ic  22/11/2002 added hotlink controltype
'DPH 25/11/2002 Max height for RQG
'DPH 26/11/2002 Space TabIndex by mnTABINDEXGAP to allow for RQG inbetween
' Formula for tabindex val = (CRFElement.ElementOrder - 1) * mnTABINDEXGAP
'ic 14/01/2003 start tabindex at 200 to avoid clashes with visit eform tabindex
' MLM Copied from modUIHTMLEForm and changed to use an EForm instead of an EForm Instance
' DPH 01/05/2003 - Align status icon / sdv cell to top
' IC/DPH 02/05/2003 - Change in pixel coordinates visit date
' DPH 08/05/2003 - bug 1618 Change Y offset to 750 to allow for eForm header
' DPH 13/05/2003 - For comments need to use font 'style' font-family & font-size
' DPH 20/05/2003 - All eForm components have a z-index > 0
' DPH 28/05/2003 - allow for lab questions in header
' DPH 04/09/2003 - Create tab order for 'zero' order elements
' ic 15/02/2005 bug 2529, only add active categories
' ic 27/05/2005 issue 2579, dont add hidden fields for comments/lines/pictures
' DPH 29/06/2005 - issue 2412 keep note/comment icon in text box view when display length shorter than
'                   actual length of field
'-------------------------------------------------------------------------------------------'

    Dim oPic As Picture
    Dim oCategory As CategoryItem
    Dim oValidation As Validation
    Dim oRQG As QGroupRO
    Dim oOwnerRQG As QGroupRO
    
    Dim sRtn As String
    Dim sCpdImages As String
    Dim nLength As Integer
    Dim nWidth As Integer
    Dim bDefaultFont As Boolean
    Dim bDefaultCaptionFont As Boolean
    Dim sStyle As String
    Dim sStyleRadio As String
    'MLM 26/09/02:
    Dim lElementX As Long
    Dim lElementY As Long
    ' DPH 18/10/2002
    Dim bRQG As Boolean
    Dim bBelongsToRQG As Boolean
    ' DPH 14/11/2002
    Dim lRQGWidth As Long
    Dim sEformLink As String
    Dim lRQGMaxHeight As Long
    Dim lRQGLowPoint As Long
    ' DPH 26/11/2002
    Dim lTabIndex As Long
    ' DPH 08/05/2003 - offset for eform header
    Dim nEFormHeadOffset As Integer
    ' DPH 29/06/2005 - string for onscroll event for text boxes
    Dim sTxtScrollEvent As String
    
    ' offset number for header
    ' DPH 28/05/2003 - allow for lab questions in header
    If bHasLabQuestions Then
        nEFormHeadOffset = 1100
    Else
        nEFormHeadOffset = 750
    End If
    
    ' DPH 18/10/2002 - Retrieve RQG element information
    Set oRQG = oCRFElement.QGroup
    Set oOwnerRQG = oCRFElement.OwnerQGroup
    If oRQG Is Nothing Then
        bRQG = False
    Else
        bRQG = True
    End If
    If oOwnerRQG Is Nothing Then
        bBelongsToRQG = False
    Else
        bBelongsToRQG = True
    End If
    sRtn = ""
    
    'ZA - 23/07/2002 - check font attributes for caption/comment
    If oCRFElement.Caption <> "" Then
        If ((oCRFElement.CaptionFontColour = lStudyFontCol) _
        And (oCRFElement.CaptionFontSize = nStudyFontSize) _
        And (oCRFElement.CaptionFontName = sStudyFontName)) _
        Or ((oCRFElement.CaptionFontColour = lStudyFontCol) _
        And (oCRFElement.CaptionFontSize = nStudyFontSize) _
        And (oCRFElement.CaptionFontName = sStudyFontName)) Then
            bDefaultCaptionFont = True
        Else
            bDefaultCaptionFont = False
        End If
    End If
    
    'do this element's font attributes match the study default font attributes, or are empty
    If ((oCRFElement.FontColour = lStudyFontCol) _
    And (oCRFElement.FontSize = nStudyFontSize) _
    And (oCRFElement.FontName = sStudyFontName)) _
    Or ((oCRFElement.FontColour = lStudyFontCol) _
    And (oCRFElement.FontSize = nStudyFontSize) _
    And (oCRFElement.FontName = sStudyFontName)) Then
        bDefaultFont = True
    Else
        bDefaultFont = False
    End If
    
    ' Special case for RQG or Question that is part of a RQG
    If bRQG Or bBelongsToRQG Then
        ' ignore questions belonging to a question group as dealt with completely in Javascript
        ' DPH 21/11/2002 - Added RQG Max height function
        lRQGMaxHeight = 0
        lRQGLowPoint = 0
        If bRQG Then
            ' put in function to tidy
            If oCRFElement.ElementUse = eElementUse.EFormVisitDate Then
                lElementY = 0 'in the header bar
                If enEFormUse = eEFormUse.User Then
                    'form date; position on the right
                    lElementX = 6386 / nIeXSCALE
                Else
                    'visit date; position on the left
                    lElementX = 2129 / nIeXSCALE
                End If
            Else
                ' DPH 08/05/2003 - Change Y offset to 750 to allow for eForm header
                '"normal" elements are moved down a bit
                lElementY = (oCRFElement.ElementY + nEFormHeadOffset) / nIeYSCALE
                lElementX = oCRFElement.ElementX / nIeXSCALE
            End If
    
            ' Calculate width
            lRQGWidth = (lEFormWidth / nIeXSCALE) - lElementX
            ' Calculate max height
            lRQGMaxHeight = CalculateRQGMaxHeight(oEform, oCRFElement)
            ' Calculate For PageLength variable
            lRQGLowPoint = (lElementY + 24) + lRQGMaxHeight
            
            'outer div enclosing RQG headers
            sRtn = sRtn & "<div id=" _
                        & "g_" & oRQG.Code & "_RQGHeadDiv " _
                        & "style=" & Chr(34) & "position:absolute; overflow: hidden;left:" & lElementX & ";" _
                        & "top:" & lElementY & ";" _
                        & "height:" & "25" & ";" _
                        & "width:" & lRQGWidth & ";" _
                        & "z-index:3;" & Chr(34) & "></div>" & vbCrLf
            
            'outer div enclosing all RQG questions
            sRtn = sRtn & "<div id=" _
                        & "g_" & oRQG.Code & "_RQGDiv " _
                        & "style=" & Chr(34) & "position:absolute; overflow: auto;left:" & lElementX & ";" _
                        & "top:" & (lElementY + 24) & ";" _
                        & "height:" & lRQGMaxHeight & ";" _
                        & "width:" & lRQGWidth & ";" _
                        & "z-index:3;" & Chr(34) & "></div>" & vbCrLf
        
            If oCRFElement.Caption <> "" Then
                sRtn = sRtn & RtnCaption(oCRFElement, bDefaultCaptionFont, oEform.displayNumbers, lStudyFontCol, nEFormHeadOffset)
            End If
        
        End If
        
        RtnEformElement = sRtn
        ' DPH 21/11/2002 - Adjust Pagelength
        If lRQGLowPoint > 0 And lRQGLowPoint > nPageLength Then
            nPageLength = lRQGLowPoint
        End If

        ' DPH 26/11/2002 - TabIndex
        If oCRFElement.ElementOrder > 1 Then
            lTabIndex = CLng(oCRFElement.ElementOrder - 1) * mnTABINDEXGAP
        Else
            lTabIndex = 1
        End If
        If lHighestTabIndex < lTabIndex Then
            lHighestTabIndex = lTabIndex
        End If

        Exit Function
    End If
    
    ' DPH 08/05/2003 - Add offset to pagelength (750) ...
    'keep track of the lowest element on eform so 'next' button can be positioned
    If CLng((oCRFElement.ElementY + nEFormHeadOffset) / nIeYSCALE) > nPageLength Then
        nPageLength = CLng((oCRFElement.ElementY + nEFormHeadOffset) / nIeYSCALE)
    End If

    If enEFormUse = eEFormUse.User Or (enEFormUse = eEFormUse.VisitEForm And oCRFElement.QuestionId > 0) Then
    'controls are either hidden or visible
    'MLM 11/10/02:
    If oCRFElement.Hidden Or (enEFormUse = eEFormUse.VisitEForm And oCRFElement.ElementUse <> eElementUse.EFormVisitDate) Then
        
        'dont add hidden fields for lines/comments/pictures. these may be in the database due to
        'ggb mis-translating
        If ((oCRFElement.ControlType > 0) And (oCRFElement.ControlType < 10000)) Then
            'hidden controls are always represented as html hidden fields
            sRtn = sRtn & "<input " _
                        & "type=" & Chr(34) & "hidden" & Chr(34) & " " _
                        & "name=" & Chr(34) & oCRFElement.WebId & Chr(34) _
                        & ">"
            
            'ic 24/04/2002 add add info field - needed for hidden fields
            sRtn = sRtn & "<input " _
                            & "type=" & Chr(34) & "hidden" & Chr(34) & " " _
                            & "name=""a" & oCRFElement.WebId & Chr(34) & ">"
        End If
                       
    Else
        'visible controls can be comments,lines,textboxes,attachments,radio buttons,pictures,drop-down lists
        
        'MLM 26/09/02: Work out the pixel coordinates of the element here
        If oCRFElement.ElementUse = eElementUse.EFormVisitDate Then
            ' IC/DPH 02/05/2003 - Change in pixel coordinates visit date
            lElementY = 27 'in the header bar
            If enEFormUse = eEFormUse.User Then
                'form date; position on the right
                lElementX = 6386 / nIeXSCALE
            Else
                'visit date; position on the left
                lElementX = 2129 / nIeXSCALE
            End If
        Else
            '"normal" elements are moved down a bit
            ' DPH 08/05/2003 - Change Y offset to 750 to allow for eForm header
            lElementY = (oCRFElement.ElementY + nEFormHeadOffset) / nIeYSCALE
            lElementX = oCRFElement.ElementX / nIeXSCALE
        End If
        
        'ic 14/01/2003 start tabindex at 200 to avoid clashes with visit eform tabindex
        If (enEFormUse = eEFormUse.User) Then
            ' DPH 04/09/2003 - Create tab order for 'zero' order elements
            If (oCRFElement.ElementOrder > 0) Then
                lTabIndex = CLng(oCRFElement.ElementOrder) * mnTABINDEXGAP
            Else
                ' if Zero can be eform date or elements that do not have a tab order
                lTabIndex = 100
            End If
        Else
            lTabIndex = oCRFElement.ElementOrder
        End If
        
        If lHighestTabIndex < lTabIndex Then
            lHighestTabIndex = lTabIndex
        End If
        
        Select Case oCRFElement.ControlType
        Case ControlType.Comment
            'comment
            If oCRFElement.Caption <> "" Then
                sRtn = sRtn & "<div style=" & Chr(34) _
                            & "top:" & lElementY & ";" _
                            & "left:" & lElementX & ";" _
                            & "z-index:1" & Chr(34) & ">"
                'ZA 22/07/2002 - changed FontName with CaptionFontName for CRFEelement
                If Not bDefaultCaptionFont Then
                    'DPH 13/05/2003 - For comments need to use font style font-family & font-size
                    sRtn = sRtn & "<font style="""
                    If (oCRFElement.CaptionFontColour <> 0) Then sRtn = sRtn & "color:" & RtnHTMLCol(oCRFElement.CaptionFontColour) & "; "
                    If (oCRFElement.CaptionFontSize <> 0) Then sRtn = sRtn & "FONT-SIZE:" & oCRFElement.CaptionFontSize & " pt; "
                    If (oCRFElement.CaptionFontName <> "") Then sRtn = sRtn & "FONT-FAMILY:" & oCRFElement.CaptionFontName & "; "
                    sRtn = sRtn & """>"
                End If
                
                If oCRFElement.CaptionFontBold Then sRtn = sRtn & "<b>"
                If oCRFElement.CaptionFontItalic Then sRtn = sRtn & "<i>"
                sRtn = sRtn & Replace(oCRFElement.Caption, vbCrLf, "<br>")
                If oCRFElement.CaptionFontItalic Then sRtn = sRtn & "</i>"
                If oCRFElement.CaptionFontBold Then sRtn = sRtn & "</b>"
                
                If Not bDefaultFont Then
                    sRtn = sRtn & "</font>"
                End If
                
                sRtn = sRtn & "</div>"
            End If

        Case ControlType.Line
            'line
            sRtn = sRtn & "<div " _
                        & "style=" & Chr(34) & "20px;left:0;top:" & lElementY & ";z-index:1;" & Chr(34) _
                        & ">" _
                        & "<hr>" _
                        & "</div>"
            
        
        Case ControlType.TextBox, ControlType.Calendar, ControlType.RichTextBox
            'textbox
            'MLM 26/09/02: Don't draw captions for form and visit dates.
            'If oCRFElement.Caption <> "" And oCRFElement.ElementUse = eElementUse.User Then
                sRtn = sRtn & RtnCaption(oCRFElement, bDefaultCaptionFont, oEform.displayNumbers, lStudyFontCol, nEFormHeadOffset, , enEFormUse)
            'End If
        
        
            SetElementWidth oCRFElement.Format, oCRFElement.QuestionLength, oCRFElement.DisplayLength, nLength, nWidth
            
        
            sRtn = sRtn & "<div id=" & Chr(34) & oCRFElement.WebId & "_inpDiv" & Chr(34) & " " _
                        & "style=" & Chr(34) & ";left:" & lElementX _
                        & ";top:" & lElementY _
                        & ";z-index:2" & Chr(34) & ">"
            
            
            
            sRtn = sRtn & "<input " _
                        & "type=" & Chr(34) & "hidden" & Chr(34) & " " _
                        & "name=""a" & oCRFElement.WebId & Chr(34) & ">"
            
            'set element style
            sStyle = Chr(34) & "background-color:#EEEEEE; "
            'MLM 14/10/02: Include in-line style to change appearance of form and visit dates
            If (oCRFElement.ElementUse = eElementUse.EFormVisitDate) Or Not bDefaultFont Then
                sStyle = sStyle & SetFontAttributes(oCRFElement, True)
            End If
            sStyle = sStyle & Chr(34)
            
            ' DPH 29/06/2005 - issue 2412 add scroll event to keep note/comment in view
            If (oCRFElement.ControlType = ControlType.TextBox) Then
                sTxtScrollEvent = " onscroll=" & Chr(34) & "o1.fnDisplayNoteStatusScroll(this);" & Chr(34)
                sTxtScrollEvent = sTxtScrollEvent & " onchange=" & Chr(34) & "o1.fnDisplayNoteStatusScroll(this);" & Chr(34)
            Else
                sTxtScrollEvent = ""
            End If
            
            sRtn = sRtn & "<table cellpadding='0' cellspacing='0'><tr><td onMouseOver=" & Chr(34) _
                        & "fnSetTooltip('" & oCRFElement.WebId & "',0,1);" & Chr(34) _
                        & "><input type=" & Chr(34) & "text" & Chr(34) & " " _
                        & "disabled " _
                        & "tabindex=" & Chr(34) & lTabIndex & Chr(34) & " " _
                        & "name=" & Chr(34) & oCRFElement.WebId & Chr(34) & " " _
                        & "Style=" & sStyle & " " _
                        & "size=" & Chr(34) & nWidth & Chr(34) & " " _
                        & "maxlength=" & Chr(34) & nLength & Chr(34) & " " _
                        & "onfocus=" & Chr(34) & "o1.fnGotFocus(this);" & Chr(34) & sTxtScrollEvent & ">"
            
                        '& "o1.fnDisplayFieldTxt(" & "deFrm.af_" & LCase(oCRFElement.Code) & ",'" & oCRFElement.Name & sBlockDelimeter & oCRFElement.Format & sBlockDelimeter & oCRFElement.QuestionLength & sBlockDelimeter & Replace(oCRFElement.Helptext, "'", "\'") & "');" & Chr(34) & ">"


            ' DPH 29/10/2002 - Added default Repeat No (zero) to fnPopMenu Call
            sRtn = sRtn & "</td><a onMouseup='javascript:fnPopMenu(" & Chr(34) & oCRFElement.WebId & Chr(34) & ",0,event);'><td><img src='../img/blank.gif' " _
                        & "name='imgc_" & oCRFElement.WebId & "'></td><td>"
                        
            sRtn = sRtn & "<table cellpadding='0' cellspacing='0'><tr><td><img src='../img/blank.gif' " _
                        & "name='img_" & oCRFElement.WebId & "' " & "onMouseOver=" & Chr(34) & "fnSetTooltip('" & oCRFElement.WebId & "',0,2);" & Chr(34) _
                        & "></td></tr><tr><td><img src='../img/blank.gif' " _
                        & "name='imgs_" & oCRFElement.WebId & "'></td></tr></table></td></a>" & GetNRCTCHTML(oCRFElement) & "</tr></table>"
            
            sRtn = sRtn & "</div>"
        
        Case ControlType.Attachment
            'multimedia
            
            If oCRFElement.Caption <> "" Then
                sRtn = sRtn & RtnCaption(oCRFElement, bDefaultCaptionFont, oEform.displayNumbers, lStudyFontCol, nEFormHeadOffset)
            End If
            
            sRtn = sRtn & "<input " _
                        & "type=" & Chr(34) & "hidden" & Chr(34) & " " _
                        & "name=""a" & oCRFElement.WebId & Chr(34) & ">"
            
            sRtn = sRtn & "<div id=" & Chr(34) & oCRFElement.WebId & "_inpDiv" & Chr(34) & " " _
                        & "style=" & Chr(34) & "left:" & lElementX _
                        & ";top:" & lElementY _
                        & ";z-index:2" & Chr(34) & ">"

            sRtn = sRtn & "<table cellpadding='0' cellspacing='0'><tr><td><input type=" & Chr(34) & "file" & Chr(34) & " " _
                        & "disabled " _
                        & "tabindex=" & Chr(34) & lTabIndex & Chr(34) & " " _
                        & "name=" & Chr(34) & oCRFElement.WebId & Chr(34) & " " _
                        & "style=" & Chr(34) & "background-color:#EEEEEE; " & Chr(34) & " " _
                        & "onfocus=" & Chr(34) & "o1.fnGotFocus(this);" & Chr(34) _
                        & " onMouseOver=" & Chr(34) & "fnSetTooltip('" & oCRFElement.WebId & "',0,1);" & Chr(34) & ">"
                        
                        '& "o1.fnDisplayFieldTxt(" & "deFrm.af_" & LCase(oCRFElement.Code) & ",'" & oCRFElement.Name & sBlockDelimeter & oCRFElement.Format & sBlockDelimeter & oCRFElement.QuestionLength & sBlockDelimeter & Replace(oCRFElement.Helptext, "'", "\'") & "');" & Chr(34) & ">"


            ' DPH 29/10/2002 - Added default Repeat No (zero) to fnPopMenu Call
            sRtn = sRtn & "</td><a onMouseup='javascript:fnPopMenu(" & Chr(34) & oCRFElement.WebId & Chr(34) & ",0,event);'><td><img src='../img/blank.gif' " _
                         & "name='imgc_" & oCRFElement.WebId & "'></td><td>"
                         
            sRtn = sRtn & "<table cellpadding='0' cellspacing='0'><tr><td><img src='../img/blank.gif' " _
                         & "name='img_" & oCRFElement.WebId & "'" & " onMouseOver=" & Chr(34) & "fnSetTooltip('" & oCRFElement.WebId & "',0,2);" & Chr(34) _
                         & "></td></tr><tr><td><img src='../img/blank.gif' " _
                         & "name='imgs_" & oCRFElement.WebId & "'></td></tr></table></td></a>" & GetNRCTCHTML(oCRFElement) & "</tr></table>"
            
            sRtn = sRtn & "</div>"


        Case ControlType.OptionButtons, ControlType.PushButtons
            'radio button group (1 or more radio buttons)
           '''''''''''''''''''''''''''''''''''''''''''''''''' sStyleRadio = "Style = """
            'outer div enclosing all radio buttons and captions in the group
            sRtn = sRtn & "<div " _
                        & "style=" & Chr(34) & ";left:" & lElementX & ";" _
                        & "top:" & lElementY & ";" _
                        & "z-index:2;" & Chr(34) & ">" & vbCrLf
            
            sRtn = sRtn & "<input " _
                        & "type=" & Chr(34) & "hidden" & Chr(34) & " " _
                        & "name=""a" & oCRFElement.WebId & Chr(34) & ">" & vbCrLf
                        
            sRtn = sRtn & "<input " _
                        & "type=" & Chr(34) & "hidden" & Chr(34) & " " _
                        & "name=" & Chr(34) & oCRFElement.WebId & Chr(34) & ">" & vbCrLf
            
'''''''            If Not bDefaultFont Then
'''''''                If (oCRFElement.FontColour <> 0) Then sStyleRadio = sStyleRadio & "COLOR: " & RtnHTMLCol(oCRFElement.FontColour) & ";"
'''''''                If (oCRFElement.FontSize <> 0) Then sStyleRadio = sStyleRadio & "FONT-SIZE: " & oCRFElement.FontSize & " pt;"
'''''''                If (oCRFElement.FontName <> "") Then sStyleRadio = sStyleRadio & "FONT-FAMILY: " & oCRFElement.FontName & ";"
'''''''                If oCRFElement.FontBold Then sStyleRadio = sStyleRadio & "FONT-WEIGHT: bold; "
'''''''                If oCRFElement.FontItalic Then sStyleRadio = sStyleRadio & "FONT-STYLE: italic;" & Chr(34)
'''''''            End If
            
'''''''            If Not bDefaultFont Then
'''''''                sStyleRadio = """" & SetFontAttributes(oCRFElement) & """"
'''''''            End If
            
            ''''''''''''''sRtn = sRtn & "<table><tr><td Style = " & sStyleRadio & ">"
            sRtn = sRtn & "<table cellpadding='0' cellspacing='0'><tr><td>"
            
            'div positioning radio buttons
            sRtn = sRtn & "<div id=" & Chr(34) & oCRFElement.WebId & "_inpDiv" & Chr(34) & " " _
                        & "style=" & Chr(34) & "position:relative;z-index:2;"
                        
            sRtn = sRtn & Chr(34) & ">"

            sRtn = sRtn & "</div></td><td valign=" & Chr(34) & "top" & Chr(34) & ">"
            
            sRtn = sRtn & "<img src='../img/blank.gif' " _
                         & "name='imgc_" & oCRFElement.WebId & "'></td>"
                         
            ' DPH 01/05/2003 - Align status icon / sdv cell to top
            sRtn = sRtn & "<td valign=" & Chr(34) & "top" & Chr(34) & ">"
            
            ' DPH 29/10/2002 - Added default Repeat No (zero) to fnPopMenu Call
            sRtn = sRtn & "<table cellpadding='0' cellspacing='0'><tr><a onMouseup='javascript:fnPopMenu(" & Chr(34) & oCRFElement.WebId & Chr(34) & ",0,event);'><td><img src='../img/blank.gif' " _
                         & "name='img_" & oCRFElement.WebId & "'" & " onMouseOver=" & Chr(34) & "fnSetTooltip('" & oCRFElement.WebId & "',0,2);" & Chr(34) _
                         & "></td></a><tr><tr><td><img src='../img/blank.gif' " _
                         & "name='imgs_" & oCRFElement.WebId & "'></td></tr></table>"
            
            sRtn = sRtn & "</td>" & GetNRCTCHTML(oCRFElement) & "</tr></table></div>" & vbCrLf
            
            
            
            If oCRFElement.Caption <> "" Then
                sRtn = sRtn & RtnCaption(oCRFElement, bDefaultCaptionFont, oEform.displayNumbers, lStudyFontCol, nEFormHeadOffset)
            End If

        
            
        Case ControlType.Picture
            'picture
            
        'if the element  cannot be found it is omitted
'        If Dir(msIMAGE_PATH & oCRFElement.Caption) <> "" Then
'            Set oPic = LoadPicture(App.Path & "\" & msIMAGE_PATH & oCRFElement.Caption)
'            nWidth = frmGenerateHTML.ScaleY(oPic.Width, vbHimetric, vbPixels) * nIeXSCALE
'            Set oPic = Nothing

            If nWidth > 7000 And nWidth < 5000 Then 'Assume it is a line, limit to width of page
                
                nLength = 611
                sRtn = sRtn & "<div id=" & Chr(34) & "lin" & oCRFElement.ElementID & "_inpDiv" & Chr(34) & " " _
                            & "style=" & Chr(34) & ";left:" & lElementX & ";" _
                            & "top:" & lElementY & ";z-index:1;" & Chr(34) & ">"
                
                sRtn = sRtn & "<img src=" & Chr(34) & sImagePath & oCRFElement.Caption & Chr(34) & " " _
                            & "width=" & Chr(34) & nLength & Chr(34) & " " _
                            & "valign=" & Chr(34) & "top" & Chr(34) & " " _
                            & "align=" & Chr(34) & "left" & Chr(34) & ">"
                sRtn = sRtn & "</div>"

            Else
                sRtn = sRtn & "<div id=" & Chr(34) & "pic" & oCRFElement.ElementID & "_inpDiv" & Chr(34) & " " _
                            & "style=" & Chr(34) & ";left:" & lElementX & ";" _
                            & "top:" & lElementY & ";z-index:1;" & Chr(34) & ">"
                
                sRtn = sRtn & "<img src=" & Chr(34) & sImagePath & oCRFElement.Caption & Chr(34) & " " _
                            & "valign=" & Chr(34) & "top" & Chr(34) & " " _
                            & "align=" & Chr(34) & "left" & Chr(34) & ">"
                sRtn = sRtn & "</div>"
            End If
            
        Case ControlType.PopUp
            'drop-down list
            

            sRtn = sRtn & "<div id=" & Chr(34) & oCRFElement.WebId & "_inpDiv" & Chr(34) & " " _
                        & "style=" & Chr(34) & ";left:" & lElementX & ";" _
                        & "top:" & lElementY & ";" _
                        & "z-index:2" & Chr(34) & ">"
            
            
            sRtn = sRtn & "<input " _
                        & "type=" & Chr(34) & "hidden" & Chr(34) & " " _
                        & "name=""a" & oCRFElement.WebId & Chr(34) & ">"
            
            'set element style
            sStyle = Chr(34) & "background-color:#EEEEEE; "
            If Not bDefaultFont Then
                sStyle = sStyle & SetFontAttributes(oCRFElement, True)
            End If
            sStyle = sStyle & Chr(34)


            sRtn = sRtn & "<table><tr>" _
                        & "<td>"

            'onblur event so that answers are evaluated when tabbing between fields as well as clicking
            sRtn = sRtn & "<select tabindex=" & Chr(34) & lTabIndex & Chr(34) & " " _
                            & "disabled " _
                            & "name=" & Chr(34) & oCRFElement.WebId & Chr(34) & " " _
                            & "Style=" & sStyle & " " _
                            & "onfocus=" & Chr(34) & "o1.fnGotFocus(this);" & Chr(34) & " " _
                            & "onchange=" & Chr(34) & "o1.fnLostFocus(this);" & Chr(34) _
                            & " onMouseOver=" & Chr(34) & "fnSetTooltip('" & oCRFElement.WebId & "',0,1);" & Chr(34) & ">"
                            
                            '& "o1.fnDisplayFieldTxt(" & "deFrm.af_" & LCase(oCRFElement.Code) & ",'" & oCRFElement.Name & sBlockDelimeter & oCRFElement.Format & sBlockDelimeter & sBlockDelimeter & Replace(oCRFElement.Helptext, "'", "\'") & "');" & Chr(34) & " " _

            'space as default in pop-up box
            sRtn = sRtn & "<option value=" & Chr(34) & Chr(34) & "> </option>"

            
            For Each oCategory In oCRFElement.Categories
                'ic 15/02/2005 bug 2529, only add active categories
                If (oCategory.Active) Then
                    sRtn = sRtn & "<option value=" & Chr(34) & oCategory.Code & Chr(34) & ">" _
                            & oCategory.Value _
                            & "</option>"
                End If
            Next
            


        sRtn = sRtn & "</select>"
        
        ' DPH 29/10/2002 - Added default Repeat No (zero) to fnPopMenu Call
        sRtn = sRtn & "</td><a onMouseup='javascript:fnPopMenu(" & Chr(34) & oCRFElement.WebId & Chr(34) & ",0,event);'><td>"
        
        sRtn = sRtn & "<img src='../img/blank.gif' " _
                         & "name='imgn_" & oCRFElement.WebId & "'>"

        sRtn = sRtn & "<img src='../img/blank.gif' " _
                         & "name='imgc_" & oCRFElement.WebId & "'></td><td>"
                         
        sRtn = sRtn & "<table cellpadding='0' cellspacing='0'><tr><td><img src='../img/blank.gif' " _
                         & "name='img_" & oCRFElement.WebId & "'" & " onMouseOver=" & Chr(34) & "fnSetTooltip('" & oCRFElement.WebId & "',0,2);" & Chr(34) _
                         & "></td></tr><tr><td><img src='../img/blank.gif' " _
                         & "name='imgs_" & oCRFElement.WebId & "'></td></tr></table>"
        
        sRtn = sRtn & "</td></a>" & GetNRCTCHTML(oCRFElement) & "</tr></table>"

        sRtn = sRtn & "</div>"

        If oCRFElement.Caption <> "" Then
            sRtn = sRtn & RtnCaption(oCRFElement, bDefaultCaptionFont, oEform.displayNumbers, lStudyFontCol, nEFormHeadOffset)
        End If
        
        Case ControlType.Hotlink
            'hotlink
            'MLM 19/02/03: Rather than evaluating the target of the hotlink here, send the hotlink.
            '   This will be intercepted using new h prefix and evaluated in SaveAndLoadEForm.
            If oCRFElement.Caption <> "" Then
'                sEformLink = CStr(oEFI.GetHotlinkTarget(oCRFElement.Hotlink))
            
                sRtn = sRtn & "<div style=" & Chr(34) _
                            & "top:" & lElementY & ";" _
                            & "left:" & lElementX & ";" _
                            & "z-index:1" & Chr(34) & ">"
                'ZA 22/07/2002 - changed FontName with CaptionFontName for CRFEelement
                If Not bDefaultCaptionFont Then
                    sRtn = sRtn & "<font "
                    If (oCRFElement.CaptionFontColour <> 0) Then sRtn = sRtn & "color=" & Chr(34) & RtnHTMLCol(oCRFElement.CaptionFontColour) & Chr(34) & " "
                    If (oCRFElement.CaptionFontSize <> 0) Then sRtn = sRtn & "size=" & Chr(34) & RtnFontSize(oCRFElement.CaptionFontSize) & Chr(34) & " "
                    If (oCRFElement.CaptionFontName <> "") Then sRtn = sRtn & "face=" & Chr(34) & oCRFElement.CaptionFontName & Chr(34)
                    sRtn = sRtn & ">"
                End If
                
                If oCRFElement.CaptionFontBold Then sRtn = sRtn & "<b>"
                If oCRFElement.CaptionFontItalic Then sRtn = sRtn & "<i>"
                
                If (sEformLink <> "0") Then
                    sRtn = sRtn & "<a href='javascript:fnSave(" & Chr(34) & "h" & oCRFElement.Hotlink & Chr(34) & ")'>"
                    sRtn = sRtn & Replace(oCRFElement.Caption, vbCrLf, "<br>")
                    sRtn = sRtn & "</a>"
                Else
                    sRtn = sRtn & Replace(oCRFElement.Caption, vbCrLf, "<br>")
                End If
                
                If oCRFElement.CaptionFontItalic Then sRtn = sRtn & "</i>"
                If oCRFElement.CaptionFontBold Then sRtn = sRtn & "</b>"
                
                If Not bDefaultFont Then
                    sRtn = sRtn & "</font>"
                End If
                
                sRtn = sRtn & "</div>"
            End If
        Case Else
        End Select
    End If
    End If

    Set oCategory = Nothing
    RtnEformElement = sRtn
End Function

'------------------------------------------------------------------------------------------------------
Private Function SetFontAttributes(ByVal oCRFElement As eFormElementRO, Optional ByVal bIncludeColour As Boolean) As String
'------------------------------------------------------------------------------------------------------
'Retrieve font attributes for a CRFElement
' MLM 14/10/02: If the element is a form or visit date, use fixed style
'------------------------------------------------------------------------------------------------------
Dim sResult As String

'    On Error GoTo ErrLabel
    
    If oCRFElement.ElementUse = eElementUse.EFormVisitDate Then
        sResult = "color:#AAAAAA;font-family:Verdana,helvetica,arial; font-size:8pt;"
    Else
        sResult = ""
        
        'don't get the colour for radio buttons
        If bIncludeColour Then
            If oCRFElement.FontColour <> 0 Then
                sResult = sResult & "color:#" & RtnHTMLCol(oCRFElement.FontColour) & "; "
            End If
        End If
        
        'get the font size
        If oCRFElement.FontSize <> 0 Then
            sResult = sResult & "FONT-SIZE: " & oCRFElement.FontSize & " pt; "
        End If
        
        'check if font is bold
        If oCRFElement.FontBold Then
            sResult = sResult & "FONT-WEIGHT: " & "bold; "
        End If
        
        'check if font is italic
        If oCRFElement.FontItalic Then
            sResult = sResult & "FONT-STYLE: italic; "
        End If
        
        'get font name
        If oCRFElement.FontName <> "" Then
            sResult = sResult & "FONT-FAMILY: " & oCRFElement.FontName & ";"
        End If
    End If
    
    SetFontAttributes = sResult
    
    Exit Function
    
'ErrLabel:
'    SetFontAttributes = ""
End Function

''-------------------------------------------------------------------------------------------'
'Public Sub PublishEformASPFiles(ByVal lStudyCode As Long, _
'                                ByVal nVersion As Integer)
'
''-------------------------------------------------------------------------------------------'
'' Creates ASP eform displayed in the browser - rewrite of CreateHTMLFiles() for v2.2.10 using
'' business objects
''-------------------------------------------------------------------------------------------'
''Revisions:
''   ic  02/04/2002  copied from v3.0 to v 2.2 for improved 2.2 release
''
'
'Dim oStudyDef As StudyDefRO
'Dim oEf As eFormRO
'
'Dim sDatabaseCnn As String
'Dim sPath As String
'Dim sFileName As String
'Dim sDatabase As String
'Dim nFile As Integer
'
'
'
'    'get user database
'    sDatabase = goUser.Database.NameOfDatabase
'
'    'get database connection
'    sDatabaseCnn = gsADOConnectString
'
'    'Create a new StudyDef class and load study def
'    Set oStudyDef = New StudyDefRO
'    Call oStudyDef.Load(sDatabaseCnn, lStudyCode, nVersion)
'
'    'get next free file number
'    nFile = FreeFile
'
'    'get new file directory path
'    sPath = gsWEB_HTML_LOCATION & "asp\eform\"
'
'
'    'loop through each eform in the study
'    For Each oEf In oStudyDef.eForms
'        Call oStudyDef.LoadElements(oEf)
'
'        'build file name - studyid_eformid_database
'        sFileName = oStudyDef.StudyId & "_" & oEf.EFormId & "_" & sDatabase
'
'
'        'make sure folder exists before opening
'        If Not FolderExistence(sPath & sFileName & ".asp") Then
'            Call DialogError("HTML files could not be created because " & gsWEB_HTML_LOCATION & " not found")
'            Exit Sub
'        End If
'
'        'open the file for output
'        Open sPath & sFileName & ".asp" For Output As #nFile
'
'        'write to the file
'        Print #nFile, RtnEformString(oStudyDef, oEf, sDatabase, sFileName)
'
'        'close the file
'        Close #nFile
'
'        'MLM 24/09/02: Also create files for visit forms; add V_ to file names
'        Open sPath & "V_" & sFileName & ".asp" For Output As #nFile
'        Print #nFile, GetVisitEFormString(oStudyDef, oEf, sDatabase, sFileName)
'        Close #nFile
'
'        Debug.Print "completed " & oEf.Name
'
'        'important - remove elements from studydef to free memory
'        Call oStudyDef.RemoveElements(oEf)
'    Next
'
'
'    'destroy object instances
'    Set oEf = Nothing
'    Set oStudyDef = Nothing
'
'End Sub

'-------------------------------------------------------------------------------------------'
Private Function CalculateRQGMaxHeight(oEform As eFormRO, _
                                oRQGElement As eFormElementRO) As Long
'-------------------------------------------------------------------------------------------'
' Calculate the maximum height a RQG can be without overlapping the next
' element (below) it on the eForm
'-------------------------------------------------------------------------------------------'
' REVISIONS
' DPH 29/01/2002 - Adapted RQG max size if last element on an eForm
'-------------------------------------------------------------------------------------------'
Dim lMaxHeight As Long
Dim lTopPosRQG As Long
Dim lLeftPosRQG As Long
Dim lBotPosRQG As Long
Dim lTopPosElement As Long
Dim lLeftPosElement As Long
Dim lRightPosElement As Long
Dim oElement As eFormElementRO
Dim oRQG As QGroupRO
Dim oOwnerRQG As QGroupRO
Dim bBelongsToRQG As Boolean
Dim bRQG As Boolean
Dim bRQGLastElement As Boolean
Dim lRowHeight As Long
Dim nMaxElements As Integer

    ' Get top (Y) position of RQG dealing with
    ' Add 24 for RQG header DIV
    lTopPosRQG = oRQGElement.ElementY + (24 * nIeYSCALE)
    lLeftPosRQG = oRQGElement.ElementX
    ' Make max height RQG an arbitary Max value
    lBotPosRQG = lTopPosRQG + 8000
    bRQGLastElement = True
    
    For Each oElement In oEform.EFormElements
        ' Loop through page elements identifying Non-RQG
        ' and Non-RQG question elements
        Set oRQG = oElement.QGroup
        Set oOwnerRQG = oElement.OwnerQGroup
        If oRQG Is Nothing Then
            bRQG = False
        Else
            bRQG = True
        End If
        If oOwnerRQG Is Nothing Then
            bBelongsToRQG = False
        Else
            bBelongsToRQG = True
        End If

        If Not bRQG And Not bBelongsToRQG Then
            ' Get position on eForm of Element
            lTopPosElement = oElement.ElementY
            lLeftPosElement = oElement.ElementX
            
            ' if lower on form than RQG element then compare else ignore
            If lTopPosElement > lTopPosRQG Then
            
                bRQGLastElement = False
                
                If lTopPosElement < lBotPosRQG Then
                    ' Now check for horizontal overlap
                    ' can do if element comparing with is to left of rqg
                    'If lLeftPosElement < lLeftPosRQG Then
                        lBotPosRQG = lTopPosElement
                    'End If
                End If
            End If
        End If
    Next
    
    ' if RQG is last question on a page
    If bRQGLastElement Then
        ' try to calculate a more realistic max height
        Set oRQG = oRQGElement.QGroup
        lMaxHeight = 0
        If Not (oRQG Is Nothing) Then
            ' set default row height
            lRowHeight = 30 * nIeYSCALE
            ' reset nMaxElements
            nMaxElements = 1
            ' loop through RQG Questions and detect any Radio Buttons
            For Each oElement In oRQG.Elements
                If oElement.ControlType = ControlType.OptionButtons Then
                    ' Get number of elements in Category (height of radio buttons)
                    If oElement.Categories.Count > nMaxElements Then
                        nMaxElements = oElement.Categories.Count
                    End If
                End If
            Next
            lMaxHeight = (lRowHeight * oRQG.DisplayRows * nMaxElements) / nIeYSCALE
        End If
    Else
        ' calculate Maximum height (with web scale) - 5 (so not attached to next element)
        lMaxHeight = ((lBotPosRQG - lTopPosRQG) / nIeYSCALE) - 5
    End If
    
    CalculateRQGMaxHeight = lMaxHeight
End Function

''-------------------------------------------------------------------------------------------'
'Private Function RtnEformString(ByVal oStudyDef As StudyDefRO, _
'                                ByVal oEf As eFormRO, _
'                                ByVal sDatabase As String, _
'                                ByVal sFileName As String) As String
''-------------------------------------------------------------------------------------------'
'' function accepts an studydef object, eform object, database name, filename
'' function returns a string representing the passed eform as an ASP : rewrite for v2.2.10 using
'' business objects
''-------------------------------------------------------------------------------------------'
''Revisions:
''   ic  24/01/2002  changed session variable names 'usrId','usrRl' to 'ssUsrId','ssUsrRl'
''                   added code to create i/o dll name based on version
''                   added fnLockMACRO() js function definition
''   ic  01/02/2002  changed include file to 'checkEFLoggedIn.asp'
''   ic  06/02/2002  added code to js showSaver() function to position saving layer relative to page scrolling
''   ic  02/04/2002  copied from v3.0 to v 2.2 for improved 2.2 release
'' MLM 24/09/02: Cope with visit eforms!
''-------------------------------------------------------------------------------------------'
'
'Dim oCRFElement As eFormElementRO
'Dim sbgcol As String
'Dim sRtn As String
'Dim sIOdllName As String
'
'
'    'get bg colour for page
'    If oEf.BackgroundColour = 0 Then
'        sbgcol = "#" & GetHexColour(oStudyDef.eFormColour)
'    Else
'        sbgcol = "#" & GetHexColour(oEf.BackgroundColour)
'    End If
'
'
'    'start of asp page
'    'page header,include files
'    sRtn = sRtn & "<%@ LANGUAGE=VBScript%>" & vbCrLf _
'                & "<%Response.Buffer = true%>" & vbCrLf _
'                & "<%Response.Expires=0%>" & vbCrLf _
'                & "<!-- #include file=" & Chr(34) & "../general/include/checkSSL.asp" & Chr(34) & " -->" & vbCrLf
'
'    sRtn = sRtn & "<%" & vbCrLf _
'                & "If (Session(" & Chr(34) & "ssUsrName" & Chr(34) & ") = " & Chr(34) & Chr(34) & ") " _
'                & "Or (Session(" & Chr(34) & "ssUsrPassword" & Chr(34) & ") = " & Chr(34) & Chr(34) & ") Then" & vbCrLf _
'                & "%>" & vbCrLf _
'                & "<script language=" & Chr(34) & "javascript" & Chr(34) & ">" & vbCrLf _
'                & "window.parent.fnReLogIn(" & Chr(34) & "<%=Server.URLEncode(" & Chr(34) & "fltEf=" & oEf.EFormId _
'                                                                         & "&fltId=" & Chr(34) & " & Request.Querystring(" & Chr(34) & "fltId" & Chr(34) & ") " _
'                                                                         & "& " & Chr(34) & "&fltDb=" & sDatabase _
'                                                                         & "&fltSt=" & oStudyDef.StudyId _
'                                                                         & "&fltSi=" & Chr(34) & " & Request.Querystring(" & Chr(34) & "fltSi" & Chr(34) & ") " _
'                                                                         & "& " & Chr(34) & "&fltSj=" & Chr(34) & " & Request.Querystring(" & Chr(34) & "fltSj" & Chr(34) & ")" _
'                                                                         & ")%>" & Chr(34) & ");" & vbCrLf _
'                & "</script>" & vbCrLf _
'                & "<%" & vbCrLf _
'                & "Response.End" & vbCrLf _
'                & "End If" & vbCrLf _
'                & "%>" & vbCrLf
'
'    sRtn = sRtn & "<!-- #include file=" & Chr(34) & "../general/include/miscFunc.asp" & Chr(34) & "-->" & vbCrLf
'
'
'    'copyright notice
'    sRtn = sRtn & "<%" & vbCrLf _
'                & "Response.Write(" & Chr(34) & "<!-- " & Chr(34) & " & Application(" & Chr(34) & "asCOPYRIGHT" & Chr(34) & ") & " & Chr(34) & " -->" & Chr(34) & ")" & vbCrLf
'
'
'    'dim statements
'    'MLM 24/09/02: Added sVisitEForm
'    sRtn = sRtn & "dim oIo" & vbCrLf _
'                & "dim usrName" & vbCrLf _
'                & "dim usrRl" & vbCrLf _
'                & "dim fltDb" & vbCrLf _
'                & "dim fltSt" & vbCrLf _
'                & "dim fltSi" & vbCrLf _
'                & "dim fltSj" & vbCrLf _
'                & "dim fltId" & vbCrLf _
'                & "dim nxt" & vbCrLf _
'                & "dim aErrors" & vbCrLf _
'                & "dim nLoop" & vbCrLf _
'                & "dim nSaveOk" & vbCrLf _
'                & "dim nStatus" & vbCrLf _
'                & "dim sModule" & vbCrLf _
'                & "dim sVisitEForm" & vbCrLf
'
'
'    'assign passed/session var values to local vars
'    sRtn = sRtn & "usrName = session(" & Chr(34) & "ssUsrName" & Chr(34) & ")" & vbCrLf _
'                & "usrRl = session(" & Chr(34) & "ssUsrRl" & Chr(34) & ")" & vbCrLf _
'                & "fltDb = " & Chr(34) & sDatabase & Chr(34) & vbCrLf _
'                & "fltSt = " & Chr(34) & oStudyDef.StudyId & Chr(34) & vbCrLf _
'                & "fltSi = Request.QueryString(" & Chr(34) & "fltSi" & Chr(34) & ")" & vbCrLf _
'                & "fltSj = Request.QueryString(" & Chr(34) & "fltSj" & Chr(34) & ")" & vbCrLf _
'                & "fltId = Request.QueryString(" & Chr(34) & "fltId" & Chr(34) & ")" & vbCrLf _
'                & "sModule = Request.QueryString(" & Chr(34) & "module" & Chr(34) & ")" & vbCrLf
'
'    'ic 24/01/2002 added code to create i/o dll name based on version
'    'create i/o dll name using application version
'    sIOdllName = "MACROWWWIO" & App.Major & App.Minor & ".clsWWW"
'
'    '<-- start of data-entry/view data permission 'if' condition
'    sRtn = sRtn & "if ((rtnArrayItem(Application(" & Chr(34) & "asFnDataEntry" & Chr(34) & ")) <> " & Chr(34) & Chr(34) & ") " _
'                & "and (rtnArrayItem(Application(" & Chr(34) & "asFnViewData" & Chr(34) & ")) <> " & Chr(34) & Chr(34) & ")) then" & vbCrLf
'
'    'create i/o object,check de permission 'if' condition
'    sRtn = sRtn & "set oIo = Server.CreateObject(" & Chr(34) & sIOdllName & Chr(34) & ")" & vbCrLf
'
'
'    'check that the subject being loaded is locked
''    sRtn = sRtn & "if thisSubjectLocked() then" & vbCrLf
'
'
'    'get 'nxt' value from hidden field. this field is in all de eforms and its name is the same as
'    'the eforms file name, prefixed with 'f_' (because field names must begin with a letter) without '.asp' extension
'    'if a value for this field is present in the form object it means that the page has submitted to itself.
'    'the field will contain an integer value. 0=reload this eform, 1=load next eform, 2=load previous eform
'    sRtn = sRtn & "nxt = Request.Form(" & Chr(34) & "f_" & sFileName & Chr(34) & ")" & vbCrLf
'
'
'    'set nStatus = 0. nStatus is used in the page to determine the content.
'    sRtn = sRtn & "nStatus = 0" & vbCrLf
'
'
'    'if 'nxt' has a value the form has submitted to itself and should be saved
'    'call SaveForm() function. if function returns 1, save was ok, else save of one or more responses failed on
'    'validation - set nStatus to 1 to display this error when the page is displayed in the browser.
'    'if save was ok and nxt value is 1 or 2 (next or previous eform) then get the url of the requested eform from
'    'MoveNext() function
'    sRtn = sRtn & "If nxt <> " & Chr(34) & Chr(34) & " Then" & vbCrLf _
'                & "saveFrm = Request.Form()" & vbCrLf _
'                & "aErrors = oIo.SaveForm(usrName,Session(" & Chr(34) & "ssUsrPassword" & Chr(34) & "),fltDb,fltSi,fltSt,fltSj,fltId,saveFrm)" & vbCrLf _
'                & "If (IsEmpty(aErrors)) Then" & vbCrLf _
'                & "Select Case (nxt)" & vbCrLf _
'                & "Case 0:" & vbCrLf _
'                & "Case 1,2: url = oIo.MoveNext(usrName,fltDb,fltSi,fltSt,fltSj,fltId,nxt)" & vbCrLf _
'                & "Case 3: nStatus = 4" & vbCrLf _
'                & "End Select" & vbCrLf _
'                & "Else" & vbCrLf _
'                & "nStatus = 1" & vbCrLf _
'                & "End If" & vbCrLf _
'                & "End If" & vbCrLf
'
'    'if url variable is empty load this page (either first visit, save failed or user requested to return to this page
'    'if url has a value it is a pipe (|) delimited string containing 2 values. first value is next non-repeating eform,
'    'second value is next repeating eform. if both values are empty this is the last (or if previous eform was requested,
'    'the first) eform in the visit. set nStatus to 2. if the values are different there are several possibilities:
'    'val1 is empty, val2 has string - there are only repeating eforms remaining in this visit
'    'val1 has string, val2 is empty - there are only non repeating eforms remaining in this visit
'    'val1 has string, val2 has string - there are repeating and non repeating eforms remaining in this visit
'    'set nStatus to 3
'    sRtn = sRtn & "If (url <> " & Chr(34) & Chr(34) & ") Then" & vbCrLf _
'                & "url = split(url," & Chr(34) & "|" & Chr(34) & ")" & vbCrLf _
'                & "If (url(0) & url(1) = " & Chr(34) & Chr(34) & ") Then" & vbCrLf _
'                & "nStatus = 2" & vbCrLf _
'                & "ElseIf (url(0) <> url(1)) Then" & vbCrLf _
'                & "nStatus = 3" & vbCrLf _
'                & "Else" & vbCrLf _
'                & "Response.Redirect(url(0) & " & Chr(34) & "&module=" & Chr(34) & " & sModule" & ")" & vbCrLf _
'                & "End If" & vbCrLf _
'                & "End If" & vbCrLf
'
'
'    'case branch on nStatus value displays different page depending on processing outcomes above
'    'first branch:
'    'nStatus is 0 - first visit to this page, display with no messages
'    'nStatus is 1 - save of one or more responses failed. display this page, but add an alert in js pageLoaded() function
'    sRtn = sRtn & "Select Case (nStatus)" & vbCrLf _
'                & "Case 0,1:" & vbCrLf _
'                & "%>"
'
'
'    'start of html
'    sRtn = sRtn & "<html>" & vbCrLf _
'                & "<head>" & vbCrLf
'
'
'    'radio control include
'    sRtn = sRtn & "<script language=" & Chr(34) & "javascript" & Chr(34) & " src=" & Chr(34) & "../../script/RadioControl.js" & Chr(34) & "></script>" & vbCrLf
'
'
'    'stylesheet for page - this is the study default, only text not matching this default will need to be defined on the page
'    sRtn = sRtn & "<style type=" & Chr(34) & "text/css" & Chr(34) & "> " _
'                & "div{position:absolute; width:auto; height:auto; " _
'                & "font-family:" & oStudyDef.FontName & " ;font-size:" & oStudyDef.FontSize & " pt; color:" & GetHexColour(oStudyDef.FontColour) & "} " _
'                & "td{font-family:" & oStudyDef.FontName & " ;font-size:" & oStudyDef.FontSize & " pt; color:" & GetHexColour(oStudyDef.FontColour) & "}" _
'                & "</style>" & vbCrLf
'
'
'    'ic 23/08/2002 changed fnInitialiseApplet() call to pass integer, not boolean to solve regional problem
'    'ic 07/02/2002 added RtnEformCategoryTxtFunc() function call
'    'ic 24/01/2002 added fnLockMACRO() js function definition
'    'javascript definitions and includes
'    'MLM 25/09/02: Added calls to InitVEFIRadios() and InitVEFICategories() to pageLoaded function
'    sRtn = sRtn & "<script language=" & Chr(34) & "javascript" & Chr(34) & ">" & vbCrLf _
'                & "var frm1;" & vbCrLf _
'                & "var frm2;" & vbCrLf _
'                & "var frm3;" & vbCrLf _
'                & "var aRadioList = new Array();" & vbCrLf _
'                & "function fnLockMACRO(rtnUrl){" & vbCrLf _
'                & "window.parent.navigate('../general/login.asp?action=lock&url=<%=Server.UrlEncode(" & Chr(34) & "eformFrm.asp?fltDb=" & sDatabase & "&fltEf=" & oEf.EFormId & "&fltId=" & Chr(34) & " & fltId & " & Chr(34) & "&fltSt=" & Chr(34) & " & fltSt & " & Chr(34) & "&fltSi=" & Chr(34) & " & fltSi & " & Chr(34) & "&fltSj=" & Chr(34) & " & fltSj & " & Chr(34) & "&module=" & Chr(34) & " & sModule)%>&rtnUrl='+rtnUrl)}" & vbCrLf _
'                & "function fnChangeUser(rtnUrl){" & vbCrLf _
'                & "window.parent.navigate('../general/login.asp?action=change&url=<%=Server.UrlEncode(" & Chr(34) & "eformFrm.asp?fltDb=" & sDatabase & "&fltEf=" & oEf.EFormId & "&fltId=" & Chr(34) & " & fltId & " & Chr(34) & "&fltSt=" & Chr(34) & " & fltSt & " & Chr(34) & "&fltSi=" & Chr(34) & " & fltSi & " & Chr(34) & "&fltSj=" & Chr(34) & " & fltSj & " & Chr(34) & "&module=" & Chr(34) & " & sModule)%>&rtnUrl='+rtnUrl)}" & vbCrLf _
'                & "function pageLoaded(){" & vbCrLf _
'                & "frm1=window.parent.frames[0];" & vbCrLf _
'                & "frm2=window.parent.frames[1].document.deFrm;" & vbCrLf _
'                & "frm3=window.parent.frames[1].document.all.tags(" & Chr(34) & "div" & Chr(34) & ");" & vbCrLf _
'                & "frm1.fnInitialiseApplet(<%=cint(rtnArrayItem(Application(" & Chr(34) & "asFnOverruleWarnings" & Chr(34) & ")) <> " & Chr(34) & Chr(34) & ")%>,frm2," & Chr(34) & sbgcol & Chr(34) & ",null,window.parent.frames[1]);" & vbCrLf _
'                & RtnRadioCreateFuncs(oEf) & vbCrLf _
'                & "InitVEFIRadios();" & vbCrLf _
'                & "JSVEInitFunc();" & vbCrLf _
'                & RtnEformCategoryTxtFunc(oEf) & vbCrLf _
'                & "InitVEFICategories();" & vbCrLf
'
'    'ic 06/02/2002 added code to js showSaver() function to position saving layer relative to page scrolling
'    'MLM 25/09/02: Added empty InitVEFIRadios(), InitVEFICategories() and VisitDateOk() functions.
'    '   Also, pass window to fnFinalisePage. Only proceed if returns true.
'    '   Include form date function
'    sRtn = sRtn & "<%If (nxt <> " & Chr(34) & Chr(34) & ") Then" & vbCrLf _
'                & "If (Not IsEmpty(aErrors)) Then " & vbCrLf _
'                & "Response.Write(" & Chr(34) & "alert('MACRO encountered problems while saving. Some responses could not be saved.\nUnsaved responses are listed below\n\n" & Chr(34) & ")" & vbCrLf _
'                & "for nLoop = lbound(aErrors,2) to ubound(aErrors,2)" & vbCrLf _
'                & "Response.Write(aErrors(0,nLoop) & " & Chr(34) & " - " & Chr(34) & " & aErrors(1,nLoop) & " & Chr(34) & "\n" & Chr(34) & ")" & vbCrLf _
'                & "next" & vbCrLf _
'                & "response.write(" & Chr(34) & " ');" & Chr(34) & ")" & vbCrLf _
'                & "End If" & vbCrLf _
'                & "End If%>" & vbCrLf _
'                & "hideLoader();}" & vbCrLf _
'                & "function saveEform(nxt){if(frm1.fnFinalisePage(window, frm2)){showSaver();document.deFrm[" & Chr(34) & "f_" & sFileName & Chr(34) & "].value=nxt;frm2.submit();}}" & vbCrLf _
'                & "function hideLoader(){document.all.wholeDiv.style.visibility='visible';window.parent.frames[0].fnDisplayMenu(true);frm1.fnDisplayFormName(" & Chr(34) & oEf.Name & Chr(34) & ");document.all.loadingDiv.style.visibility='hidden';}" & vbCrLf _
'                & "function showSaver(){document.all.wholeDiv.style.visibility='hidden';document.all.savingDiv.style.pixelLeft=document.body.scrollLeft+50;document.all.savingDiv.style.pixelTop=document.body.scrollTop+50;document.all.savingDiv.style.visibility='visible';}" & vbCrLf _
'                & "function logOut(){if(confirm('About to close current MACRO session?'))window.parent.navigate(" & Chr(34) & "../general/Logout.asp" & Chr(34) & ")}" & vbCrLf _
'                & "function InitVEFIRadios(){}" & vbCrLf _
'                & "function InitVEFICategories(){}" & vbCrLf _
'                & "function VisitDateOk(){return true;}" & vbCrLf _
'                & GetDateOkFunc(eEFormUse.User, oEf) & vbCrLf _
'                & "</script>" & vbCrLf
'
'
'    'style definitions displaying loading/saving messages
'    sRtn = sRtn & "<link rel=" & Chr(34) & "stylesheet" & Chr(34) & " href=" & Chr(34) & "../../style/messagebox.css" & Chr(34) & " type=" & Chr(34) & "text/css" & Chr(34) & ">" & vbCrLf _
'                & "<link rel=" & Chr(34) & "stylesheet" & Chr(34) & " href=" & Chr(34) & "../../style/MACRO1.css" & Chr(34) & " type=" & Chr(34) & "text/css" & Chr(34) & ">" & vbCrLf _
'                & "<style type=" & Chr(34) & "text/css" & Chr(34) & ">" _
'                & "#wholeDiv{position:relative; visibility:hidden;}" _
'                & "#loadingDiv{position:absolute; visibility:visible; left:50; top:50; width:300px; height:100px; z-index:100;}" _
'                & "#savingDiv{position:absolute; visibility:hidden; left:50; top:50; width:300px; height:100px; z-index:100;}" _
'                & "</style>" & vbCrLf _
'                & "</head>" & vbCrLf
'
'
'    'start of html body
'    sRtn = sRtn & "<body bgcolor=" & Chr(34) & sbgcol & Chr(34) & " onload=" & Chr(34) & "javascript:pageLoaded();" & Chr(34) & ">"
'
'
'    'layer definitions displaying loading/saving messages. note these must be inside the body tag for ie4
'    sRtn = sRtn & "<div class=" & Chr(34) & "PROCESS_MESSAGE_BOX" & Chr(34) & " id=" & Chr(34) & "loadingDiv" & Chr(34) & ">" & vbCrLf _
'                & "<table height=" & Chr(34) & "100%" & Chr(34) & " align=" & Chr(34) & "center" & Chr(34) & " width=" & Chr(34) & "90%" & Chr(34) & " class=" & Chr(34) & "MESSAGE" & Chr(34) & ">" & vbCrLf _
'                & "<tr><td valign=" & Chr(34) & "middle" & Chr(34) & "><b>please wait</b><br><br><img src=" & Chr(34) & "../../img/clock.gif" & Chr(34) & ">" & vbCrLf _
'                & "&nbsp;&nbsp;loading " & oEf.Name & "...</td></tr></table></div>" & vbCrLf _
'                & "<div class=" & Chr(34) & "PROCESS_MESSAGE_BOX" & Chr(34) & " id=" & Chr(34) & "savingDiv" & Chr(34) & ">" & vbCrLf _
'                & "<table height=" & Chr(34) & "100%" & Chr(34) & " align=" & Chr(34) & "center" & Chr(34) & " width=" & Chr(34) & "90%" & Chr(34) & " class=" & Chr(34) & "MESSAGE" & Chr(34) & ">" & vbCrLf _
'                & "<tr><td valign=" & Chr(34) & "middle" & Chr(34) & "><b>please wait</b><br><br><img src=" & Chr(34) & "../../img/clock.gif" & Chr(34) & ">" & vbCrLf _
'                & "&nbsp;&nbsp;saving " & oEf.Name & "...</td></tr></table></div>" & vbCrLf
'
'
'    'asp flush command to display loading messages on page before the page has completed loading
'    sRtn = sRtn & "<%Response.Flush" & vbCrLf
'
'
'    'write JSVE initialisation function
'    sRtn = sRtn & "Response.Write(oIo.getJSVEInitFunc(usrName,Session(" & Chr(34) & "ssUsrPassword" & Chr(34) & "),fltDb,fltSi,fltSt,fltSj,fltId))" & vbCrLf _
'                & "%>" & vbCrLf
'
'
'    'write 'wholeDiv' enclosing page for hiding on save
'    sRtn = sRtn & "<div id=" & Chr(34) & "wholeDiv" & Chr(34) & ">"
'
'
'    'form definition, the form submits back to this asp page for whatever processing is requested. the hidden field holds
'    'this asp's filename so that when processing submitted form elements, the asp can tell which page the elements belong to
'    sRtn = sRtn & "<form name=" & Chr(34) & "deFrm" & Chr(34) & " method=" & Chr(34) & "post" & Chr(34) & " action=" & Chr(34) & sFileName & ".asp?fltSi=<%=fltSi%>&fltSj=<%=fltSj%>&fltId=<%=fltId%>&module=<%=sModule%>" & Chr(34) & ">" & vbCrLf _
'                & "<input type=" & Chr(34) & "hidden" & Chr(34) & " name=" & Chr(34) & "f_" & sFileName & Chr(34) & ">" & vbCrLf
'
'
'    'loop through each eform element on the eform
'    For Each oCRFElement In oEf.EFormElements
'
'        'call function to get html string representing the eform element
'        sRtn = sRtn & RtnEformElementString(oStudyDef, oCRFElement, sbgcol, oEf.displayNumbers)
'    Next
'
'    'MLM 25/09/02: Here, do all the stuff pertaining to the visit form, by .Executing its ASP if needed
'    sRtn = sRtn & vbCrLf & "<%sVisitEForm = oIo.GetVisitEForm(usrName,Session(" & Chr(34) & "ssUsrPassword" & Chr(34) & "),fltDb,fltSi,fltSt,fltSj,fltId)" & vbCrLf & _
'        "If sVisitEForm <> """" Then" & vbCrLf & _
'        "   Server.Execute sVisitEForm" & vbCrLf & _
'        "End If%>" & vbCrLf
'
'    sRtn = sRtn & "<div onfocus=" & Chr(34) & "if(!frm1.fnLostFocus(frm1.sLastFocusID))frm1.setFieldFocus(frm1.sLastFocusID);" & Chr(34) _
'              & "tabindex=" & Chr(34) & "999" & Chr(34) & "> </div>"
'
'    'close the form
'    sRtn = sRtn & "</form>" _
'                & "</div>" _
'                & "</body>" _
'                & "</html>"
'
'
'    'second branch:
'    'nStatus is 2 - next or previous eform was requested, but this is the last or first eform in the current visit
'    sRtn = sRtn & "<% Case 2:%>" & vbCrLf _
'                & "<html>" & vbCrLf _
'                & "<head>" & vbCrLf _
'                & "<link rel=" & Chr(34) & "stylesheet" & Chr(34) & " href=" & Chr(34) & "../../style/MACRO1.css" & Chr(34) & " type=" & Chr(34) & "text/css" & Chr(34) & ">" & vbCrLf _
'                & "<meta http-equiv=" & Chr(34) & "Pragma" & Chr(34) & " content=" & Chr(34) & "no-cache" & Chr(34) & ">" & vbCrLf
'
'    sRtn = sRtn & "<script language=" & Chr(34) & "javascript" & Chr(34) & ">" & vbCrLf _
'                & "function showMsg(){" & vbCrLf _
'                & "<%If (nxt = 1) Then%>" & vbCrLf _
'                & "alert(" & Chr(34) & "This is the last form in the current visit." & Chr(34) & ");" & vbCrLf _
'                & "<%Else%>" & vbCrLf _
'                & "alert(" & Chr(34) & "This is the first form in the current visit." & Chr(34) & ");" & vbCrLf _
'                & "<%End If%>" & vbCrLf _
'                & "window.parent.frames[0].fnGoMenu(4);" & vbCrLf _
'                & "}</script>" & vbCrLf _
'                & "</head>" & vbCrLf _
'                & "<body onload=" & Chr(34) & "showMsg();" & Chr(34) & ">" & vbCrLf _
'                & "</body>" & vbCrLf _
'                & "</html>"
'
'
'    'third branch:
'    'nStatus is 3 - the next eform was requested, but the next eform is cycling
'    sRtn = sRtn & "<% Case 3:%>" & vbCrLf _
'                & "<html>" & vbCrLf _
'                & "<head>" & vbCrLf _
'                & "<link rel=" & Chr(34) & "stylesheet" & Chr(34) & " href=" & Chr(34) & "../../style/MACRO1.css" & Chr(34) & " type=" & Chr(34) & "text/css" & Chr(34) & ">" & vbCrLf _
'                & "<meta http-equiv=" & Chr(34) & "Pragma" & Chr(34) & " content=" & Chr(34) & "no-cache" & Chr(34) & ">" & vbCrLf
'
'
'    sRtn = sRtn & "<script language=" & Chr(34) & "javascript" & Chr(34) & ">" & vbCrLf _
'                & "function showMsg(){" & vbCrLf _
'                & "var bNew=(confirm(" & Chr(34) & "Would you like to open a new " & oEf.Name & " form?" & Chr(34) & "));" & vbCrLf _
'                & "<%If (url(0) <> " & Chr(34) & Chr(34) & ") and (url(1) <> " & Chr(34) & Chr(34) & ") Then%>" & vbCrLf _
'                & "if (bNew){" _
'                & "window.location.replace(" & Chr(34) & "<%=url(1) & " & Chr(34) & "&module=" & Chr(34) & " & sModule%>" & Chr(34) & ")" _
'                & "}else{" _
'                & "window.location.replace(" & Chr(34) & "<%=url(0) & " & Chr(34) & "&module=" & Chr(34) & " & sModule%>" & Chr(34) & ")" _
'                & "}" & vbCrLf _
'                & "<%Else%>" & vbCrLf _
'                & "if (bNew){" & "window.location.replace(" & Chr(34) & "<%=url(1) & " & Chr(34) & "&module=" & Chr(34) & " & sModule%>" & Chr(34) & ")" & "}else{" _
'                & "alert(" & Chr(34) & "This is the last form in the current visit." & Chr(34) & ");" & vbCrLf _
'                & "window.parent.frames[0].fnGoMenu(4);}" & vbCrLf _
'                & "<%End If%>" & vbCrLf _
'                & "}</script>" & vbCrLf _
'                & "</head>" & vbCrLf _
'                & "<body onload=" & Chr(34) & "showMsg();" & Chr(34) & ">" & vbCrLf _
'                & "</body>" & vbCrLf _
'                & "</html>"
'
'
'    'final branch:
'    'nStatus is 4 - save and return to schedule/databrowser was requested
'    sRtn = sRtn & "<% Case Else:%>" & vbCrLf _
'                & "<html>" & vbCrLf _
'                & "<head>" & vbCrLf _
'                & "</head>" & vbCrLf _
'                & "<meta http-equiv=" & Chr(34) & "Pragma" & Chr(34) & " content=" & Chr(34) & "no-cache" & Chr(34) & ">" & vbCrLf _
'                & "</head>" & vbCrLf _
'                & "<body onload=" & Chr(34) & "window.parent.frames[0].fnGoMenu(4);" & Chr(34) & ">" & vbCrLf _
'                & "</body>" & vbCrLf _
'                & "</html>"
'
'
'    'end of case branch
'    sRtn = sRtn & "<%End Select" & vbCrLf
'
'
''    'the subject wasnt correctly locked
''    sRtn = sRtn & "else%>" & vbCrLf _
''                & "<html>" & vbCrLf _
''                & "<head>" & vbCrLf _
''                & "<link rel=" & Chr(34) & "stylesheet" & Chr(34) & " href=" & Chr(34) & "../../style/MACRO1.css" & Chr(34) & " type=" & Chr(34) & "text/css" & Chr(34) & ">" & vbCrLf _
''                & "<meta http-equiv=" & Chr(34) & "Pragma" & Chr(34) & " content=" & Chr(34) & "no-cache" & Chr(34) & ">" & vbCrLf _
''                & "</head>" & vbCrLf _
''                & "<body>" & vbCrLf _
''                & "<table width=" & Chr(34) & "100%" & Chr(34) & " height=" & Chr(34) & "95%" & Chr(34) & ">" & vbCrLf _
''                & "<tr><td align=" & Chr(34) & "center" & Chr(34) & " class=" & Chr(34) & "MESSAGE" & Chr(34) & ">" _
''                & "This subject has not been correctly locked. <a href=" & Chr(34) & "../general/selectDatabase.asp" & Chr(34) & " target=" & Chr(34) & "_top" & Chr(34) & ">Please re-select the subject</a>.</td></tr>" & vbCrLf _
''                & "</table>" & vbCrLf _
''                & "</body>" & vbCrLf _
''                & "</html>" _
''                & "<%End If%>" & vbcrlf
'
'
'    'the user doesnt have de permission
'    sRtn = sRtn & "Else%>" & vbCrLf _
'                & "<html>" & vbCrLf _
'                & "<head>" & vbCrLf _
'                & "<link rel=" & Chr(34) & "stylesheet" & Chr(34) & " href=" & Chr(34) & "../../style/MACRO1.css" & Chr(34) & " type=" & Chr(34) & "text/css" & Chr(34) & ">" & vbCrLf _
'                & "<meta http-equiv=" & Chr(34) & "Pragma" & Chr(34) & " content=" & Chr(34) & "no-cache" & Chr(34) & ">" & vbCrLf _
'                & "</head>" & vbCrLf _
'                & "<body>" & vbCrLf _
'                & "<table width=" & Chr(34) & "100%" & Chr(34) & " height=" & Chr(34) & "95%" & Chr(34) & ">" & vbCrLf _
'                & "<tr><td align=" & Chr(34) & "center" & Chr(34) & " class=" & Chr(34) & "MESSAGE" & Chr(34) & ">" _
'                & "You do not have permission to access this subject in MACRO Web Data Entry.</td></tr>" & vbCrLf _
'                & "</table>" & vbCrLf _
'                & "</body>" & vbCrLf _
'                & "</html>"
'
'
'    'end of conditions
'    sRtn = sRtn & "<%End If" & vbCrLf
'
'
'    'destroy i/o
'    sRtn = sRtn & "Set oIo = Nothing%>"
'
'
'    Set oCRFElement = Nothing
'    RtnEformString = sRtn
'End Function

'-------------------------------------------------------------------------------------------'
Private Function GetDateOkFunc(ByRef enEFormUse As eEFormUse, ByRef oEform As eFormRO) As String
'-------------------------------------------------------------------------------------------'
' MLM 25/09/02: Created. Returns a JavaScript function that returns false if there is an blank
'   enterable form/visit date (which should prevent the form from being saved).
'   enEFormUse determines the name of the function.
'-------------------------------------------------------------------------------------------'

Dim sResult As String
Dim oFormDate As eFormElementRO

    'function header
    If enEFormUse = eEFormUse.VisitEForm Then
        sResult = "function VisitDateOk(){"
    Else
        sResult = "function FormDateOk(){"
    End If
    
    'function body
    Set oFormDate = oEform.EFormDateElement
    If oFormDate Is Nothing Then
        'always ok to save
        sResult = sResult & "return true"
    Else
        sResult = sResult & "return !(frm1.fnEnterable('" & oFormDate.WebId _
            & "') && (''==frm1.fnGetFormatted('" & oFormDate.WebId & "')))"
    End If
    
    'function terminator
    GetDateOkFunc = sResult & ";}"

End Function

''-------------------------------------------------------------------------------------------'
'Private Function GetEFormString(ByRef oEForm As eFormRO, _
'                                ByRef sDatabase As String, _
'                                ByRef sFileName As String) As String
''-------------------------------------------------------------------------------------------'
'' MLM 24/09/02: Created. Creates the contents of an ASP representing a visit eForm.
''-------------------------------------------------------------------------------------------'
'
'Dim sResult As String
'Dim oCRFElement As eFormElementRO
'Dim sbgcol As String
'
'    'get bg colour for page
'    If oEf.BackgroundColour = 0 Then
'        sbgcol = "#" & GetHexColour(oStudyDef.eFormColour)
'    Else
'        sbgcol = "#" & GetHexColour(oEf.BackgroundColour)
'    End If
'
'    'ASP headers
'    sResult = "<%@ LANGUAGE=VBScript%>" & vbCrLf
'
'    'We need a JavaScript code block,
'    sResult = sResult & "<script language=javascript>" & vbCrLf
'    'containing functions for the visit form's radio buttons,
'    sResult = sResult & "function InitVEFIRadios(){" & vbCrLf & _
'        RtnRadioCreateFuncs(oEf) & "}" & vbCrLf
'    'category lists
'    sResult = sResult & "function InitVEFICategories(){" & vbCrLf & _
'        RtnEformCategoryTxtFunc(oEf) & "}" & vbCrLf
'    'and visit date.
'    sResult = sResult & GetDateOkFunc(eEFormUse.VisitEForm, oEf) & vbCrLf & _
'        "</script>" & vbCrLf
'
'    'and write the necessary form elements
'    '(should all be hidden apart from the visit date, if there is one)
'    For Each oCRFElement In oEf.EFormElements
'        sResult = sResult & RtnEformElementString(oStudyDef, oCRFElement, sbgcol, oEf.displayNumbers, eEFormUse.VisitEForm)
'    Next
'
'    GetVisitEFormString = sResult
'
'End Function


'-------------------------------------------------------------------------------------------'
Private Function RtnRadioCreateFuncs(ByVal oEf As eFormRO) As String
'-------------------------------------------------------------------------------------------'
' function accepts an eform object
' function builds a string made up of js function calls to fnCreateRadio
' these calls initialise custom radio controls on the eform
'-------------------------------------------------------------------------------------------'
'Revisions:
'ZA 31/07/2002 - pass font properties while creating radio buttons
'DPH 15/10/2002 - fnGotFocus, fnLostFocus calls to use 'this' parameter
'               Repeat object for Radio buttons
'DPH 26/11/2002 Space TabIndex by mnTABINDEXGAP to allow for RQG inbetween
'DPH 20/06/2003 Get colour & style separately
' DPH 04/09/2003 - will not have a proper tab order but set to 150 to avoid VisitDate/eFormDate
' ic 17/09/2003 skip hidden controls
'-------------------------------------------------------------------------------------------'
Dim oCRFElement As eFormElementRO
Dim oCategory As CategoryItem
Dim sRtn As String
Dim sStyle As String
Dim nRepeat As Integer
Dim oRQG As QGroupRO
Dim oOwnerRQG As QGroupRO
' DPH 26/11/2002
Dim lTabIndex As Long
' DPH 20/06/2003 sColour added
Dim sColour As String

    sRtn = ""
    sStyle = ""
    nRepeat = 0


            
    'loop through all eform elements
    For Each oCRFElement In oEf.EFormElements
    
        ' DPH 18/10/2002 - Retrieve RQG element information
        Set oRQG = oCRFElement.QGroup
        Set oOwnerRQG = oCRFElement.OwnerQGroup

        ' Only create if not belonging to a RQG
        'ic 17/09/2003 skip hidden controls too
        If (oRQG Is Nothing) And (oOwnerRQG Is Nothing) And Not oCRFElement.Hidden Then
            sStyle = Chr(34) & SetFontAttributes(oCRFElement, False) & Chr(34)
            ' DPH 20/06/2003 - Get font colour
            If oCRFElement.FontColour <> 0 Then
                sColour = Chr(34) & RtnHTMLCol(oCRFElement.FontColour) & Chr(34)
            Else
                sColour = Chr(34) & "000000" & Chr(34)
            End If
            
            ' DPH 26/11/2002 - TabIndex
            If oCRFElement.ElementOrder > 1 Then
                lTabIndex = CLng(oCRFElement.ElementOrder - 1) * mnTABINDEXGAP
            Else
                ' DPH 04/09/2003 - will not have a proper tab order but set to 150 to avoid VisitDate/eFormDate
                lTabIndex = 150
            End If
            
            Select Case oCRFElement.ControlType
            Case ControlType.OptionButtons, ControlType.PushButtons
             
                ' DPH 15/10/2002 Repeat Object for Radio Buttons
                sRtn = sRtn & "if(aRadioList[" & Chr(34) & oCRFElement.WebId & Chr(34) & "]==null)"
                sRtn = sRtn & "{aRadioList[" & Chr(34) & oCRFElement.WebId & Chr(34) & "]=new Object();}" & vbCrLf
        
                sRtn = sRtn & "if(aRadioList[" & Chr(34) & oCRFElement.WebId & Chr(34) & "].olRepeat==null)"
                sRtn = sRtn & "{aRadioList[" & Chr(34) & oCRFElement.WebId & Chr(34) & "].olRepeat=new Array();}" & vbCrLf
     
                sRtn = sRtn & "aRadioList[" & Chr(34) & oCRFElement.WebId & Chr(34) & "].olRepeat[" & nRepeat & "]=new Object();" & vbCrLf
                
    '            sRtn = sRtn & "aRadioList[" & Chr(34) & oCRFElement.WebId & Chr(34) & "] = fnRadioCreate(o2,s1," _
    '                                       & Chr(34) & oCRFElement.WebId & "_inpDiv" & Chr(34) & "," & Chr(34)
                                          
                sRtn = sRtn & "aRadioList[" & Chr(34) & oCRFElement.WebId & Chr(34) & "].olRepeat[" & nRepeat & "] = fnRadioCreate(o2,s1," _
                                           & Chr(34) & oCRFElement.WebId & "_inpDiv" & Chr(34) & "," & Chr(34)
                
                For Each oCategory In oCRFElement.Categories
                    ' DPH 20/06/2003 - Make sure category is active before adding to list
                    If oCategory.Active Then
                        sRtn = sRtn & oCategory.Code & "¬" & oCategory.Value & "~"
                    End If
                Next
                sRtn = Left(sRtn, (Len(sRtn) - 1)) & Chr(34) & ","
                
                ' DPH 20/06/2003 - Added sColour
                sRtn = sRtn & "null" & "," _
                            & lTabIndex & "," _
                            & Chr(34) & oCRFElement.WebId _
                            & Chr(34) & "," & sColour & "," & sStyle & ");" & vbCrLf
                            
                ' Now set up location for use in JavaScript
                sRtn = sRtn & "aRadioList[" & Chr(34) & oCRFElement.WebId & Chr(34) & "].olRepeat[" & nRepeat & "].oValueRef=eval('"
                sRtn = sRtn & "o2.document.all[" & Chr(34) & oCRFElement.WebId & Chr(34) & "]');" & vbCrLf
                'aRadioList[oFieldTemp.sID].olRepeat[i].oValueRef=eval('frm2.document.all["'+oFieldTemp.sID+'"]['+i+']');
                
                'DPH 15/10/2002 - fnGotFocus, fnLostFocus calls to use 'this' parameter
                'sRtn = sRtn & "aRadioList[" & Chr(34) & oCRFElement.WebId & Chr(34) & "].onfocus=function(){o1.fnGotFocus('" & oCRFElement.WebId & "');};" & vbCrLf
                'sRtn = sRtn & "aRadioList[" & Chr(34) & oCRFElement.WebId & Chr(34) & "].onchange=function(){o1.fnLostFocus('" & oCRFElement.WebId & "');};" & vbCrLf
                sRtn = sRtn & "aRadioList[" & Chr(34) & oCRFElement.WebId & Chr(34) & "].olRepeat[" & nRepeat & "].onfocus=function(){o1.fnGotFocus(this);};" & vbCrLf
                sRtn = sRtn & "aRadioList[" & Chr(34) & oCRFElement.WebId & Chr(34) & "].olRepeat[" & nRepeat & "].onchange=function(){o1.fnLostFocus(this);};" & vbCrLf
                
            Case Else
            End Select
        
        End If
    Next
    
    RtnRadioCreateFuncs = sRtn

End Function


'-------------------------------------------------------------------------------------------'
Private Function RtnEformCategoryTxtFunc(ByVal oEf As eFormRO) As String
'-------------------------------------------------------------------------------------------'
' function accepts an eform object
' function builds a string made up of js function calls to JSVE function fnSetCategoryText()
' - one function call for each category field on the eform. this is so that the JSVE contains
' both the code and string value for each category response. this is needed for displaying
' in de when adding mimessages
'-------------------------------------------------------------------------------------------'
'Revisions:
'   ic  02/04/2002  copied from v3.0 to v 2.2 for improved 2.2 release
'   DPH 21/10/2002 Reinstated creating list for Radio buttons for RQGs
'   DPH 20/06/2003 - only use active categories
'   ic 17/09/2003 skip hidden controls
'
Dim oCRFElement As eFormElementRO
Dim oCategory As CategoryItem
Dim sRtn As String

    sRtn = ""
    
    'loop through all eform elements
    For Each oCRFElement In oEf.EFormElements
        'ic 17/09/2003 skip hidden controls
        If Not oCRFElement.Hidden Then
            'ic 17/09/2003 skip hidden controls
            Select Case oCRFElement.ControlType
            'ic 17/04/2002 dont need this for radios anymore as we're using custom radios
            ' DPH 21/10/2002 - Reinstated for RQG button controls
            Case ControlType.OptionButtons, ControlType.PushButtons, ControlType.PopUp
    '        Case ControlType.PopUp
                For Each oCategory In oCRFElement.Categories
                    ' DPH 20/06/2003 - only use active categories
                    If oCategory.Active Then
                        sRtn = sRtn & "o1.fnSetCategoryText(" & Chr(34) & oCRFElement.WebId & Chr(34) & "," _
                                                                & Chr(34) & oCategory.Code & Chr(34) & "," _
                                                                & Chr(34) & oCategory.Value & Chr(34) _
                                                          & ");" & vbCrLf
                    End If
                Next
                
            Case Else
            
            End Select
        End If
    Next
    
    RtnEformCategoryTxtFunc = sRtn

End Function


''-------------------------------------------------------------------------------------------'
'Private Function RtnEformElementString(ByVal oStudyDef As StudyDefRO, _
'                                       ByVal oCRFElement As eFormElementRO, _
'                                       ByVal sbgcol As String, _
'                                       ByVal bDisplayNumbers As Boolean, _
'                                       Optional ByVal enEFormUse As eEFormUse = eEFormUse.User) As String
''-------------------------------------------------------------------------------------------'
'' function accepts a studydef object, eform element object, form bg colour
'' function returns a string representing the passed element as vbscript and html: rewrite
'' for v3.0 using business objects
''
'' MLM 26/09/02: Added enEFormUse argument, and use to hide elements on the visit eform that aren't the date.
''   Don't display captions for form or visit dates.
''   Use specific font and positioning for form and visit dates.
''-------------------------------------------------------------------------------------------'
''Revisions:
''ic 28/02/2002 changed user eform picture path to /img/eForm/
''ZA 30/07/2002 added style sheet for font/colour for question/caption/comment
'
'
'    Dim oPic As Picture
'    Dim oCategory As CategoryItem
'    Dim oValidation As Validation
'
'    Dim sRtn As String
'    Dim sCpdImages As String
'    Dim nLength As Integer
'    Dim nWidth As Integer
'    Dim bDefaultFont As Boolean
'    Dim bDefaultCaptionFont As Boolean
'    Dim sStyle As String
'    Dim sStyleRadio As String
'    'MLM 26/09/02:
'    Dim lElementX As Long
'    Dim lElementY As Long
'
'
'    'ZA - 23/07/2002 - check font attributes for caption/comment
'    If oCRFElement.Caption <> "" Then
'        If ((oCRFElement.CaptionFontColour = oStudyDef.FontColour) _
'        And (oCRFElement.CaptionFontSize = oStudyDef.FontSize) _
'        And (oCRFElement.CaptionFontName = oStudyDef.FontName)) _
'        Or ((oCRFElement.CaptionFontColour = 0) _
'        And (oCRFElement.CaptionFontSize = 0) _
'        And (oCRFElement.CaptionFontName = "")) Then
'            bDefaultCaptionFont = True
'        Else
'            bDefaultCaptionFont = False
'        End If
'    End If
'
'    'do this element's font attributes match the study default font attributes, or are empty
'    If ((oCRFElement.FontColour = oStudyDef.FontColour) _
'    And (oCRFElement.FontSize = oStudyDef.FontSize) _
'    And (oCRFElement.FontName = oStudyDef.FontName)) _
'    Or ((oCRFElement.FontColour = 0) _
'    And (oCRFElement.FontSize = 0) _
'    And (oCRFElement.FontName = "")) Then
'        bDefaultFont = True
'    Else
'        bDefaultFont = False
'    End If
'
'
'    'controls are either hidden or visible
'    'MLM 26/09/02: hide user elements on visit eforms
'    If oCRFElement.Hidden Or (enEFormUse = eEFormUse.VisitEForm And oCRFElement.ElementUse = eEFormUse.User) Then
'        'hidden controls are always represented as html hidden fields
'        sRtn = sRtn & "<input " _
'                    & "type=" & Chr(34) & "hidden" & Chr(34) & " " _
'                    & "name=" & Chr(34) & oCRFElement.WebId & Chr(34) _
'                    & ">"
'
'        'ic 24/04/2002 add add info field - needed for hidden fields
'        sRtn = sRtn & "<input " _
'                        & "type=" & Chr(34) & "hidden" & Chr(34) & " " _
'                        & "name=""a" & oCRFElement.WebId & Chr(34) & ">"
'
''ic debug 24/04/2002
''        sRtn = sRtn & "<img src=" & Chr(34) & "../../img/blank.gif" & Chr(34) & " " _
''                        & "name=" & Chr(34) & "img_" & LCase(oCRFElement.Code) & Chr(34) & ">"
''end of debug
'
'    Else
'        'visible controls can be comments,lines,textboxes,attachments,radio buttons,pictures,drop-down lists
'
'        'MLM 26/09/02: Work out the pixel coordinates of the element here
'        If oCRFElement.ElementUse = eElementUse.EFormVisitDate Then
'            lElementY = 0 'in the header bar
'            If enEFormUse = eEFormUse.User Then
'                'form date; position on the right
'                lElementX = 6386 / nIeXSCALE
'            Else
'                'visit date; position on the left
'                lElementX = 2129 / nIeXSCALE
'            End If
'        Else
'            '"normal" elements are moved down a bit
'            lElementY = (oCRFElement.ElementY + 375) / nIeYSCALE
'            lElementX = oCRFElement.ElementX / nIeXSCALE
'        End If
'
'        Select Case oCRFElement.ControlType
'        Case ControlType.Comment
'            'comment
'            If oCRFElement.Caption <> "" Then
'                sRtn = sRtn & "<div style=" & Chr(34) _
'                            & "top:" & lElementY & ";" _
'                            & "left:" & lElementX & ";" _
'                            & "z-index:1" & Chr(34) & ">"
'                'ZA 22/07/2002 - changed FontName with CaptionFontName for CRFEelement
'                If Not bDefaultCaptionFont Then
'                    sRtn = sRtn & "<font "
'                    If (oCRFElement.CaptionFontColour <> 0) Then sRtn = sRtn & "color=" & Chr(34) & GetHexColour(oCRFElement.CaptionFontColour) & Chr(34) & " "
'                    If (oCRFElement.CaptionFontSize <> 0) Then sRtn = sRtn & "size=" & Chr(34) & getFontSize(oCRFElement.CaptionFontSize) & Chr(34) & " "
'                    If (oCRFElement.CaptionFontName <> "") Then sRtn = sRtn & "face=" & Chr(34) & oCRFElement.CaptionFontName & Chr(34)
'                    sRtn = sRtn & ">"
'                End If
'
'                If oCRFElement.CaptionFontBold Then sRtn = sRtn & "<b>"
'                If oCRFElement.CaptionFontItalic Then sRtn = sRtn & "<i>"
'                sRtn = sRtn & Replace(oCRFElement.Caption, vbCrLf, "<br>")
'                If oCRFElement.CaptionFontItalic Then sRtn = sRtn & "</i>"
'                If oCRFElement.CaptionFontBold Then sRtn = sRtn & "</b>"
'
'                If Not bDefaultFont Then
'                    sRtn = sRtn & "</font>"
'                End If
'
'                sRtn = sRtn & "</div>"
'            End If
'
'        Case ControlType.Line
'            'line
'            sRtn = sRtn & "<div " _
'                        & "style=" & Chr(34) & "20px;left:0;top:" & lElementY & Chr(34) _
'                        & ">" _
'                        & "<hr>" _
'                        & "</div>"
'
'
'        Case ControlType.TextBox, ControlType.Calendar, ControlType.RichTextBox
'            'textbox
'            'MLM 26/09/02: Don't draw captions for form and visit dates.
'            If oCRFElement.Caption <> "" And oCRFElement.ElementUse = eElementUse.User Then
'                sRtn = sRtn & RtnCaption(oCRFElement, bDefaultCaptionFont, bDisplayNumbers, oStudyDef.FontColour)
'            End If
'
'
'            SetElementWidth oCRFElement.Format, oCRFElement.QuestionLength, nLength, nWidth
'
'
'            sRtn = sRtn & "<div id=" & Chr(34) & oCRFElement.WebId & "_inpDiv" & Chr(34) & " " _
'                        & "style=" & Chr(34) & ";left:" & lElementX _
'                        & ";top:" & lElementY _
'                        & ";z-index:2" & Chr(34) & ">"
'
'
'
'            sRtn = sRtn & "<input " _
'                        & "type=" & Chr(34) & "hidden" & Chr(34) & " " _
'                        & "name=""a" & oCRFElement.WebId & Chr(34) & ">"
'
'            'set element style
'            sStyle = Chr(34) & "background-color:#EEEEEE; "
'            If Not bDefaultFont Then
'                sStyle = sStyle & SetFontAttributes(oCRFElement, True)
'            End If
'            sStyle = sStyle & Chr(34)
'
'
'            sRtn = sRtn & "<input type=" & Chr(34) & "text" & Chr(34) & " " _
'                        & "disabled " _
'                        & "tabindex=" & Chr(34) & oCRFElement.ElementOrder & Chr(34) & " " _
'                        & "name=" & Chr(34) & oCRFElement.WebId & Chr(34) & " " _
'                        & "Style=" & sStyle & " " _
'                        & "size=" & Chr(34) & nWidth & Chr(34) & " " _
'                        & "maxlength=" & Chr(34) & nLength & Chr(34) & " " _
'                        & "onfocus=" & Chr(34) & "frm1.fnGotFocus(" & "'" & oCRFElement.WebId & "');" & Chr(34) & ">"
'
'                        '& "frm1.fnDisplayFieldTxt(" & "deFrm.af_" & LCase(oCRFElement.Code) & ",'" & oCRFElement.Name & sBlockDelimeter & oCRFElement.Format & sBlockDelimeter & oCRFElement.QuestionLength & sBlockDelimeter & Replace(oCRFElement.Helptext, "'", "\'") & "');" & Chr(34) & ">"
'
'
'
'            sRtn = sRtn & "<img src=" & Chr(34) & "../../img/blank.gif" & Chr(34) & " " _
'                        & "name=" & Chr(34) & "imgc_" & LCase(oCRFElement.Code) & Chr(34) & ">"
'
'            sRtn = sRtn & "<img src=" & Chr(34) & "../../img/blank.gif" & Chr(34) & " " _
'                        & "name=" & Chr(34) & "img_" & LCase(oCRFElement.Code) & Chr(34) & ">"
'
'            sRtn = sRtn & "</div>"
'
'        Case ControlType.Attachment
'            'multimedia
'
'            If oCRFElement.Caption <> "" Then
'                sRtn = sRtn & RtnCaption(oCRFElement, bDefaultCaptionFont, bDisplayNumbers, oStudyDef.FontColour)
'            End If
'
'            sRtn = sRtn & "<input " _
'                        & "type=" & Chr(34) & "hidden" & Chr(34) & " " _
'                        & "name=""a" & oCRFElement.WebId & Chr(34) & ">"
'
'            sRtn = sRtn & "<div id=" & Chr(34) & oCRFElement.WebId & "_inpDiv" & Chr(34) & " " _
'                        & "style=" & Chr(34) & "left:" & lElementX _
'                        & ";top:" & lElementY _
'                        & ";z-index:2" & Chr(34) & ">"
'
'            sRtn = sRtn & "<input type=" & Chr(34) & "file" & Chr(34) & " " _
'                        & "disabled " _
'                        & "tabindex=" & Chr(34) & oCRFElement.ElementOrder & Chr(34) & " " _
'                        & "name=" & Chr(34) & oCRFElement.WebId & Chr(34) & " " _
'                        & "style=" & Chr(34) & "background-color:#EEEEEE; " & Chr(34) & " " _
'                        & "onfocus=" & Chr(34) & "frm1.fnGotFocus(" & "'" & oCRFElement.WebId & "');" & Chr(34) & ">"
'
'                        '& "frm1.fnDisplayFieldTxt(" & "deFrm.af_" & LCase(oCRFElement.Code) & ",'" & oCRFElement.Name & sBlockDelimeter & oCRFElement.Format & sBlockDelimeter & oCRFElement.QuestionLength & sBlockDelimeter & Replace(oCRFElement.Helptext, "'", "\'") & "');" & Chr(34) & ">"
'
'
'
'            sRtn = sRtn & "<img src=" & Chr(34) & "../../img/blank.gif" & Chr(34) & " " _
'                         & "name=" & Chr(34) & "imgc_" & LCase(oCRFElement.Code) & Chr(34) & ">"
'
'            sRtn = sRtn & "<img src=" & Chr(34) & "../../img/blank.gif" & Chr(34) & " " _
'                         & "name=" & Chr(34) & "img_" & LCase(oCRFElement.Code) & Chr(34) & ">"
'
'            sRtn = sRtn & "</div>"
'
'
'        Case ControlType.OptionButtons, ControlType.PushButtons
'            'radio button group (1 or more radio buttons)
'           '''''''''''''''''''''''''''''''''''''''''''''''''' sStyleRadio = "Style = """
'            'outer div enclosing all radio buttons and captions in the group
'            sRtn = sRtn & "<div " _
'                        & "style=" & Chr(34) & ";left:" & lElementX & ";" _
'                        & "top:" & lElementY & ";" _
'                        & "z-index:2;" & Chr(34) & ">" & vbCrLf
'
'
''            sRtn = sRtn & "<div id=" & Chr(34) & "f_" & LCase(oCRFElement.Code) & "_capDiv" & Chr(34) & " " _
''                        & "style=" & Chr(34) & "top:0;left:0;" & Chr(34) & ">"
'
'            sRtn = sRtn & "<input " _
'                        & "type=" & Chr(34) & "hidden" & Chr(34) & " " _
'                        & "name=""a" & oCRFElement.WebId & Chr(34) & ">" & vbCrLf
'
'            sRtn = sRtn & "<input " _
'                        & "type=" & Chr(34) & "hidden" & Chr(34) & " " _
'                        & "name=""a" & oCRFElement.WebId & Chr(34) & ">" & vbCrLf
'
''''''''            If Not bDefaultFont Then
''''''''                If (oCRFElement.FontColour <> 0) Then sStyleRadio = sStyleRadio & "COLOR: " & GetHexColour(oCRFElement.FontColour) & ";"
''''''''                If (oCRFElement.FontSize <> 0) Then sStyleRadio = sStyleRadio & "FONT-SIZE: " & oCRFElement.FontSize & " pt;"
''''''''                If (oCRFElement.FontName <> "") Then sStyleRadio = sStyleRadio & "FONT-FAMILY: " & oCRFElement.FontName & ";"
''''''''                If oCRFElement.FontBold Then sStyleRadio = sStyleRadio & "FONT-WEIGHT: bold; "
''''''''                If oCRFElement.FontItalic Then sStyleRadio = sStyleRadio & "FONT-STYLE: italic;" & Chr(34)
''''''''            End If
'
''''''''            If Not bDefaultFont Then
''''''''                sStyleRadio = """" & SetFontAttributes(oCRFElement) & """"
''''''''            End If
'
'            ''''''''''''''sRtn = sRtn & "<table><tr><td Style = " & sStyleRadio & ">"
'            sRtn = sRtn & "<table><tr><td>"
'
'            'div positioning radio buttons
'            sRtn = sRtn & "<div id=" & Chr(34) & oCRFElement.WebId & "_inpDiv" & Chr(34) & " " _
'                        & "style=" & Chr(34) & "position:relative;z-index:2;"
'
'
'
'
'
'            sRtn = sRtn & Chr(34) & ">"
'
'
'            'sRtn = sRtn & "<table><tr><td onclick=" & Chr(34) & "frm1.fnLostFocus(" & "'f_" & LCase(oCRFElement.Code) & "','#FFFF80');" & Chr(34) & ">"
''            sRtn = sRtn & "<table><tr><td>"
''
''
''            For Each oCategory In oCRFElement.Categories
''                sRtn = sRtn & "<input type=" & Chr(34) & "radio" & Chr(34) & " " _
''                            & "disabled " _
''                            & "tabindex=" & Chr(34) & oCRFElement.ElementOrder & Chr(34) & " " _
''                            & "name=" & Chr(34) & "f_" & LCase(oCRFElement.Code) & Chr(34) & " " _
''                            & "value=" & Chr(34) & oCategory.Code & Chr(34) & " " _
''                            & "onfocus=" & Chr(34) & "frm1.fnGotFocus(" & "'f_" & LCase(oCRFElement.Code) & "','" & oCategory.Code & "');" & Chr(34) & ">"
''
''                            '& "frm1.fnDisplayFieldTxt(" & "deFrm.af_" & LCase(oCRFElement.Code) & ",'" & oCRFElement.Name & sBlockDelimeter & oCRFElement.Format & sBlockDelimeter & sBlockDelimeter & Replace(oCRFElement.Helptext, "'", "\'") & "');" & Chr(34) & " " _
''
''                If Not bDefaultFont Then
''                    sRtn = sRtn & "<font "
''                    If (oCRFElement.FontColour <> 0) Then sRtn = sRtn & "color=" & Chr(34) & GetHexColour(oCRFElement.FontColour) & Chr(34) & " "
''                    If (oCRFElement.FontSize <> 0) Then sRtn = sRtn & "size=" & Chr(34) & getFontSize(oCRFElement.FontSize) & Chr(34) & " "
''                    If (oCRFElement.FontName <> "") Then sRtn = sRtn & "face=" & Chr(34) & oCRFElement.FontName & Chr(34)
''                    sRtn = sRtn & ">"
''
''                    sRtn = sRtn & oCategory.Value & "</font><br>"
''                Else
''                    sRtn = sRtn & oCategory.Value & "<br>"
''                End If
''
''            Next
''
''
''
''            'if only one radio button, add a hidden field of the same name to create a control array
''            If oCRFElement.Categories.Count = 1 Then
''                sRtn = sRtn & "<input type=" & Chr(34) & "hidden" & Chr(34) & " disabled name=" & Chr(34) & "f_" & LCase(oCRFElement.Code) & Chr(34) & ">"
''            End If
''
''            sRtn = sRtn & "</td><td valign=top>"
'
'
'
'
''            sRtn = sRtn & "</td></tr></table>"
'
'            sRtn = sRtn & "</div></td><td valign=" & Chr(34) & "top" & Chr(34) & ">"
'
'
'            sRtn = sRtn & "<img src=" & Chr(34) & "../../img/blank.gif" & Chr(34) & " " _
'                         & "name=" & Chr(34) & "imgc_" & LCase(oCRFElement.Code) & Chr(34) & ">"
'
'            sRtn = sRtn & "<img src=" & Chr(34) & "../../img/blank.gif" & Chr(34) & " " _
'                         & "name=" & Chr(34) & "img_" & LCase(oCRFElement.Code) & Chr(34) & ">"
'
'            sRtn = sRtn & "</td></tr></table></div>" & vbCrLf
'
'
'
'            If oCRFElement.Caption <> "" Then
'                sRtn = sRtn & RtnCaption(oCRFElement, bDefaultCaptionFont, bDisplayNumbers, oStudyDef.FontColour)
'            End If
'
'
'
'        Case ControlType.Picture
'            'picture
'
'        'if the element  cannot be found it is omitted
'        If Dir(gsDOCUMENTS_PATH & oCRFElement.Caption) <> "" Then
'            Set oPic = LoadPicture(gsDOCUMENTS_PATH & oCRFElement.Caption)
'            nWidth = frmGenerateHTML.ScaleY(oPic.Width, vbHimetric, vbPixels) * nXSCALE
'            Set oPic = Nothing
'
'            If nWidth > 7000 And nWidth < 5000 Then 'Assume it is a line, limit to width of page
'
'                nLength = 611
'                sRtn = sRtn & "<div id=" & Chr(34) & "lin" & oCRFElement.ElementID & "_inpDiv" & Chr(34) & " " _
'                            & "style=" & Chr(34) & ";left:" & lElementX & ";" _
'                            & "top:" & lElementY & Chr(34) & ">"
'
'                'ic 28/02/2002
'                'changed user eform picture path to /img/eForm/
'                sRtn = sRtn & "<img src=" & Chr(34) & "../../img/eForm/" & oCRFElement.Caption & Chr(34) & " " _
'                            & "width=" & Chr(34) & nLength & Chr(34) & " " _
'                            & "valign=" & Chr(34) & "top" & Chr(34) & " " _
'                            & "align=" & Chr(34) & "left" & Chr(34) & ">"
'                sRtn = sRtn & "</div>"
'
'            Else
'                sRtn = sRtn & "<div id=" & Chr(34) & "pic" & oCRFElement.ElementID & "_inpDiv" & Chr(34) & " " _
'                            & "style=" & Chr(34) & ";left:" & lElementX & ";" _
'                            & "top:" & lElementY & Chr(34) & ">"
'
'                'ic 28/02/2002
'                'changed user eform picture path to /img/eForm/
'                sRtn = sRtn & "<img src=" & Chr(34) & "../../img/eForm/" & oCRFElement.Caption & Chr(34) & " " _
'                            & "valign=" & Chr(34) & "top" & Chr(34) & " " _
'                            & "align=" & Chr(34) & "left" & Chr(34) & ">"
'                sRtn = sRtn & "</div>"
'            End If
'
'
'            'check we havent already copied this file by looking for the filename in our copied images variable
'            'which is just a comma delimited list of files we add to as we process this function
'            If InStr(sCpdImages, "," & oCRFElement.Caption) = 0 Then
'
'                'if the file exists in the html folder already, delete it
'                'ic 28/02/2002
'                'changed user eform picture path to /img/eForm/
'                If Dir(gsWEB_HTML_LOCATION & "img\eform\" & oCRFElement.Caption) <> "" Then Kill (gsWEB_HTML_LOCATION & "img\eForm\" & oCRFElement.Caption)
'
'                'copy the file from source to destination
'                'ic 28/02/2002
'                'changed user eform picture path to /img/eForm/
'                FileCopy gsDOCUMENTS_PATH & oCRFElement.Caption, gsWEB_HTML_LOCATION & "img\eForm\" & oCRFElement.Caption
'
'                'add to our list of copied images
'                sCpdImages = sCpdImages & "," & oCRFElement.Caption
'            End If
'
'        Else
''                   MsgBox "Image " & gsDOCUMENTS_PATH & oCrfElement.Caption & " cannot be found!", vbOKOnly
'        End If
'
'        Case ControlType.PopUp
'            'drop-down list
'
'
'            sRtn = sRtn & "<div id=" & Chr(34) & oCRFElement.WebId & "_inpDiv" & Chr(34) & " " _
'                        & "style=" & Chr(34) & ";left:" & lElementX & ";" _
'                        & "top:" & lElementY & ";" _
'                        & "z-index:2" & Chr(34) & ">"
'
'
'            sRtn = sRtn & "<input " _
'                        & "type=" & Chr(34) & "hidden" & Chr(34) & " " _
'                        & "name=""a" & oCRFElement.WebId & Chr(34) & ">"
'
'            'set element style
'            sStyle = Chr(34) & "background-color:#EEEEEE; "
'            If Not bDefaultFont Then
'                sStyle = sStyle & SetFontAttributes(oCRFElement, True)
'            End If
'            sStyle = sStyle & Chr(34)
'
'
'            sRtn = sRtn & "<table><tr>" _
'                        & "<td id=" & Chr(34) & "tbl_" & LCase(oCRFElement.Code) & Chr(34) & ">"
'
'            'onblur event so that answers are evaluated when tabbing between fields as well as clicking
'            sRtn = sRtn & "<select tabindex=" & Chr(34) & oCRFElement.ElementOrder & Chr(34) & " " _
'                            & "disabled " _
'                            & "name=" & Chr(34) & oCRFElement.WebId & Chr(34) & " " _
'                            & "Style=" & sStyle & " " _
'                            & "onfocus=" & Chr(34) & "frm1.fnGotFocus('" & oCRFElement.WebId & "');" & Chr(34) & " " _
'                            & "onchange=" & Chr(34) & "frm1.fnLostFocus('" & oCRFElement.WebId & "');" & Chr(34) & ">"
'
'                            '& "frm1.fnDisplayFieldTxt(" & "deFrm.af_" & LCase(oCRFElement.Code) & ",'" & oCRFElement.Name & sBlockDelimeter & oCRFElement.Format & sBlockDelimeter & sBlockDelimeter & Replace(oCRFElement.Helptext, "'", "\'") & "');" & Chr(34) & " " _
'
'            'space as default in pop-up box
'            sRtn = sRtn & "<option value=" & Chr(34) & Chr(34) & "> </option>"
'
'
'            For Each oCategory In oCRFElement.Categories
'                sRtn = sRtn & "<option value=" & Chr(34) & oCategory.Code & Chr(34) & ">" _
'                            & oCategory.Value _
'                            & "</option>"
'            Next
'
'
'
'        sRtn = sRtn & "</select>"
'
'        sRtn = sRtn & "</td><td>"
'
'        sRtn = sRtn & "<img src=" & Chr(34) & "../../img/blank.gif" & Chr(34) & " " _
'                         & "name=" & Chr(34) & "imgn_" & LCase(oCRFElement.Code) & Chr(34) & ">"
'
'        sRtn = sRtn & "<img src=" & Chr(34) & "../../img/blank.gif" & Chr(34) & " " _
'                         & "name=" & Chr(34) & "imgc_" & LCase(oCRFElement.Code) & Chr(34) & ">"
'
'        sRtn = sRtn & "<img src=" & Chr(34) & "../../img/blank.gif" & Chr(34) & " " _
'                         & "name=" & Chr(34) & "img_" & LCase(oCRFElement.Code) & Chr(34) & ">"
'
'        sRtn = sRtn & "</td></tr></table>"
'
'        sRtn = sRtn & "</div>"
'
'        If oCRFElement.Caption <> "" Then
'            sRtn = sRtn & RtnCaption(oCRFElement, bDefaultCaptionFont, bDisplayNumbers, oStudyDef.FontColour)
'        End If
'
'        Case Else
'        End Select
'    End If
'
'    Set oCategory = Nothing
'    RtnEformElementString = sRtn
'End Function


'-------------------------------------------------------------------------------------------'
Private Function RtnCaption(ByVal oCRFElement As eFormElementRO, _
                            ByVal bDefaultFont As Boolean, _
                            ByVal bDisplayNumbers As Boolean, _
                            ByVal lDefFontCol As Long, _
                            ByVal neFormHeaderHeight As Integer, _
                            Optional ByVal bRQGHeader As Boolean = False, _
                            Optional ByRef enElementUse As eElementUse = eElementUse.User) As String
'-------------------------------------------------------------------------------------------'
' function accepts an eform element object
' function returns the html required to display the passed elements caption : rewrite for
' v3.0 using business objects
'
' MLM 26/06/03: Pass in enElementUse and use in positioning of captions for form and visit dates.
'-------------------------------------------------------------------------------------------'
'Revisions:
'   ic  02/04/2002  copied from v3.0 to v 2.2 for improved 2.2 release
'   DPH 11/12/2002  Want caption without <DIV> for RQG
'   DPH 08/05/2003 - Change Y offset to 750 to allow for eForm header
'   DPH 12/05/2003 - Set caption left/width position
'   DPH 28/05/2003 - include eformheaderheight as parameter
    Dim sCap As String
    Dim lCapLeft As Long
    Dim lCapWidth As Long
    Dim lCapTop As Long
    
    If Not bRQGHeader Then
        sCap = "<div id=" & Chr(34) & oCRFElement.WebId & "_"
    
        'if this is a radio button caption then add an 'r' to the div id so that the
        'proper div can surround the entire group
    '    If (oCRFElement.ControlType = ControlType.OptionButtons) _
    '    Or (oCRFElement.ControlType = ControlType.PushButtons) Then
    '       sCap = sCap & "r"
    '    End If
        
        ' DPH 12/05/2003 - calculate caption left position & set
        'MLM 26/06/03: Absolute positioning of captions for form and visit dates
        If oCRFElement.ElementUse = eElementUse.EFormVisitDate Then
            If enElementUse = eElementUse.EFormVisitDate Then
                lCapLeft = 45
            Else
                lCapLeft = 4257
            End If
            lCapTop = 435 '405
            lCapWidth = 0
        ElseIf oCRFElement.CaptionX > 0 Then
            lCapLeft = oCRFElement.CaptionX
            lCapTop = oCRFElement.CaptionY + neFormHeaderHeight
            lCapWidth = 0
        Else
            'lCapLeft = oCRFElement.ElementX - .Width - 50
            lCapLeft = 0
            lCapTop = oCRFElement.CaptionY + neFormHeaderHeight
            lCapWidth = oCRFElement.ElementX - 50
        End If
        
        ' DPH 08/05/2003 - Change Y offset to 750 to allow for eForm header
        sCap = sCap & "CapDiv" & Chr(34) & " " _
                    & "style = " & Chr(34) _
                    & "top:" & CLng(lCapTop / nIeYSCALE) & ";" _
                    & "left:" & CLng(lCapLeft / nIeXSCALE) & ";"
        If lCapWidth > 0 Then
            sCap = sCap & "width:" & CLng(lCapWidth / nIeXSCALE) & ";"
        End If
        sCap = sCap & "z-index:1" & Chr(34) & ">"
    
    Else
        sCap = ""
    End If
    
    'MLM 26/06/03: Fixed style for form and visit date captions
    If oCRFElement.ElementUse = eElementUse.User Then
        'ic added extra conditions for 2.2 version as business objects dont fill in all defaults
        'ZA 22/07/2002 - updated CRFElement font/colour properties to use Caption properties
        sCap = sCap & "<font style = "
        
        'ZA 30/07/2002 -  added style sheet
        
        If (oCRFElement.CaptionFontColour = 0) Then
        sCap = sCap & Chr(34) & "color:" & RtnHTMLCol(lDefFontCol) & Chr(34) & "; "
        Else
            sCap = sCap & Chr(34) & "color:" & RtnHTMLCol(oCRFElement.CaptionFontColour) & "; "
        End If
        
        If Not bDefaultFont Then
            If (oCRFElement.CaptionFontSize <> 0) Then sCap = sCap & "FONT-SIZE: " & oCRFElement.CaptionFontSize & " pt; "
            If (oCRFElement.CaptionFontName <> "") Then sCap = sCap & "FONT-FAMILY: " & oCRFElement.CaptionFontName & ";"
        End If
        sCap = sCap & Chr(34) & ">"
        
        If oCRFElement.CaptionFontBold Then sCap = sCap & "<b>"
        If oCRFElement.CaptionFontItalic Then sCap = sCap & "<i>"
        If bDisplayNumbers Then sCap = sCap & oCRFElement.ElementOrder & "."
        sCap = sCap & Replace(oCRFElement.Caption, vbCrLf, "<br>")
        
        'ic 27/04/2002
        'added unit to caption
        If oCRFElement.Unit <> "" Then sCap = sCap & " (" & oCRFElement.Unit & ")"
        
        If oCRFElement.CaptionFontItalic Then sCap = sCap & "</i>"
        If oCRFElement.CaptionFontBold Then sCap = sCap & "</b>"
    
        'If Not bDefaultFont Then
        sCap = sCap & "</font>"
    Else
        sCap = sCap & "<font style=""color:#AAAAAA;font-family:Verdana,helvetica,arial; font-size:8pt;"">" & _
            oCRFElement.Caption & "</font>"
    End If
    
    If Not bRQGHeader Then
        sCap = sCap & "</div>" '& vbCrLf
    Else
        sCap = sCap '& vbCrLf
    End If
    
    RtnCaption = sCap
End Function

'ic 02/04/2002 commented out, replaced with PublishEFormASPFiles()
''-------------------------------------------------------------------------------------------'
'Public Sub CreateHTMLFiles(ByVal vClinicalTrialId As Long, _
'                           ByVal vVersionId As Integer)
''-------------------------------------------------------------------------------------------'
'' Creates ASP eform displayed in the browser.
''-------------------------------------------------------------------------------------------'
''Revisions:
''ic  21/08/01   rewrite for v2.2
''rem 19/10/01 bug fix 88: removed stylesheet link from eform as it was setting background colour of eform
''rem 22/10/01 added "style=z-inex:100" to <div> tag on Message boxes so that it appeared in front of text boxes on eform
''DPH 29/10/2001 Fixes to not show active categories and fix question number display problem
''DPH 5/11/2001 - Order Categories to ValueOrder
''REM 5/11/2001 - bug fix 122: added call to FinishoALM VB function to close Arezzo instance after saving an eform
'
'Dim rsStudyDefinition As ADODB.Recordset
'Dim rsCRFPages As ADODB.Recordset
'Dim rsCRFElement As ADODB.Recordset
'Dim rsDataItem As ADODB.Recordset
'Dim rsDerivedValues As ADODB.Recordset
'Dim sSQL As String
'
'Dim sPath As String
'Dim sFile As String
'Dim sExt As String
'Dim nFile As Integer
'
'Dim lCRFPageId As Long
'Dim lDataItemId As Long
'Dim rsValueData As ADODB.Recordset
'Dim sDefaultFontAttr As String
'Dim sPositionAttr As String
'Dim sDisabledSetting As String
'Dim bNetscape As Boolean
'Dim nRCount As Integer 'recordcount
'Dim bNoOfelements As Boolean
'Dim bNoOfpages As Boolean
'Dim lControlLeft As Long
'Dim lImageXposition As Long
'Dim bDispNum As Boolean
'Dim bCommentDispNum As Boolean  ' DPH 29/10/2001 Added because Comments causing question number display problem
'
'Dim nLengthOfItem As Integer     'JL 16/06/00 Added for use by radios & pop-ups
'Dim nMaxLengthSoFar As Integer   'JL 16/06/00 Added for use by radios & pop-ups
'Dim nTextBoxCharWidth As Integer
'Dim nMaxLength As Integer
'Dim sComments As String
'Dim sUnitOfMeasurement As String
'
'Dim nDefaultFontSize As Integer
'Dim lDefaultFontColour As Long
'Dim sDefaultFontFamily As String
'Dim nDefaultFontWeight As Integer   'Bold/medium 'Default 0=medium, 1=bold
'Dim nDefaultFontStyle As Integer    'Default 0=normal, 1=Italic
'
'Dim nNumberOfSkips As Integer
'Dim nNumberOfRadios As Integer
'
'Dim sQuestionRoleCode As String 'For Question
'Dim sDerivation As String
'
'Const sDISABLED = "DISABLED=""Disabled"""
'Const sENABLED = ""
'
'Const bIEflag = 0
'Const nMINIMALCONTROLWIDTH = 70
'Const nMINIMAL_POPUP_WIDTH = 40
'Dim ElementWidth As Integer
'Dim oPicture As Picture
'Dim nMaxImageLength As Integer
'
'
''ic 05/12/00
''added declaration to hold list of images files already copied
'Dim sCpdImages As String
'
'
'Dim sBgCol As String
'
'    'On Error GoTo ErrHandler
'
'    bNetscape = False
'
'
'
'
'    nFile = FreeFile
'    bNoOfelements = True
'    bNoOfpages = True
'
'
'    'get Default font attributes
'    sSQL = "SELECT * FROM StudyDefinition" _
'         & " WHERE ClinicalTrialId = " & vClinicalTrialId _
'         & " AND VersionId = " & vVersionId
'
'    Set rsStudyDefinition = New ADODB.Recordset
'    rsStudyDefinition.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
'
'    If rsStudyDefinition.RecordCount <> 0 Then
'        sDefaultFontAttr = "<style type=" & Chr(34) & "text/css" & Chr(34) & "> " _
'                         & "div{position:absolute;width:auto;height:auto}" _
'                         & "</style>"
'
'
'
''        sDefaultFontAttr = "<style type=" & Chr$(34) & "text/css" & Chr$(34) & "> " & _
''            "div    {position:absolute;width:auto;height:auto}" & _
''            "p      {color:" & GetFontColour(rsStudyDefinition!DefaultFontColour) & ";font-family:" & _
''                    rsStudyDefinition!DefaultFontName & ";font-size:" & rsStudyDefinition!DefaultFontSize & _
''                    " " & "pt}  " & _
''            "input.T { font-size: 10pt;}" & _
''                    "</style>"
'
'     End If
'
'     nDefaultFontSize = rsStudyDefinition!DefaultFontSize
'     lDefaultFontColour = rsStudyDefinition!DefaultFontColour
'     sDefaultFontFamily = rsStudyDefinition!DefaultFontName
'     nDefaultFontWeight = rsStudyDefinition!DefaultFontBold
'     nDefaultFontStyle = rsStudyDefinition!DefaultFontItalic
'
'
'
'    'Identify a trial
'    sSQL = "SELECT * FROM CRFPage " _
'         & " WHERE ClinicalTrialId = " & vClinicalTrialId _
'         & " AND VersionId = " & vVersionId
'
'    Set rsCRFPages = New ADODB.Recordset
'    rsCRFPages.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
'
'
'
'
'
'
'    Do While Not rsCRFPages.EOF
'
'        bDispNum = CBool(rsCRFPages!displayNumbers)
'        lCRFPageId = rsCRFPages!CRFPageId
'
'
'        sSQL = "SELECT * FROM CRFElement " _
'             & "WHERE CRFPageId =" & lCRFPageId & " " _
'             & "AND ClinicalTrialId =" & vClinicalTrialId & " " _
'             & "AND VersionId=" & vVersionId & " " _
'             & "ORDER BY CRFElementId"
'
'        Set rsCRFElement = New ADODB.Recordset
'        rsCRFElement.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
'
'
'        sPath = gsWEB_HTML_LOCATION
'        sFile = LTrim(str(rsCRFPages!ClinicalTrialId)) _
'              & "_" & LTrim(str(rsCRFPages!CRFPageId)) _
'              & "_" & gUser.DatabaseName
'        sExt = ".asp"
'
'        ' DPH 17/10/2001 Make sure folder exists before opening
'        If Not FolderExistence(sPath & sFile & sExt) Then
'            Call DialogError("HTML files could not be created because " & gsWEB_HTML_LOCATION & " not found")
'            Exit Sub
'        End If
'
'        Open sPath & sFile & sExt For Output As #nFile
'
'
'        'start of asp page
'        'page header,include files
'        Print #nFile, "<%@ LANGUAGE=VBScript%>" & vbCrLf _
'                    & "<%Response.Buffer = true%>" & vbCrLf _
'                    & "<%Response.Expires=0%>" & vbCrLf _
'                    & "<!-- #include file=" & Chr(34) & "../general/include/checkSSL.asp" & Chr(34) & " -->" & vbCrLf _
'                    & "<!-- #include file=" & Chr(34) & "../general/include/checkLoggedIn.asp" & Chr(34) & "-->" & vbCrLf _
'                    & "<!-- #include file=" & Chr(34) & "../general/include/miscFunc.asp" & Chr(34) & "-->" & vbCrLf
'
'        Print #nFile, "<%" & vbCrLf _
'                    & "Response.Write(" & Chr(34) & "<!-- " & Chr(34) & " & application(" & Chr(34) & "HTMLCOPYRIGHT" & Chr(34) & ") & " & Chr(34); " -->" & Chr(34) & ")" & vbCrLf
'
'
'        'dim statements
'        Print #nFile, "dim oIo" & vbCrLf _
'                    & "dim usrId" & vbCrLf _
'                    & "dim usrRl" & vbCrLf _
'                    & "dim fltDb" & vbCrLf _
'                    & "dim fltSt" & vbCrLf _
'                    & "dim fltSi" & vbCrLf _
'                    & "dim fltSj" & vbCrLf _
'                    & "dim fltId" & vbCrLf _
'                    & "dim nxt" & vbCrLf _
'                    & "dim saveFrm" & vbCrLf _
'                    & "dim nSaveOk" & vbCrLf _
'                    & "dim nStatus" & vbCrLf
'
'
'        'assign passed/session var values to local vars
'        Print #nFile, "usrId = session(" & Chr(34) & "usrId" & Chr(34) & ")" & vbCrLf _
'                    & "usrRl = session(" & Chr(34) & "usrRl" & Chr(34) & ")" & vbCrLf _
'                    & "fltDb = " & Chr(34) & gUser.DatabaseName & Chr(34) & vbCrLf _
'                    & "fltSt = " & Chr(34) & LTrim(str(rsCRFPages!ClinicalTrialId)) & Chr(34) & vbCrLf _
'                    & "fltSi = Request.QueryString(" & Chr(34) & "fltSi" & Chr(34) & ")" & vbCrLf _
'                    & "fltSj = Request.QueryString(" & Chr(34) & "fltSj" & Chr(34) & ")" & vbCrLf _
'                    & "fltId = Request.QueryString(" & Chr(34) & "fltId" & Chr(34) & ")" & vbCrLf
'
'
'        'create i/o object,check de permission 'if' condition
'        '<-- start of de permission 'if' condition
'        Print #nFile, "set oIo = Server.CreateObject(" & Chr(34) & "MACROWWWIO22.clsWWW" & Chr(34) & ")" & vbCrLf _
'                    & "If oIo.Permission(usrId,fltDb,usrRl," & Chr(34) & "F5002" & Chr(34) & ") Then" & vbCrLf
'
'        Print #nFile, "if thisSubjectLocked() then"
'
'        'get 'nxt' value from hidden field. this field is in all de eforms and its name is the same as
'        'the eforms file name, prefixed with 'f_' (because field names must begin with a letter) without '.asp' extension
'        'if a value for this field is present in the form object it means that the page has submitted to itself.
'        'the field will contain an integer value. 0=reload this eform, 1=load next eform, 2=load previous eform
'        Print #nFile, "nxt = Request.Form(" & Chr(34) & "f_" & sFile & Chr(34) & ")" & vbCrLf
'
'
'        'set nStatus = 0. nStatus is used in the page to determine the content.
'        Print #nFile, "nStatus = 0"
'
'
'        'if 'nxt' has a value the form has submitted to itself and should be saved
'        'call SaveForm() function. if function returns 1, save was ok, else save of one or more responses failed on
'        'validation - set nStatus to 1 to display this error when the page is displayed in the browser.
'        'if save was ok and nxt value is 1 or 2 (next or previous eform) then get the url of the requested eform from
'        'MoveNext() function
'        Print #nFile, "If nxt <> " & Chr(34) & Chr(34) & " Then" & vbCrLf _
'                    & "saveFrm = Request.Form()" & vbCrLf _
'                    & "nSaveOk = oIo.SaveForm(usrId,fltDb,fltSi,fltSt,fltSj,fltId,saveFrm)" & vbCrLf _
'                    & "set session(" & Chr(34) & "usrStudyDef" & Chr(34) & ") = oIo.RtnStudyDef()" & vbCrLf _
'                    & "session(" & Chr(34) & "usrStudyDefOK" & Chr(34) & ") = " & Chr(34) & "True" & Chr(34) & vbCrLf _
'                    & "If (nSaveOk = 1) Then" & vbCrLf _
'                    & "Select Case (nxt)" & vbCrLf _
'                    & "Case 0:" & vbCrLf _
'                    & "Case 1,2: url = oIo.MoveNext(usrId,fltDb,fltSi,fltSt,fltSj,fltId,nxt)" & vbCrLf _
'                    & "Case 3: nStatus = 4" & vbCrLf _
'                    & "End Select" & vbCrLf _
'                    & "Else" & vbCrLf _
'                    & "nStatus = 1" & vbCrLf _
'                    & "End If" & vbCrLf _
'                    & "End If" & vbCrLf
'
'
'        'if url variable is empty load this page (either first visit, save failed or user requested to return to this page
'        'if url has a value it is a pipe (|) delimited string containing 2 values. first value is next non-repeating eform,
'        'second value is next repeating eform. if both values are empty this is the last (or if previous eform was requested,
'        'the first) eform in the visit. set nStatus to 2. if the values are different there are several possibilities:
'        'val1 is empty, val2 has string - there are only repeating eforms remaining in this visit
'        'val1 has string, val2 is empty - there are only non repeating eforms remaining in this visit
'        'val1 has string, val2 has string - there are repeating and non repeating eforms remaining in this visit
'        'set nStatus to 3
'        Print #nFile, "If (url <> " & Chr(34) & Chr(34) & ") Then" & vbCrLf _
'                    & "url = split(url," & Chr(34) & "|" & Chr(34) & ")" & vbCrLf _
'                    & "If (url(0) & url(1) = " & Chr(34) & Chr(34) & ") Then" & vbCrLf _
'                    & "nStatus = 2" & vbCrLf _
'                    & "ElseIf (url(0) <> url(1)) Then" & vbCrLf _
'                    & "nStatus = 3" & vbCrLf _
'                    & "Else" & vbCrLf _
'                    & "Response.Redirect(url(0))" & vbCrLf _
'                    & "End If" & vbCrLf _
'                    & "End If" & vbCrLf
'
'
'        'branch on nStatus value
'        'first branch:
'        'nStatus is 0 - first visit to this page, display with no messages
'        'nStatus is 1 - save of one or more responses failed. display this page, but add an alert in js pageLoaded() function
'        Print #nFile, "Select Case (nStatus)" & vbCrLf _
'                    & "Case 0,1:" & vbCrLf _
'                    & "%>"
'
'
''        Print #nFile, "<html>" & vbCrLf _
''                    & "<head>" & vbCrLf _
''                    & "<link rel=" & Chr(34) & "stylesheet" & Chr(34) & " href=" & Chr(34) & "../../style/MACRO1.css" & Chr(34) & " type=" & Chr(34) & "text/css" & Chr(34) & ">" & vbCrLf _
''                    & "<title>InferMed MACRO v<%=application(" & Chr(34) & "APPVERSION" & Chr(34) & ")%> - Data Entry - " _
''                    & rsCRFPages!CRFTitle & "</title>" & vbCrLf _
''                    & sDefaultFontAttr & vbCrLf _
''                    & "<script language=" & Chr(34) & "javascript" & Chr(34) & ">" & vbCrLf _
''                    & "var frm1;" & vbCrLf _
''                    & "var frm2;" & vbCrLf _
''                    & "function saveEform(nxt){document.deFrm.h" & sFile & ".value=nxt;frm1.fnFinalisePage(document.deFrm);}" & vbCrLf _
''                    & "function hideLoader(){document.all.loadingDiv.style.visibility='hidden';}" & vbCrLf _
''                    & "</script>" & vbCrLf
'
'
'        ' rem 19/10/01 removed stylesheet link as it was setting background colour of eform
'        Print #nFile, "<html>" & vbCrLf _
'                    & "<head>" & vbCrLf _
'                    & sDefaultFontAttr & vbCrLf _
'                    & "<script language=" & Chr(34) & "javascript" & Chr(34) & ">" & vbCrLf _
'                    & "var frm1;" & vbCrLf _
'                    & "var frm2;" & vbCrLf _
'                    & "var frm3;" & vbCrLf _
'                    & "function pageLoaded(){" & vbCrLf _
'                    & "window.parent.frames[0].fnInitialiseApplet();" & vbCrLf _
'                    & "JSVEInitFunc();" & vbCrLf _
'                    & "<%If (nStatus = 1) Then Response.Write(" & Chr(34) & "alert('MACRO encountered a problem while saving.\nOne or more responses could not be saved');" & Chr(34) & ")%>" & vbCrLf _
'                    & "hideLoader();}" & vbCrLf _
'                    & "function saveEform(nxt){showSaver();document.deFrm[" & Chr(34) & "f_" & sFile & Chr(34) & "].value=nxt;frm1.fnFinalisePage(frm2);}" & vbCrLf _
'                    & "function hideLoader(){window.parent.frames[0].dispMnu(true);document.all.loadingDiv.style.visibility='hidden';}" & vbCrLf _
'                    & "function showSaver(){window.parent.frames[0].dispMnu(false);document.all.savingDiv.style.visibility='visible';}" & vbCrLf _
'                    & "function logOut(){if(confirm('About to close current MACRO session?'))window.parent.navigate(" & Chr(34) & "../general/Logout.asp" & Chr(34) & ")}" & vbCrLf _
'                    & "</script>" & vbCrLf
'
'
'
'        Print #nFile, "<link rel=" & Chr(34) & "stylesheet" & Chr(34) & " href=" & Chr(34) & "../../style/messagebox.css" & Chr(34) & " type=" & Chr(34) & "text/css" & Chr(34) & ">" & vbCrLf _
'                    & "<style type=" & Chr(34) & "text/css" & Chr(34) & ">" _
'                    & "#loadingDiv{position:absolute; visibility:visible; left:50; top:50; width:300px; height:100px; background-color:#d3d3d3 z-index:100;}" _
'                    & "#savingDiv{position:absolute; visibility:hidden; left:50; top:50; width:300px; height:100px; background-color:#d3d3d3 z-index:100;}" _
'                    & "</style>" & vbCrLf _
'                    & "</head>" & vbCrLf
'
''    '   ATN 7/12/99
''    '   Variable definitions will now be set up in the web rde dll.
''    '   added CRFElement.CRFElementId.
''        'JL 8/09/00. Added UnitOfMeasurement in the select for passing to new function JavascriptDisableCaptions.
''        sSQL = "SELECT DataItemCode, UnitOfMeasurement" _
''                & " FROM DataItem, CRFElement where " _
''                & " CRFPageId =" & lCRFPageId _
''                & " AND DataItem.DataItemId = CRFElement.DataItemId " _
''                & " AND DataItem.ClinicalTrialId = CRFElement.ClinicalTrialId " _
''                & " AND DataItem.VersionId = CRFElement.VersionId " _
''                & " AND DataItem.ClinicalTrialId =" & vClinicalTrialId _
''                & " AND DataItem.VersionId=" & vVersionId
''
''        Set rsDataItem = New ADODB.Recordset
''        rsDataItem.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
''
'''       'JL 26/10/00
''        If Not rsDataItem.EOF Then
''            If IsNull(rsDataItem!UnitOfMeasurement) Then
''                sUnitOfMeasurement = RemoveNull(rsDataItem!UnitOfMeasurement)
''            End If
''        End If
'
'
'        If (rsCRFPages!BackgroundColour = 0) Then
'            sBgCol = "#" & GetFontColour(rsStudyDefinition!DefaultCRFPageColour)
'        Else
'            sBgCol = "#" & GetFontColour(rsCRFPages!BackgroundColour)
'        End If
'        Print #nFile, "<body bgcolor=" & Chr(34) & sBgCol & Chr(34) & " onload=" & Chr(34) & "javascript:pageLoaded();" & Chr(34) & ">"
'
'        'rem 22/10/01 added "style=z-index:100" to <div> tag of each message box to ensure it appeared infront of text boxes on eform
'        Print #nFile, "<div class=" & Chr(34) & "PROCESS_MESSAGE_BOX" & Chr(34) & " id=" & Chr(34) & "loadingDiv" & Chr(34) & " style=" & Chr(34) & "z-index:100" & Chr(34) & ">" & vbCrLf _
'                    & "<table bgcolor=" & Chr(34) & "#d3d3d3" & Chr(34) & " height=" & Chr(34) & "100%" & Chr(34) & " align=" & Chr(34) & "center" & Chr(34) & " width=" & Chr(34) & "90%" & Chr(34) & " class=" & Chr(34) & "MESSAGE" & Chr(34) & vbCrLf _
'                    & "<tr><td valign=" & Chr(34) & "middle" & Chr(34) & "><b>please wait</b><br><br><img src=" & Chr(34) & "../../img/clock.gif" & Chr(34) & ">" & vbCrLf _
'                    & "&nbsp;&nbsp;loading " & rsCRFPages!CRFTitle & "..</td></tr></table></div>" & vbCrLf _
'                    & "<div class=" & Chr(34) & "PROCESS_MESSAGE_BOX" & Chr(34) & " id=" & Chr(34) & "savingDiv" & Chr(34) & " style=" & Chr(34) & "z-index:100" & Chr(34) & ">" & vbCrLf _
'                    & "<table bgcolor=" & Chr(34) & "#d3d3d3" & Chr(34) & " height=" & Chr(34) & "100%" & Chr(34) & " align=" & Chr(34) & "center" & Chr(34) & " width=" & Chr(34) & "90%" & Chr(34) & " class=" & Chr(34) & "MESSAGE" & Chr(34) & vbCrLf _
'                    & "<tr><td valign=" & Chr(34) & "middle" & Chr(34) & "><b>please wait</b><br><br><img src=" & Chr(34) & "../../img/clock.gif" & Chr(34) & ">" & vbCrLf _
'                    & "&nbsp;&nbsp;saving " & rsCRFPages!CRFTitle & "..</td></tr></table></div>" & vbCrLf _
'                    & "<%Response.Flush" & vbCrLf _
'                    & "if session(" & Chr(34) & "usrStudyDefOK" & Chr(34) & ") = " & Chr(34) & "True" & Chr(34) & " Then" & vbCrLf _
'                    & "Response.Write(oIo.getJSVEInitFunc(usrId,fltDb,fltSi,fltSt,fltSj,fltId,session(" & Chr(34) & "usrStudyDef" & Chr(34) & ")))" & vbCrLf _
'                    & "Else" & vbCrLf _
'                    & "Response.Write(oIo.getJSVEInitFunc(usrId,fltDb,fltSi,fltSt,fltSj,fltId))" & vbCrLf _
'                    & "End If" & vbCrLf _
'                    & "%>" & vbCrLf
'
'
''        Print #nFile, "<script language=" & Chr(34) & "javaScript" & Chr(34) & " src=" & Chr(34) & "../../script/menu2.js" & Chr(34) & "></script>" & vbCrLf _
''                    & "<script language=" & Chr(34) & "javaScript" & Chr(34) & ">" & vbCrLf _
''                    & "function logOut(){if(confirm('About to close current MACRO session?'))window.navigate(" & Chr(34) & "../general/Logout.asp" & Chr(34) & ")}" & vbCrLf _
''                    & "</script>" & vbCrLf _
''                    & "<%=getMnu(7)%>" _
''                    & "<script language=" & Chr(34) & "JavaScript" & Chr(34) & ">" & vbCrLf _
''                    & "function UpdateIt(){document.all['mnuBarDiv'].style.top=document.body.scrollTop;setTimeout('UpdateIt()',200);}" & vbCrLf _
''                    & "UpdateIt();" & vbCrLf _
''                    & "</script>"
'
'
'        Print #nFile, "<form name=" & Chr(34) & "deFrm" & Chr(34); " method=" & Chr(34) & "post" & Chr(34) & " action=" & Chr(34) & sFile & sExt & "?fltSi=<%=fltSi%>&fltSj=<%=fltSj%>&fltId=<%=fltId%>" & Chr(34) & ">" & vbCrLf _
'                    & "<input type=" & Chr(34) & "hidden" & Chr(34) & " name=" & Chr(34) & "f_" & sFile & Chr(34) & ">" & vbCrLf
'
'
'        If rsCRFElement.EOF Then
'            bNoOfelements = False
'        End If
'
''        'JL 07/09/00. Build Javascript function to dim captions grey if disabled.
''        nNumberOfSkips = 0
''        nNumberOfRadios = 0
''        Do While Not rsCRFElement.EOF
''            If Not IsNull(rsCRFElement!SkipCondition) And rsCRFElement!SkipCondition > "" Then
''                nNumberOfSkips = nNumberOfSkips + 1
''            End If
''            If rsCRFElement!ControlType = gsOPTION_BUTTONS Or rsCRFElement!ControlType = gsPUSH_BUTTONS Then
''                nNumberOfRadios = nNumberOfRadios + 1
''            End If
''
''            lDataItemId = rsCRFElement!DataItemId
''
''            rsCRFElement.MoveNext
''        Loop
''
''        If nNumberOfRadios > 0 Then
''            Print #nFile, JavascriptUpdateRadioValues(vClinicalTrialId, vVersionId, lCRFPageId)
''        End If
''
''        If nNumberOfSkips > 0 Then
''                Print #nFile, JavascriptDisableCaptions(rsStudyDefinition, vClinicalTrialId, vVersionId, lCRFPageId, bDispNum, _
''                                    sUnitOfMeasurement, bNetscape)
''        Else
''                'JL 12/02/01 Bug. still need a function as getCaption is called in JS functions (e.g. DisplayForm())
''                Print #nFile, "<script language =""javascript"">" & vbNewLine & _
''                                            vbTab & "function getCaption(){ return """" }" & vbNewLine & "</script>"
''        End If
''        'JL 26/10/00
''        If Not bNoOfelements = False Then
''            rsCRFElement.MoveFirst
''        End If
''        'CRFELEMENT
'
'
'
'        Do While Not rsCRFElement.EOF
'        '-------------------------------------------------------------------------------------------------
'        'JL 16/08/00. Added checking of question RoleCode
'        '-------------------------------------------------------------------------------------------------
'
'        sQuestionRoleCode = RemoveNull(rsCRFElement!RoleCode)
'
'        'JL 16/11/00. Remove uneccesary server side code.
''        Print #nFile, "<% sQuestionStatus = GetStatusText(getCurrentStatus(lthisId)) %>"
'
''        lImageXposition = 0
'
'            lDataItemId = rsCRFElement!DataItemId
'
'                sSQL = "SELECT * FROM DataItem WHERE " _
'                    & " ClinicalTrialId =" & vClinicalTrialId _
'                    & " AND DataItemId =" & lDataItemId _
'                    & " AND VersionID =" & vVersionId
'
'                Set rsDataItem = New ADODB.Recordset
'                rsDataItem.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
''                If lDataItemId > 0 Then
''                Debug.Print rsDataItem!DataItemCode
''                    'Remove any Nulls right at the start.
''                    sDerivation = RemoveNull(rsDataItem!Derivation)
''
''                    If sDerivation > "" Then
''                        sDisabledSetting = sDISABLED
''                    Else
''                        sDisabledSetting = sENABLED
''                    End If
''
''                End If
'
'
'        'lThisId used as argument for routines in dataValues.asp
'        'Dont print for visual elements
''        If rsCRFElement!ControlType <> gsCOMMENT And rsCRFElement!ControlType <> gsPICTURE And _
''                    rsCRFElement!ControlType <> gsLINE Then
'
''            Print #nFile, "<% lthisId= " & rsCRFElement!CRFElementId & "%>"
''        End If
'
'        '---------------------------------------------------------------------------'
'        '                               *** COMMENTS LABEL  ***                     '
'        '---------------------------------------------------------------------------'
'        If rsCRFElement!ControlType = ControlType.Comment Then
'            sComments = ""
'                    Print #nFile, "<div Style=" & Chr(34) & "top:" & CLng(rsCRFElement!Y / nIeYSCALE) & ";left:" & CLng(rsCRFElement!X / nIeXSCALE) & ";z-index:1" & Chr(34) & ">"
'
''                    Print #nFile, "<p>"
'
'                    ' DPH 29/10/2001 - Use special bool because no question number with a comment
'                    'bDispNum = False
'                    bCommentDispNum = False
'
'                    sComments = sComments & "<font"
'                    sComments = sComments & printFontAttributes(rsCRFElement!FontColour, rsCRFElement!FontSize, rsCRFElement!FontName _
'                            , rsStudyDefinition!DefaultFontColour, rsStudyDefinition!DefaultFontSize, rsStudyDefinition!DefaultFontName, bNetscape) & _
'                    FieldAttributes_Explorer(rsCRFElement!FontBold, rsCRFElement!FontItalic, rsCRFElement!FieldOrder, rsCRFElement!Caption _
'                                , bCommentDispNum)
'                    sComments = sComments & "</font>"
'
'                    Print #nFile, sComments
'
' '                   Print #nFile, "</p>"
'                    Print #nFile, "</div>" & vbCrLf
'
'        '---------------------------------------------------------------------------'
'        '                                  *** Hidden  ***                          '
'        '---------------------------------------------------------------------------'
'        ElseIf rsCRFElement!Hidden = 1 Then
'                    Print #nFile, "<input type=" & Chr(34) & "hidden" & Chr(34) & " " _
'                                       & "name=" & Chr(34) & "f_" & LCase(rsDataItem!DataItemCode) & Chr(34) & ">"
''                    Print #nFile, CreateFormObject(rsDataItem, rsCRFElement, rsStudyDefinition)
'
'        '---------------------------------------------------------------------------'
'        '                                  *** LINE  ***                            '
'        '---------------------------------------------------------------------------'
'        ElseIf rsCRFElement!ControlType = ControlType.Line Then
'                Print #nFile, "<div Style=" & Chr(34) & "20px;left:0;top:" & CLng(rsCRFElement!Y / nIeYSCALE) & Chr(34) & ">"
'                Print #nFile, "<hr>"
'
'                Print #nFile, "</div>" & vbCrLf
'        '---------------------------------------------------------------------------'
'        '                           *** TEXTBOX or DATE  ***                        '
'        '---------------------------------------------------------------------------'
'        ElseIf rsCRFElement!ControlType = ControlType.TextBox Or rsCRFElement!ControlType = ControlType.Calendar _
'            Or rsCRFElement!ControlType = ControlType.RichTextBox Then
'        'jl RichTextBoxes are not supported in the web, but display anyway for now.
'                            'Jl 30/05/00 Code transfered to new printCaption.
'                            '------------------------------------'
'                            '        *** PRINT CAPTION  ***      '
'                            '------------------------------------'
'                            '07/08/00
'            Print #nFile, printCaption(rsCRFElement, rsStudyDefinition, rsDataItem, _
'                     bDispNum, bNetscape, nDefaultFontSize, lDefaultFontColour, _
'                     sDefaultFontFamily, nDefaultFontWeight, nDefaultFontStyle)
'
'                    If IsNull(rsDataItem!DataItemFormat) Then
'                        SetElementWidth "", rsDataItem!DataItemLength, nMaxLength, nTextBoxCharWidth
'                    Else
'                        SetElementWidth rsDataItem!DataItemFormat, rsDataItem!DataItemLength, nMaxLength, nTextBoxCharWidth
'                    End If
'
'                '2##
'                Print #nFile, "<div id=" & Chr(34) & "f_" & LCase(rsDataItem!DataItemCode) & "_inpDiv" & Chr(34) & " style=" & Chr$(34) & _
'                                    ";left:" & CLng(rsCRFElement!X / nIeXSCALE) & ";top:" & CLng(rsCRFElement!Y / nIeYSCALE) & ";z-index:2" & Chr$(34) & ">"
'
'                'JL 07/08/00 Replacing use of R recordset.
'                'If a disabled control no need for event handlers.
'
'                Print #nFile, "<input type=" & Chr(34) & "text" & Chr(34) & " " _
'                                   & "disabled " _
'                                   & "tabindex=" & Chr(34) & rsCRFElement!FieldOrder & Chr(34) & " " _
'                                   & "name=" & Chr(34) & "f_" & LCase(rsDataItem!DataItemCode) & Chr(34) & " " _
'                                   & "size=" & Chr(34) & nTextBoxCharWidth & Chr(34) & " " _
'                                   & "maxlength=" & Chr(34) & nMaxLength & Chr(34) & " " _
'                                   & "onblur=" & Chr(34) & "frm1.fnLostFocus(" & "'f_" & LCase(rsDataItem!DataItemCode) & "','#FFFFFF');" & Chr(34) & " " _
'                                   & "onfocus=" & Chr(34) & "frm1.fnGotFocus(" & "'f_" & LCase(rsDataItem!DataItemCode) & "','#FFFF80');" & Chr(34) & ">"
'
'
'    '   ATN 3/12/99
'    '   Javascript validation now created at form level at run-time
'    '           'Populate the Form Objects property value
'                'Print #nFile, GetDataValue(rsDataItem!DataItemCode)
''                Print #nFile, CreateFormObject(rsDataItem, rsCRFElement, rsStudyDefinition)
''                Print #nFile, "</div>"
'
''                If nMaxLength > nTextBoxCharWidth Then
''                    nTextBoxCharWidth = nMaxLength
''                End If
'
'                'jl 16/06/00 Replaced lControlLeft with CLng(rsCRFElement!CaptionX / nIeXSCALE)
''                lImageXposition = CalulateStatusXPosition(nTextBoxCharWidth, CLng(rsCRFElement!X / nIeXSCALE), gsTEXT_BOX)
'
'
'                'JL 13/07/00.
'                'Always call SetStatusImage, if status=Requested then the image is set to a dummy transparent image.
''                Print #nFile, "<div id=" & Chr(34) & "fld" & rsCRFElement!CRFElementId & "_imgDiv" & Chr(34) & " style=""Left:" & lImageXposition & ";top:" & CLng(rsCRFElement!Y / nIeYSCALE) & """>"
'
'                Print #nFile, "<img src=" & Chr(34) & "../../img/blank.gif" & Chr(34) & " " _
'                                 & "name=" & Chr(34) & "img_" & LCase(rsDataItem!DataItemCode) & Chr(34) & ">"
'
'                Print #nFile, "</div>" & vbCrLf
'            '---------------------------------------------------------------------------'
'            '                           *** MULTIMEDIA CONTROL  ***                     '
'            '---------------------------------------------------------------------------'
'        ElseIf rsCRFElement!ControlType = ControlType.Attachment Then
'
'                                'Jl 30/05/00 Code transfered to new printCaption.
'                                '------------------------------------'
'                                '        *** PRINT CAPTION  ***      '
'                                '------------------------------------'
'                '07/08/00
'                Print #nFile, printCaption(rsCRFElement, rsStudyDefinition, rsDataItem, _
'                         bDispNum, bNetscape, nDefaultFontSize, lDefaultFontColour, _
'                         sDefaultFontFamily, nDefaultFontWeight, nDefaultFontStyle)
'
'                'JL 23/05/00
'                Print #nFile, "<div id=" & Chr(34) & "f_" & LCase(rsDataItem!DataItemCode) & "_inpDiv" & Chr(34) & " style=" & Chr$(34) & _
'                    "left:" & CLng(rsCRFElement!X / nIeXSCALE) & ";top:" & CLng(rsCRFElement!Y / nIeYSCALE) & Chr$(34) & ">"
'
'                Print #nFile, "<input type=" & Chr(34) & "file" & Chr(34) & " " _
'                                   & "disabled " _
'                                   & "tabindex=" & Chr(34) & rsCRFElement!FieldOrder & Chr(34) & " " _
'                                   & "name=" & Chr(34) & "f_" & LCase(rsDataItem!DataItemCode) & Chr(34) & " " _
'                                   & "onblur=" & Chr(34) & "frm1.fnLostFocus(" & "'f_" & LCase(rsDataItem!DataItemCode) & "','#FFFFFF');" & Chr(34) & " " _
'                                   & "onfocus=" & Chr(34) & "frm1.fnGotFocus(" & "'f_" & LCase(rsDataItem!DataItemCode) & "','#FFFF80');" & Chr(34) & ">"
'
'
''                Print #nFile, "</div>"
'                'If a Multimedia Control allow for more space on the left.
'
''               lImageXposition = CalulateStatusXPosition(nTextBoxCharWidth, CLng(rsCRFElement!X / nIeXSCALE), gsATTACHMENT)
'
'                'JL 07/04/00 Print out Question Status Image.
'                'Prints out status image if there is a record in the DataItemResponse table.
'
'                'JL 13/07/00. Always call SetStatusImage.
''                 Print #nFile, "<div id=" & Chr(34) & "fld" & rsCRFElement!CRFElementId & "_imgDiv" & Chr(34) & " style=""Left:" & lImageXposition & ";top:" & CLng(rsCRFElement!Y / nIeYSCALE) & """> "
'
'                 Print #nFile, "&nbsp;<img src=" & Chr(34) & "../../img/blank.gif" & Chr(34) & " " _
'                                 & "name=" & Chr(34) & "img_" & LCase(rsDataItem!DataItemCode) & Chr(34) & ">"
'
'                 Print #nFile, "</div>" & vbCrLf
'
'
' '               Print #nFile, CreateFormObject(rsDataItem, rsCRFElement, rsStudyDefinition)
'
'            '---------------------------------------------------------------------------'
'            '                         *** RADIOS and OPTION BOXES  ***                  '
'            '---------------------------------------------------------------------------'
'            ElseIf rsCRFElement!ControlType = ControlType.OptionButtons Or rsCRFElement!ControlType = ControlType.PushButtons Then 'check gsPUSH_BUTTONS
'               'Get values for option boxs etc
'
'                ' DPH 29/10/2001 - Get only active codes
'                ' DPH 5/11/2001 - Order Categories to ValueOrder
'                sSQL = "SELECT * FROM ValueData WHERE" & _
'                    " DataItemId = " & lDataItemId & " AND" & _
'                    " ClinicalTrialId = " & vClinicalTrialId & "AND " & _
'                    " VersionId = " & vVersionId & " AND Active = 1" & _
'                    " ORDER BY ValueOrder "
'
'
'                Set rsValueData = New ADODB.Recordset
'                rsValueData.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
'    '   ATn 26/4/99
'    '   Added check for at least one value being defined for a category data item
'    'lImageXposition
'                If rsValueData.RecordCount > 0 Then
'
'                'Jl 30/05/00 Code transfered to new printCaption.
'
'                    Print #nFile, "<div id=" & Chr(34) & "f_" & LCase(rsDataItem!DataItemCode) & "_capDiv" & Chr(34) & " style=" & Chr(34) & "top:0;left:0;" & Chr(34) & ">"
'
'                    Print #nFile, "<div id=" & Chr(34) & "f_" & LCase(rsDataItem!DataItemCode) & "_inpDiv" & Chr(34) & " style=" & Chr$(34) & ";left:" & _
'                                       CLng(rsCRFElement!X / nIeXSCALE) & ";top:" & CLng(rsCRFElement!Y / nIeYSCALE) & ";z-index:2" & Chr$(34) & ">"
'
'                     rsValueData.MoveLast
'
' '                   Print #nFile, "<% redim v(" & rsValueData.RecordCount & ")"
' '                   Print #nFile, "redim c(" & rsValueData.RecordCount & ")"
'
'                    Print #nFile, "<table><tr><td onclick=" & Chr(34) & "frm1.fnLostFocus(" & "'f_" & LCase(rsDataItem!DataItemCode) & "','#FFFF80');" & Chr(34) & ">"
'
'                    nRCount = 1
'
'                    nLengthOfItem = 0
'                    nMaxLengthSoFar = 0
'                    rsValueData.MoveFirst
'                    Do While Not rsValueData.EOF 'For Each Radio Item
'                        Print #nFile, "<input type=" & Chr(34) & "radio" & Chr(34) & " " _
'                                           & "disabled " _
'                                           & "tabindex=" & Chr(34) & rsCRFElement!FieldOrder & Chr(34) & " " _
'                                           & "name=" & Chr(34) & "f_" & LCase(rsDataItem!DataItemCode) & Chr(34) & " " _
'                                           & "value=" & Chr(34) & rsValueData!ValueCode & Chr(34) & " " _
'                                           & "onblur=" & Chr(34) & "frm1.setFieldBGColour(" & "'f_" & LCase(rsDataItem!DataItemCode) & "','" & sBgCol & "');" & Chr(34) & " " _
'                                           & "onfocus=" & Chr(34) & "frm1.fnGotFocus(" & "'f_" & LCase(rsDataItem!DataItemCode) & "','#FFFF80');" & Chr(34) & " " _
'                                           & "onchange=" & Chr(34) & "frm1.fnLostFocus(" & "'f_" & LCase(rsDataItem!DataItemCode) & "','" & sBgCol & "');" & Chr(34) & ">" _
'                                           & "<font color=" & Chr(34) & GetFontColour(lDefaultFontColour) & Chr(34) & " face=" & Chr(34) & sDefaultFontFamily & Chr(34) & " size=" & Chr(34) & getFontSize(nDefaultFontSize) & Chr(34) & ">" & rsValueData!ItemValue & "</font><br>"
'
'
'
''                        'JL 06/02/00 Replaced rsValueData.AbsolutePosition with nRCount for recordcount, see notes.
''                        Print #nFile, "c(" & nRCount & ") =""" & rsValueData!ValueCode & """"
''                        Print #nFile, "v(" & nRCount & ") =""" & rsValueData!ItemValue & """"
''                        nRCount = nRCount + 1
''                        'JL 16/06/00
''                        nLengthOfItem = Len(rsValueData!ItemValue)
''                        If nLengthOfItem > nMaxLengthSoFar Then
''                            nMaxLengthSoFar = nLengthOfItem
''                        End If
'                        rsValueData.MoveNext
'                    Loop
'
'                    'if only one radio button, add a hidden field of the same name to create a control array
'                    If rsValueData.RecordCount = 1 Then
'                        rsValueData.MoveFirst
'                        Print #nFile, "<input type=" & Chr(34) & "hidden" & Chr(34) & " disabled name=" & Chr(34) & "f_" & LCase(rsDataItem!DataItemCode) & Chr(34) & ">"
'                    End If
'
'                    Print #nFile, "</td><td valign=top>"
'
'                    'JL 16/08/00 Added extra argument to Radio for disabling control.
'                    'JL 16/08/00 Added Check for authoisation code.
'
''                        Print #nFile, " Radio " & """opt" & LCase$(rsDataItem!DataItemCode) & """, """ & LCase$(rsDataItem!DataItemCode) & """, " & rsCRFElement!FieldOrder & ",getResponseValue(" & rsCRFElement!CRFElementId & ") %>" & vbNewLine
'
'
'                '   ATN 3/12/99
'                '   Javascript validation now created at form level at run-time
'                '   Print #nFile, JavascriptValidation(rsDataItem!DataItemCode, rsDataItem!DataType, rsDataItem!MandatoryValidation, rsDataItem!WarningValidation)
'                    'Print #nFile, GetDataValue(rsDataItem!DataItemCode)
''                    Print #nFile, CreateFormObject(rsDataItem, rsCRFElement, rsStudyDefinition)
'
''                    Print #nFile, "</div>"
'                    'JL 07/04/00 Print out Status Image.
'                    'JL 16/06/00 Replaced lImageXposition with call to calulateStatusXPosition
'                    'For th control width if Radio need to calulate the max with of the ItemValue (caption for each item) and add it
'                    'to the control width.
''                    lImageXposition = CalulateStatusXPosition(70, CLng(rsCRFElement!X / nIeXSCALE), gsOPTION_BUTTONS, nMaxLengthSoFar)
'
''                    Print #nFile, "<div id=" & Chr(34) & "fld" & rsCRFElement!CRFElementId & "_imgDiv" & Chr(34) & " style=""Left:" & lImageXposition & ";top:" & CLng(rsCRFElement!Y / nIeYSCALE) & """> "
'
'                    Print #nFile, "&nbsp;<img src=" & Chr(34) & "../../img/blank.gif" & Chr(34) & " " _
'                                 & "name=" & Chr(34) & "img_" & LCase(rsDataItem!DataItemCode) & Chr(34) & ">"
'
'                    Print #nFile, "</td></tr></table>"
'
'                    Print #nFile, "</div>"
'                     '07/08/00
'                        Print #nFile, printCaption(rsCRFElement, rsStudyDefinition, rsDataItem, _
'                        bDispNum, bNetscape, nDefaultFontSize, lDefaultFontColour, _
'                        sDefaultFontFamily, nDefaultFontWeight, nDefaultFontStyle)
'
'                    Print #nFile, "</div>" & vbCrLf
'
'                End If
'            '---------------------------------------------------------------------------'
'            '                               *** PICTURE  ***                            '
'            '---------------------------------------------------------------------------'
'            ElseIf rsCRFElement!ControlType = ControlType.Picture Then
'    '   ATN 4/5/99
'    '   Use virtual directory 'WebRDEDocuments
'    '   JL 16/06/00. Have to get this set up in the meantime change to Documents which actually exists.
'
'                'ic 01/12/00
'                'added condition so that if the element  cannot be found it is omitted
'                If Dir(gsDOCUMENTS_PATH & rsCRFElement!Caption) <> "" Then
'                    Set oPicture = LoadPicture(gsDOCUMENTS_PATH & rsCRFElement!Caption)
'                    ElementWidth = frmGenerateHTML.ScaleY(oPicture.Width, vbHimetric, vbPixels) * nXSCALE
'                    Set oPicture = Nothing
'
'                    If ElementWidth > 7000 And ElementWidth < 5000 Then 'Assume it is a line
'                        Dim maxImageLength As Integer               'limit it to the width of the page.
'                        nMaxImageLength = 611
'                        Print #nFile, "<div id=" & Chr(34) & rsCRFElement!CRFElementId & "_inpDiv" & Chr(34) & " style=" & Chr$(34) & _
'                                           ";left:" & CLng(rsCRFElement!X / nIeXSCALE) & ";top:" & CLng(rsCRFElement!Y / nIeYSCALE) & Chr$(34) & ">"
'
'                        'ic 12/02/01
'                        'removed gsDOCUMENTS_PATH constant as it was being written in the html
'                        Print #nFile, "<img src=" & Chr$(34) & rsCRFElement!Caption & Chr$(34) & " width=" & nMaxImageLength & "  align=""top"" align=""left"">"
'                        Print #nFile, "</div>"
'
'                    Else
'                        Print #nFile, "<div id=" & Chr(34) & rsCRFElement!CRFElementId & Chr(34) & "_inpDiv" & " style=" & Chr$(34) & _
'                                           ";left:" & CLng(rsCRFElement!X / nIeXSCALE) & ";top:" & CLng(rsCRFElement!Y / nIeYSCALE) & Chr$(34) & ">"
'
'                        'ic 12/02/01
'                        'removed gsDOCUMENTS_PATH constant as it was being written in the html
'                        Print #nFile, "<img src=" & Chr$(34) & rsCRFElement!Caption & Chr$(34) & " align=""top"" align=""left"">"
'                        Print #nFile, "</div>"
'                    End If
'
'
'
'
'                    'ic 05/12/00
'                    'copy the file to the web folder
'                    'because writing the asp is currently done for both IE and Netscape, but the
'                    'pictures will be the same for both, im only going to copy them in the IE asp function. if the 2 functions are
'                    'ever combined this code will need to be included.
'                    'it would be preferrable to add more flexible error handling here to check if the file we're about to overwrite
'                    'is currently open (eg in an image editor) but for now we'll assume we can just go ahead
'
'                    'check we havent already copied this file by looking for the filename in our copied images variable
'                    'which is just a comma delimited list of files we add to as we process this function
'                    If InStr(sCpdImages, "," & rsCRFElement!Caption) = 0 Then
'
'                        'if the file exists in the html folder already, delete it
'                        If Dir(gsWEB_HTML_LOCATION & rsCRFElement!Caption) <> "" Then Kill (gsWEB_HTML_LOCATION & rsCRFElement!Caption)
'
'                        'copy the file from source to destination
'                        FileCopy gsDOCUMENTS_PATH & rsCRFElement!Caption, gsWEB_HTML_LOCATION & rsCRFElement!Caption
'
'                        'add to our list of copied images
'                        sCpdImages = sCpdImages & "," & rsCRFElement!Caption
'                    End If
'
'                Else
''                   MsgBox "Image " & gsDOCUMENTS_PATH & rsCRFElement!Caption & " cannot be found!", vbOKOnly
'                End If
'            '---------------------------------------------------------------------------'
'            '                               *** POP-UP LIST  ***                        '
'            '---------------------------------------------------------------------------'
'            ElseIf rsCRFElement!ControlType = ControlType.PopUp Then
'                   'Get values for option boxs etc
'               ' vDataItemId2 = rsCRFElement!DataItemId
'                ' DPH 29/10/2001 - Get only active codes
'                ' DPH 5/11/2001 - Order Categories to ValueOrder
'                sSQL = "SELECT * FROM ValueData WHERE" & _
'                    " DataItemId = " & lDataItemId & " AND" & _
'                    " ClinicalTrialId = " & vClinicalTrialId & "AND " & _
'                    " VersionId = " & vVersionId & " AND Active = 1" & _
'                    " ORDER BY ValueOrder "
'
'                Set rsValueData = New ADODB.Recordset
'                rsValueData.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
'
'                'Jl 30/05/00 Code to print caption transfered to new function printCaption.
'                Print #nFile, "<div id=" & Chr(34) & "f_" & LCase(rsDataItem!DataItemCode) & "_inpDiv" & Chr(34) & " style=" & Chr$(34) & _
'                    sPositionAttr & ";left:" & CLng(rsCRFElement!X / nIeXSCALE) & ";top:" & CLng(rsCRFElement!Y / nIeYSCALE) & ";z-index:2" & Chr$(34) & ">"
'
'
'                'Print #nFile, "<select Disabled tabindex=" & Chr$(34) & rsCRFElement!FieldOrder & Chr$(34) & " Name =""cbo" _
'                '& LCase$(rsDataItem!DataItemCode) & """" & " onfocus=""setFocusColour(f.cbo" & LCase(rsDataItem!DataItemCode) & ");""" & _
'                '" onClick=""RefreshForm(f.cbo" & LCase$(rsDataItem!DataItemCode) & ")"" " & ">"
'
'                'ic 21/12/00
'                'added onblur event so that answers are evaluated when tabbing between fields as well as clicking
'                Print #nFile, "<select tabindex=" & Chr(34) & rsCRFElement!FieldOrder & Chr(34) & " " _
'                                    & "disabled " _
'                                    & "name =" & Chr(34) & "f_" & LCase(rsDataItem!DataItemCode) & Chr(34) & " " _
'                                    & "onfocus=" & Chr(34) & "frm1.fnGotFocus(" & "'f_" & LCase(rsDataItem!DataItemCode) & "','#FFFF80');" & Chr(34) & " " _
'                                    & "onchange=" & Chr(34) & "frm1.fnLostFocus(" & "'f_" & LCase(rsDataItem!DataItemCode) & "','#FFFF80');" & Chr(34) & " " _
'                                    & "onblur=" & Chr(34) & "frm1.setFieldBGColour(" & "'f_" & LCase(rsDataItem!DataItemCode) & "','" & sBgCol & "');" & Chr(34) & ">"
'
'                 'REM 12/10/01 To put space as default in pop-up box
'                 Print #nFile, "<option value=" & Chr(34) & Chr(34) & "> </option>"
'
'                 Do While Not rsValueData.EOF
'                    Print #nFile, "<option value=" & Chr(34) & rsValueData!ValueCode & Chr(34); ">" _
'                                & rsValueData!ItemValue _
'                                & "</option>"
'                    rsValueData.MoveNext
'
'                 Loop
'
'
''                'Test for at least 1 datavalue
''                If rsValueData.RecordCount > 0 Then
''                    rsValueData.MoveLast
''
''                    Print #nFile, "<% redim v(" & rsValueData.RecordCount & ")"
''                    Print #nFile, "redim c(" & rsValueData.RecordCount & ")"
''
''                    nLengthOfItem = 0
''                    nMaxLengthSoFar = 0
''                    nRCount = 1
''                    Dim nUpper As Integer
''                    Dim nLower As Integer
''                    Dim nCategoryValue As String
''                    Dim bMostlyUpperCase As Boolean
''                    Dim i As Integer
''                    rsValueData.MoveFirst
''                    Do While Not rsValueData.EOF
''                        '06/02/00 JL nRCount replaces (rsValueData.AbsolutePosition + 1) AbsolutePosition not working.
''                        Print #nFile, "c(" & nRCount & ") =""" & rsValueData!ValueCode & """"
''                        Print #nFile, "v(" & nRCount & ") =""" & rsValueData!ItemValue & """"
''                        nRCount = nRCount + 1
''                        nLengthOfItem = Len(rsValueData!ItemValue)
''                        If nLengthOfItem > nMaxLengthSoFar Then
''                            nMaxLengthSoFar = nLengthOfItem
''                            nCategoryValue = rsValueData!ItemValue
''                        End If
''                        rsValueData.MoveNext
''                    Loop
''                End If 'Test for at least 1 datavalue
''
''                bMostlyUpperCase = mostlyUpperCase(nCategoryValue)
''
''                Print #nFile, "SelValues getResponseValue(lthisId) %>" & vbNewLine
'
'                Print #nFile, "</select>"
'
'    '   ATN 3/12/99
'    '   Javascript validation now created at form level at run-time
'
' '               Print #nFile, CreateFormObject(rsDataItem, rsCRFElement, rsStudyDefinition)
'
'                Print #nFile, "&nbsp;<img src=" & Chr(34) & "../../img/blank.gif" & Chr(34) & " " _
'                                 & "name=" & Chr(34) & "img_" & LCase(rsDataItem!DataItemCode) & Chr(34) & ">"
'
'                Print #nFile, "</div>"
'                'JL 16/06/00 Print out Status Image.
'
''                lImageXposition = CalulateStatusXPosition(nMINIMAL_POPUP_WIDTH, CLng(rsCRFElement!X / nIeXSCALE), gsPOPUP_LIST, nMaxLengthSoFar, bMostlyUpperCase)
'
''                Print #nFile, "<div id=" & Chr(34) & "fld" & rsCRFElement!CRFElementId & "_imgDiv" & Chr(34) & " style=""Left:" & lImageXposition & ";top:" & CLng(rsCRFElement!Y / nIeYSCALE) & """> "
'
''                Print #nFile, "<img src=" & Chr(34) & "../../img/blank.gif" & Chr(34) & " " _
''                                 & "name=" & Chr(34) & "fld" & rsCRFElement!CRFElementId & "_img" & Chr(34) & ">"
'
''                Print #nFile, "</div> "
'
'                 Print #nFile, printCaption(rsCRFElement, rsStudyDefinition, rsDataItem, _
'                        bDispNum, bNetscape, nDefaultFontSize, lDefaultFontColour, _
'                        sDefaultFontFamily, nDefaultFontWeight, nDefaultFontStyle) & vbCrLf
'
'             End If
'
'
'
'            'End If
'            'debug element id
'            rsCRFElement.MoveNext
'
'
'        Loop
'
'        Print #nFile, "</form>"
''        Print #nFile, "<Img src=""dummy_img.gif"">"
'        Print #nFile, "</body>"
'        Print #nFile, "</html>"
''        Print #nFile, "<!--#include file=""CloseDataConnection.txt"" -->"
''        'MLM 18/9/00: close statement block for role function check
'
'        'this prints the second case statement.
'        Print #nFile, "<% Case 2:%>" & vbCrLf _
'                    & "<html>" & vbCrLf _
'                    & "<head>" & vbCrLf _
'                    & "<link rel=" & Chr(34) & "stylesheet" & Chr(34) & " href=" & Chr(34) & "../../style/MACRO1.css" & Chr(34) & " type=" & Chr(34) & "text/css" & Chr(34) & ">" & vbCrLf _
'                    & "<meta http-equiv=" & Chr(34) & "Pragma" & Chr(34) & " content=" & Chr(34) & "no-cache" & Chr(34) & ">" & vbCrLf
'
'        Print #nFile, "<script language=" & Chr(34) & "javascript" & Chr(34) & ">" & vbCrLf _
'                    & "function showMsg(){" & vbCrLf _
'                    & "<%If (nxt = 1) Then%>" & vbCrLf _
'                    & "alert(" & Chr(34) & "This is the last form in the current visit." & Chr(34) & ");" & vbCrLf _
'                    & "<%Else%>" & vbCrLf _
'                    & "alert(" & Chr(34) & "This is the first form in the current visit." & Chr(34) & ");" & vbCrLf _
'                    & "<%End If%>" & vbCrLf _
'                    & "window.parent.frames[0].goMnu(4);" & vbCrLf _
'                    & "}</script>" & vbCrLf _
'                    & "</head>" & vbCrLf _
'                    & "<body onload=" & Chr(34) & "showMsg();" & Chr(34) & ">" & vbCrLf _
'                    & "</body>" & vbCrLf _
'                    & "</html>"
'
'        Print #nFile, "<% Case 3:%>" & vbCrLf _
'                    & "<html>" & vbCrLf _
'                    & "<head>" & vbCrLf _
'                    & "<link rel=" & Chr(34) & "stylesheet" & Chr(34) & " href=" & Chr(34) & "../../style/MACRO1.css" & Chr(34) & " type=" & Chr(34) & "text/css" & Chr(34) & ">" & vbCrLf _
'                    & "<meta http-equiv=" & Chr(34) & "Pragma" & Chr(34) & " content=" & Chr(34) & "no-cache" & Chr(34) & ">" & vbCrLf
'
'
'        Print #nFile, "<script language=" & Chr(34) & "javascript" & Chr(34) & ">" & vbCrLf _
'                    & "function showMsg(){" & vbCrLf _
'                    & "var bNew=(confirm(" & Chr(34) & "Creating new cycling form. Click OK to continue, or CANCEL to move to next non-cycling form" & Chr(34) & "));" & vbCrLf _
'                    & "<%If (url(0) <> " & Chr(34) & Chr(34) & ") and (url(1) <> " & Chr(34) & Chr(34) & ") Then%>" & vbCrLf _
'                    & "if (bNew){" _
'                    & "window.location.replace(" & Chr(34) & "<%=url(1)%>" & Chr(34) & ")" _
'                    & "}else{" _
'                    & "window.location.replace(" & Chr(34) & "<%=url(0)%>" & Chr(34) & ")" _
'                    & "}" & vbCrLf _
'                    & "<%Else%>" & vbCrLf _
'                    & "if (bNew){" & "window.location.replace(" & Chr(34) & "<%=url(1)%>" & Chr(34) & ")" & "}else{" _
'                    & "alert(" & Chr(34) & "This is the last form in the current visit." & Chr(34) & ");" & vbCrLf _
'                    & "window.parent.frames[0].goMnu(4);}" & vbCrLf _
'                    & "<%End If%>" & vbCrLf _
'                    & "}</script>" & vbCrLf _
'                    & "</head>" & vbCrLf _
'                    & "<body onload=" & Chr(34) & "showMsg();" & Chr(34) & ">" & vbCrLf _
'                    & "</body>" & vbCrLf _
'                    & "</html>"
'
'
'        Print #nFile, "<% Case Else:%>" & vbCrLf _
'                    & "<html>" & vbCrLf _
'                    & "<head>" & vbCrLf _
'                    & "</head>" & vbCrLf _
'                    & "<meta http-equiv=" & Chr(34) & "Pragma" & Chr(34) & " content=" & Chr(34) & "no-cache" & Chr(34) & ">" & vbCrLf _
'                    & "</head>" & vbCrLf _
'                    & "<body onload=" & Chr(34) & "window.parent.frames[0].goMnu(4);" & Chr(34) & ">" & vbCrLf _
'                    & "</body>" & vbCrLf _
'                    & "</html>"
'
'
'
'        Print #nFile, "<%End Select"
'
'
'        Print #nFile, "else%>" & vbCrLf _
'                    & "<html>" & vbCrLf _
'                    & "<head>" & vbCrLf _
'                    & "<link rel=" & Chr(34) & "stylesheet" & Chr(34) & " href=" & Chr(34) & "../../style/MACRO1.css" & Chr(34) & " type=" & Chr(34) & "text/css" & Chr(34) & ">" & vbCrLf _
'                    & "<meta http-equiv=" & Chr(34) & "Pragma" & Chr(34) & " content=" & Chr(34) & "no-cache" & Chr(34) & ">" & vbCrLf _
'                    & "</head>" & vbCrLf _
'                    & "<body>" & vbCrLf _
'                    & "<table width=" & Chr(34) & "100%" & Chr(34) & " height=" & Chr(34) & "95%" & Chr(34) & ">" & vbCrLf _
'                    & "<tr><td align=" & Chr(34) & "center" & Chr(34) & " class=" & Chr(34) & "MESSAGE" & Chr(34) & ">" _
'                    & "This subject has not been correctly locked. <a href=" & Chr(34) & "../general/selectDatabase.asp" & Chr(34) & " target=" & Chr(34) & "_top" & Chr(34) & ">Please re-select the subject</a>.</td></tr>" & vbCrLf _
'                    & "</table>" & vbCrLf _
'                    & "</body>" & vbCrLf _
'                    & "</html>"
'
'
'        Print #nFile, "<%End if" & vbCrLf _
'                    & "Else%>" & vbCrLf _
'                    & "<html>" & vbCrLf _
'                    & "<head>" & vbCrLf _
'                    & "<link rel=" & Chr(34) & "stylesheet" & Chr(34) & " href=" & Chr(34) & "../../style/MACRO1.css" & Chr(34) & " type=" & Chr(34) & "text/css" & Chr(34) & ">" & vbCrLf _
'                    & "<meta http-equiv=" & Chr(34) & "Pragma" & Chr(34) & " content=" & Chr(34) & "no-cache" & Chr(34) & ">" & vbCrLf _
'                    & "</head>" & vbCrLf _
'                    & "<body>" & vbCrLf _
'                    & "<table width=" & Chr(34) & "100%" & Chr(34) & " height=" & Chr(34) & "95%" & Chr(34) & ">" & vbCrLf _
'                    & "<tr><td align=" & Chr(34) & "center" & Chr(34) & " class=" & Chr(34) & "MESSAGE" & Chr(34) & ">" _
'                    & "You do not have permission to use Web Data Entry.</td></tr>" & vbCrLf _
'                    & "</table>" & vbCrLf _
'                    & "</body>" & vbCrLf _
'                    & "</html>"
'
'                    'REM 05/11/01 Bug fix 122: added oIo.FinishoALM VB function call and moved set oIo = nothing to end of page
'        Print #nFile, "<%End If" & vbCrLf _
'                    & "if session(" & Chr(34) & "usrStudyDefOK" & Chr(34) & ") = " & Chr(34) & "True" & Chr(34) & " Then" & vbCrLf _
'                    & "set session(" & Chr(34) & "usrStudyDef" & Chr(34) & ") = oIo.FinishoALM(session(" & Chr(34) & "usrStudyDef" & Chr(34) & "))" & vbCrLf _
'                    & "set session(" & Chr(34) & "usrStudyDef" & Chr(34) & ") = nothing" & vbCrLf _
'                    & "session(" & Chr(34) & "usrStudyDefOK" & Chr(34) & ") = " & Chr(34) & "False" & Chr(34) & vbCrLf _
'                    & "End If" & vbCrLf _
'                    & "Set oIo = Nothing%>"
'
''        Print #nFile, "<!-- #Include file=""formCyclelist.asp"" -->"
''        Print #nFile, "<!--#include file=""eventHandlers.asp""-->"
''        'MLM 26/10/00 added "Change Data" role function to arguments of GetFormArezzoDetails
''        Print #nFile, "<% response.write session(""P"").GetFormArezzoDetails(session(""sRoleCode""),request.querystring(""trial"") ,request.querystring(""VersionId""), request.querystring(""trialname""), request.querystring(""site"") ,clng(request.querystring(""person"")) , cint(request.querystring(""CRFPageId"")), clng(request.querystring(""crfpagetaskid"")), lCase(request.querystring(""VisitCode"")), lCase(request.querystring(""CRFPageCode"")), session(""F5003"") ) %>"
''
''        Print #nFile, "<script language =""javascript"">"
''        'Print #nFile, "   setInitialCaptions();" & vbNewLine
''        Print #nFile, "function fnInitialisePage()"
''        Print #nFile, "{"
''        Print #nFile, "   ApplyInitialFormat();"
''        Print #nFile, "   DisplayForm();"
''        Print #nFile, "   load_Links();"
''        Print #nFile, "}"
''        Print #nFile, "</script>"
'        Close #nFile
'
'            FormDefinition rsCRFPages!ClinicalTrialId, rsCRFPages!VersionId, rsCRFPages
'            'End If
'            '### debug pageId
'        rsCRFPages.MoveNext
'
'    Loop
'
'
'    If Not (rsCRFPages.BOF And rsCRFPages.EOF) Then
'        rsCRFPages.Close
'        Set rsCRFPages = Nothing
'        If bNoOfelements = True Then
'            rsCRFElement.Close
'            Set rsCRFElement = Nothing
'            rsDataItem.Close
'            Set rsDataItem = Nothing
'        End If
'    End If
'
'    CreateHTMLComponents vClinicalTrialId, vVersionId
'
'    '   ATN 8/12/99
'    '   Added call to create netscape files here, so that its always done.
'    'CreateNetscapefiles vClinicalTrialId, vVersionId
'
'
'    'destroy recordsets
'    rsStudyDefinition.Close
'    Set rsStudyDefinition = Nothing
'
'    Exit Sub
'
'ErrHandler:
'  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "CreateHTMLFIles", "modHTML")
'        Case OnErrorAction.Ignore
'            Resume Next
'        Case OnErrorAction.Retry
'            Resume
'        Case OnErrorAction.QuitMACRO
'            Call ExitMACRO
'            Call MACROEnd
'   End Select
'
'End Sub

'ic  21/08/01   commented out CreateHTMLFiles sub, rewrote for new JSVE structure
''-------------------------------------------------------------------------------------------'
'Public Sub CreateHTMLFiles(ByVal vClinicalTrialId As Long, _
'                           ByVal vVersionId As Integer)
''-------------------------------------------------------------------------------------------'
'' Creates HTML eform displayed in the browser.
''-------------------------------------------------------------------------------------------'
''Revisions:
''
''MLM 18/09/00   Check user has logged in and has the "View data" role function
''MLM 26/10/00   added "Change Data" role function to arguments of GetFormArezzoDetails
''MLM 09/11/00   Check user has permission to use site
''JL  16/11/00   Removed unneccasary server side code.
''MLM 21/11/00   Also check subject locking
''ic  01/12/00   added condition so that if the picture cannot be found it is omitted
''ic  05/12/00   added declaration to hold list of copied images
''               added code to copy images to web folder
''MLM 07/12/00   Use study definition's default background colour if a form's colour isn't specifed.
''JL  18/12/00    Added an extra argument current user's role code to getFormArezzoDetails for
''                       disabling a question which needs authorisation.
''ic  21/12/00   added onblur event to radios so that answers are evaluated when tabbing between fields as well as clicking
''ic  12/02/01   removed gsDOCUMENTS_PATH constant from picture code as it was being written in the html
''JL  12/02/01   Changes to function JavascriptDisableCaptions.
''jl  12/03/01   Changes for oracle.
''-------------------------------------------------------------------------------------------'
''Exit Sub
''-------------------------------------------------------------------------------------------'
'
'Dim rsStudyDefinition As ADODB.Recordset
'Dim rsCRFPages As ADODB.Recordset
'Dim rsCRFElement As ADODB.Recordset
'Dim rsDataItem As ADODB.Recordset
'Dim rsDerivedValues As ADODB.Recordset
'Dim sSQL As String
'Dim sFileName As String
'Dim nFileNumber As Integer
'
'Dim lCRFPageId As Long
'Dim lDataItemId As Long
'Dim rsValueData As ADODB.Recordset
'Dim sDefaultFontAttr As String
'Dim sPositionAttr As String
'Dim sDisabledSetting As String
'Dim bNetscape As Boolean
'Dim nRCount As Integer 'recordcount
'Dim bNoOfelements As Boolean
'Dim bNoOfpages As Boolean
'Dim lControlLeft As Long
'Dim lImageXposition As Long
'Dim bDisplayNumbers As Boolean
'
'Dim nLengthOfItem As Integer     'JL 16/06/00 Added for use by radios & pop-ups
'Dim nMaxLengthSoFar As Integer   'JL 16/06/00 Added for use by radios & pop-ups
'Dim nTextBoxCharWidth As Integer
'Dim nMaxLength As Integer
'Dim sComments As String
'Dim sUnitOfMeasurement As String
'
'Dim nDefaultFontSize As Integer
'Dim lDefaultFontColour As Long
'Dim sDefaultFontFamily As String
'Dim nDefaultFontWeight As Integer   'Bold/medium 'Default 0=medium, 1=bold
'Dim nDefaultFontStyle As Integer    'Default 0=normal, 1=Italic
'
'Dim nNumberOfSkips As Integer
'Dim nNumberOfRadios As Integer
'
'Dim sQuestionRoleCode As String 'For Question
'Dim sDerivation As String
'
'Const sDISABLED = "DISABLED=""Disabled"""
'Const sENABLED = ""
'
'Const bIEflag = 0
'Const nMINIMALCONTROLWIDTH = 70
'Const nMINIMAL_POPUP_WIDTH = 40
'Dim ElementWidth As Integer
'Dim oPicture As Picture
'Dim nMaxImageLength As Integer
'
'
''ic 05/12/00
''added declaration to hold list of images files already copied
'Dim sCpdImages As String
'
'
'    'On Error GoTo ErrHandler
'
'    bNetscape = False
'
'    nFileNumber = FreeFile
'    bNoOfelements = True
'    bNoOfpages = True
'
'    'get Default font attributes
'    sSQL = "SELECT * FROM StudyDefinition" _
'            & " WHERE ClinicalTrialId = " & vClinicalTrialId _
'            & " AND VersionId = " & vVersionId
'
'    Set rsStudyDefinition = New ADODB.Recordset
'    rsStudyDefinition.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
'
'    If rsStudyDefinition.RecordCount <> 0 Then
'        sDefaultFontAttr = "<style type=" & Chr$(34) & "text/css" & Chr$(34) & "> " & _
'            "div    {position:absolute;width:auto;height:auto}" & _
'            "p      {color:" & GetFontColour(rsStudyDefinition!DefaultFontColour) & ";font-family:" & _
'                    rsStudyDefinition!DefaultFontName & ";font-size:" & rsStudyDefinition!DefaultFontSize & _
'                    " " & "pt}  " & _
'            "input.T { font-size: 10pt;}" & _
'                    "</style>"
'
'     End If
'
'     nDefaultFontSize = rsStudyDefinition!DefaultFontSize
'     lDefaultFontColour = rsStudyDefinition!DefaultFontColour
'     sDefaultFontFamily = rsStudyDefinition!DefaultFontName
'     nDefaultFontWeight = rsStudyDefinition!DefaultFontBold
'     nDefaultFontStyle = rsStudyDefinition!DefaultFontItalic
'
'    'Identify a trial
'    sSQL = "SELECT * FROM CRFPage " _
'            & " WHERE ClinicalTrialId = " & vClinicalTrialId _
'            & " AND VersionId = " & vVersionId
'
'    Set rsCRFPages = New ADODB.Recordset
'    rsCRFPages.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
'
'    If rsCRFPages.EOF Then
'        bNoOfpages = False
'    End If
'    '---------------------------------------------------------------------------'
'    '##    ###                   *** DEBUG PAGE  ***                              '
'    '---------------------------------------------------------------------------'
'    Do While Not rsCRFPages.EOF
'    'JL 01/06/00 Added numbers switch.
'        'If rsCRFPages!CRFPageId = 10001 Or rsCRFPages!CRFPageId = 10046 Then '#1
'
'        'If rsCRFPages!CRFPageId = 10019 Then
'
'        If rsCRFPages!displayNumbers = 1 Then
'            bDisplayNumbers = True
'        Else
'            bDisplayNumbers = False
'        End If
'        lCRFPageId = rsCRFPages!CRFPageId
'
'        sSQL = "SELECT * FROM CRFElement where " _
'                & "CRFPageId =" & lCRFPageId & " AND " & _
'                "ClinicalTrialId =" & vClinicalTrialId & _
'                " AND VersionId=" & vVersionId & " ORDER BY CRFElementId"
'
'        Set rsCRFElement = New ADODB.Recordset
'        rsCRFElement.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
'
'        'SDM 27/01/00 SR2794
'        sFileName = gsWEB_HTML_LOCATION & LTrim(str(rsCRFPages!ClinicalTrialId)) _
'                            & "_" & LTrim(str(rsCRFPages!CRFPageId)) & "_" & gUser.DatabaseName & ".asp"
'
'        Debug.Print "In CreateHTMLFiles - CRFpage" & sFileName
'        Open sFileName For Output As #nFileNumber
'
'        Print #nFileNumber, "<%@ LANGUAGE = VBScript %>" & vbNewLine & _
'                            "<% Option Explicit" & vbNewLine & _
'                            "Dim cnnMACRO    " & vbNewLine & _
'                            "dim rsSiteUser" & vbNewLine & _
'                            "dim sErrMsg" & vbNewLine & _
'                            "Response.Expires = 0 %>"
'
'        'MLM 18/9/00: check user's permissions
'        Print #nFileNumber, "<!-- #include file=""CheckLogin.asp"" -->"
'        Print #nFileNumber, "<!-- #include file=""format.htm"" -->"
'        Print #nFileNumber, "<!-- #include file=""arezzo.htm"" -->"
'        'MLM 10/11/00 added OpenDataConnection.txt so that site permission can be checked. This used to be #included in DataValues.asp
'        Print #nFileNumber, "<!-- #Include file=""OpenDataConnection.txt"" -->"
'
'        'retrieve site permission
'        Print #nFileNumber, "<%set rsSiteUser = cnnMACRO.Execute(""SELECT COUNT(*) AS Permission FROM SiteUser" & _
'            " WHERE Site = '"" & Request.QueryString(""site"") & ""'" & _
'            " AND UserCode = '"" & session(""UserName"") & ""'"")"
'        Print #nFileNumber, "sErrMsg = session(""P"").OpenSubject(Request.QueryString(""trial""), Request.QueryString(""site""), Request.QueryString(""person""))"
'
'        Print #nFileNumber, "if not session(""F5002"") then%>"
'        Print #nFileNumber, "    <html><body bgcolor=white><table width=100% height=100%><tr><td align=center>"
'        Print #nFileNumber, "        You do not have permission to view subject data."
'        Print #nFileNumber, "    </td></tr></table></body></html>"
'
'        'MLM 9/11/00: check user can view this site
'        'jl  12/03/01 use cint for oracle.
'        Print #nFileNumber, "<%elseif cint(rsSiteUser(""Permission"")) = 0 then%>"
'        Print #nFileNumber, "    <html><body bgcolor=white><table width=100% height=100%><tr><td align=center>"
'        Print #nFileNumber, "        You do not have permission to access <%=Request.QueryString(""site"")%>."
'        Print #nFileNumber, "    </td></tr></table></body></html>"
'
'        'MLM 21/11/00 also check file locking
'        Print #nFileNumber, "<%elseif sErrMsg > """" then%>"
'        Print #nFileNumber, "    <html><body bgcolor=white><table width=100% height=100%><tr><td align=center>"
'        Print #nFileNumber, "        <%=sErrMsg%>"
'        Print #nFileNumber, "    </td></tr></table></body></html>"
'        Print #nFileNumber, "<%else%>" & vbNewLine
'
'        Print #nFileNumber, "<html>"
'        Print #nFileNumber, "<head>"
'
'        Print #nFileNumber, "<title>" & Chr$(34) & "Form : "; rsCRFPages!CRFTitle&; Chr$(34) & "</title>"
'
'        'Print #nFileNumber, "<!-- #Include file=" & Chr(34) & "LogoutUser.txt" & Chr(34) & " -->"
'        Print #nFileNumber, "<!-- #Include file=" & Chr(34) & "HandleIEDocument.asp" & Chr(34) & " -->"
'
'        Print #nFileNumber, sDefaultFontAttr  'STYLE sheet
'        Print #nFileNumber, "</head>"
'
'    '   ATN 7/12/99
'    '   Variable definitions will now be set up in the web rde dll.
'    '   added CRFElement.CRFElementId.
'        'JL 8/09/00. Added UnitOfMeasurement in the select for passing to new function JavascriptDisableCaptions.
'        sSQL = "SELECT DataItemCode, UnitOfMeasurement" _
'                & " FROM DataItem, CRFElement where " _
'                & " CRFPageId =" & lCRFPageId _
'                & " AND DataItem.DataItemId = CRFElement.DataItemId " _
'                & " AND DataItem.ClinicalTrialId = CRFElement.ClinicalTrialId " _
'                & " AND DataItem.VersionId = CRFElement.VersionId " _
'                & " AND DataItem.ClinicalTrialId =" & vClinicalTrialId _
'                & " AND DataItem.VersionId=" & vVersionId
'
'        Set rsDataItem = New ADODB.Recordset
'        rsDataItem.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
'
''       'JL 26/10/00
'        If Not rsDataItem.EOF Then
'            If IsNull(rsDataItem!UnitOfMeasurement) Then
'                sUnitOfMeasurement = RemoveNull(rsDataItem!UnitOfMeasurement)
'            End If
'        End If
'
'        'start the HTML body
'        If rsCRFPages!BackgroundColour = 0 Then
'            'MLM 7/12/00 Use study's default background colour (not white) if form's colour is unspecified
'            Print #nFileNumber, "<body bgcolor=#" & GetFontColour(rsStudyDefinition!DefaultCRFPageColour) & " onUnload=""disableLinks();"" onLoad=""fnInitialisePage();"">"
'        Else
'            Print #nFileNumber, "<body bgcolor=""#" & GetFontColour(rsCRFPages!BackgroundColour) & """  onUnload=""disableLinks();"" onLoad=""fnInitialisePage();"">"
'        End If
'
'        Print #nFileNumber, "<script language =""javascript"">" & vbNewLine & _
'            "var f = window.parent.frames[2].document.myform ;"
'
'        Print #nFileNumber, "</script>"
'
'        Print #nFileNumber, "<form name=""myform"" method=" & Chr$(34) & "post" & Chr$(34) & " action=" & Chr$(34) & "SaveForm.asp" & Chr$(34) & ">"
'        'MLM 10/11/00 Leave DataValues here because it refers to the form object.
'        Print #nFileNumber, "<!-- #Include file=""DataValues.asp"" -->"
'
'        If rsCRFElement.EOF Then
'            bNoOfelements = False
'        End If
'
'        'JL 07/09/00. Build Javascript function to dim captions grey if disabled.
'        nNumberOfSkips = 0
'        nNumberOfRadios = 0
'        Do While Not rsCRFElement.EOF
'            If Not IsNull(rsCRFElement!SkipCondition) And rsCRFElement!SkipCondition > "" Then
'                nNumberOfSkips = nNumberOfSkips + 1
'            End If
'            If rsCRFElement!ControlType = gsOPTION_BUTTONS Or rsCRFElement!ControlType = gsPUSH_BUTTONS Then
'                nNumberOfRadios = nNumberOfRadios + 1
'            End If
'
'            lDataItemId = rsCRFElement!DataItemId
'
'            rsCRFElement.MoveNext
'        Loop
'
'        If nNumberOfRadios > 0 Then
'            Print #nFileNumber, JavascriptUpdateRadioValues(vClinicalTrialId, vVersionId, lCRFPageId)
'        End If
'
'        If nNumberOfSkips > 0 Then
'                Print #nFileNumber, JavascriptDisableCaptions(rsStudyDefinition, vClinicalTrialId, vVersionId, lCRFPageId, bDisplayNumbers, _
'                                    sUnitOfMeasurement, bNetscape)
'        Else
'                'JL 12/02/01 Bug. still need a function as getCaption is called in JS functions (e.g. DisplayForm())
'                Print #nFileNumber, "<script language =""javascript"">" & vbNewLine & _
'                                            vbTab & "function getCaption(){ return """" }" & vbNewLine & "</script>"
'        End If
'        'JL 26/10/00
'        If Not bNoOfelements = False Then
'            rsCRFElement.MoveFirst
'        End If
'        'CRFELEMENT
'        Do While Not rsCRFElement.EOF
'        '-------------------------------------------------------------------------------------------------
'        'JL 16/08/00. Added checking of question RoleCode
'        '-------------------------------------------------------------------------------------------------
'
'        sQuestionRoleCode = RemoveNull(rsCRFElement!RoleCode)
'
'        'JL 16/11/00. Remove uneccesary server side code.
'        'Print #nFileNumber, "<% sQuestionStatus = GetStatusText(getCurrentStatus(lthisId)) %>"
'
'        lImageXposition = 0
'
'            lDataItemId = rsCRFElement!DataItemId
'
'                sSQL = "SELECT * FROM DataItem WHERE " _
'                    & " ClinicalTrialId =" & vClinicalTrialId _
'                    & " AND DataItemId =" & lDataItemId _
'                    & " AND VersionID =" & vVersionId
'
'                Set rsDataItem = New ADODB.Recordset
'                rsDataItem.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
'                If lDataItemId > 0 Then
'                Debug.Print rsDataItem!DataItemCode
'                    'Remove any Nulls right at the start.
'                    sDerivation = RemoveNull(rsDataItem!Derivation)
'
'                    If sDerivation > "" Then
'                        sDisabledSetting = sDISABLED
'                    Else
'                        sDisabledSetting = sENABLED
'                    End If
'
'                End If
'
'
'        'lThisId used as argument for routines in dataValues.asp
'        'Dont print for visual elements
'        If rsCRFElement!ControlType <> gsCOMMENT And rsCRFElement!ControlType <> gsPICTURE And _
'                    rsCRFElement!ControlType <> gsLINE Then
'
'            Print #nFileNumber, "<% lthisId= " & rsCRFElement!CRFElementId & "%>"
'        End If
'
'        '---------------------------------------------------------------------------'
'        '                               *** COMMENTS LABEL  ***                     '
'        '---------------------------------------------------------------------------'
'        If rsCRFElement!ControlType = ControlType.Comment Then
'            sComments = ""
'                    Print #nFileNumber, "<div id=" & Chr$(34) & "L" & rsCRFElement!CRFElementId & Chr$(34) & " Style=" & Chr$(34) & _
'                                        ";top:" & CLng(rsCRFElement!Y / nIeYSCALE) & ";left:" & CLng(rsCRFElement!X / nIeXSCALE) & ";z-index:1" & Chr$(34) & ">"
'
'                    Print #nFileNumber, "<p>"
'
'                    bDisplayNumbers = False
'
'                    sComments = sComments & "<font"
'                    sComments = sComments & printFontAttributes(rsCRFElement!FontColour, rsCRFElement!FontSize, rsCRFElement!FontName _
'                            , rsStudyDefinition!DefaultFontColour, rsStudyDefinition!DefaultFontSize, rsStudyDefinition!DefaultFontName, bNetscape) & _
'                    FieldAttributes_Explorer(rsCRFElement!FontBold, rsCRFElement!FontItalic, rsCRFElement!FieldOrder, rsCRFElement!Caption _
'                                , bDisplayNumbers)
'                    sComments = sComments & "</font>"
'
'                    Print #nFileNumber, sComments
'
'                    Print #nFileNumber, "</p>"
'                    Print #nFileNumber, "</div>"
'        '---------------------------------------------------------------------------'
'        '                                  *** Hidden  ***                          '
'        '---------------------------------------------------------------------------'
'        ElseIf rsCRFElement!Hidden = 1 Then
'                    Print #nFileNumber, "<INPUT TYPE=""HIDDEN"" NAME=""hid" & LCase(rsDataItem!DataItemCode) & """>"
'                    Print #nFileNumber, CreateFormObject(rsDataItem, rsCRFElement, rsStudyDefinition)
'
'        '---------------------------------------------------------------------------'
'        '                                  *** LINE  ***                            '
'        '---------------------------------------------------------------------------'
'        ElseIf rsCRFElement!ControlType = ControlType.Line Then
'                Print #nFileNumber, "<div id=" & Chr$(34) & "L" & rsCRFElement!CRFElementId & Chr$(34) & " style=" & Chr$(34) & _
'                                    "20px;left:0;top:" & CLng(rsCRFElement!Y / nIeYSCALE) & Chr$(34) & " >"
'                Print #nFileNumber, "<hr>"
'
'                Print #nFileNumber, "</div>"
'        '---------------------------------------------------------------------------'
'        '                           *** TEXTBOX or DATE  ***                        '
'        '---------------------------------------------------------------------------'
'        ElseIf rsCRFElement!ControlType = ControlType.TextBox Or rsCRFElement!ControlType = ControlType.Calendar _
'            Or rsCRFElement!ControlType = ControlType.RichTextBox Then
'        'jl RichTextBoxes are not supported in the web, but display anyway for now.
'                            'Jl 30/05/00 Code transfered to new printCaption.
'                            '------------------------------------'
'                            '        *** PRINT CAPTION  ***      '
'                            '------------------------------------'
'                            '07/08/00
'            Print #nFileNumber, printCaption(rsCRFElement, rsStudyDefinition, rsDataItem, _
'                     bDisplayNumbers, bNetscape, nDefaultFontSize, lDefaultFontColour, _
'                     sDefaultFontFamily, nDefaultFontWeight, nDefaultFontStyle)
'
'                    If IsNull(rsDataItem!DataItemFormat) Then
'                        SetElementWidth "", rsDataItem!DataItemLength, nMaxLength, nTextBoxCharWidth
'                    Else
'                        SetElementWidth rsDataItem!DataItemFormat, rsDataItem!DataItemLength, nMaxLength, nTextBoxCharWidth
'                    End If
'
'                '2##
'                Print #nFileNumber, "<div id=" & Chr$(34) & "L" & rsCRFElement!CRFElementId & Chr$(34) & " style=" & Chr$(34) & _
'                                    ";left:" & CLng(rsCRFElement!X / nIeXSCALE) & ";top:" & CLng(rsCRFElement!Y / nIeYSCALE) & ";z-index:2" & Chr$(34) & ">"
'
'                'JL 07/08/00 Replacing use of R recordset.
'                'If a disabled control no need for event handlers.
'
'                    Print #nFileNumber, "<input type=" & Chr$(34) & "TEXT" & Chr$(34) & " Disabled tabindex=" & Chr$(34) & rsCRFElement!FieldOrder & Chr$(34) & _
'                                    " name=""txt" & LCase(rsDataItem!DataItemCode) & """ size=" & nTextBoxCharWidth & _
'                                    " maxlength=" & Chr$(34) & nMaxLength & Chr$(34) & " onblur=""RefreshForm(f.txt" & LCase(rsDataItem!DataItemCode) & ");""" & _
'                                    " onfocus=""setFocusColour(f.txt" & LCase(rsDataItem!DataItemCode) & ");"">"
'
'
'    '   ATN 3/12/99
'    '   Javascript validation now created at form level at run-time
'    '           'Populate the Form Objects property value
'                'Print #nFileNumber, GetDataValue(rsDataItem!DataItemCode)
'                Print #nFileNumber, CreateFormObject(rsDataItem, rsCRFElement, rsStudyDefinition)
'                Print #nFileNumber, "</div>"
'
''                If nMaxLength > nTextBoxCharWidth Then
''                    nTextBoxCharWidth = nMaxLength
''                End If
'
'                'jl 16/06/00 Replaced lControlLeft with CLng(rsCRFElement!CaptionX / nIeXSCALE)
'                lImageXposition = CalulateStatusXPosition(nTextBoxCharWidth, CLng(rsCRFElement!X / nIeXSCALE), gsTEXT_BOX)
'
'
'                'JL 13/07/00.
'                'Always call SetStatusImage, if status=Requested then the image is set to a dummy transparent image.
'                Print #nFileNumber, "<div style=""Left:" & lImageXposition & ";top:" & CLng(rsCRFElement!Y / nIeYSCALE) & """> " & vbNewLine
'
'                Print #nFileNumber, "    <%= SetStatusImage(getCurrentStatus(lthisId), """ & LCase(rsDataItem!DataItemCode) & """) %>" & vbNewLine
'
'                Print #nFileNumber, "</div> "
'            '---------------------------------------------------------------------------'
'            '                           *** MULTIMEDIA CONTROL  ***                     '
'            '---------------------------------------------------------------------------'
'            ElseIf rsCRFElement!ControlType = ControlType.Attachment Then
'
'                                'Jl 30/05/00 Code transfered to new printCaption.
'                                '------------------------------------'
'                                '        *** PRINT CAPTION  ***      '
'                                '------------------------------------'
'                '07/08/00
'                Print #nFileNumber, printCaption(rsCRFElement, rsStudyDefinition, rsDataItem, _
'                         bDisplayNumbers, bNetscape, nDefaultFontSize, lDefaultFontColour, _
'                         sDefaultFontFamily, nDefaultFontWeight, nDefaultFontStyle)
'
'                'JL 23/05/00
'                Print #nFileNumber, "<div id=" & Chr$(34) & "L" & rsCRFElement!CRFElementId & Chr$(34) & " style=" & Chr$(34) & _
'                    "left:" & CLng(rsCRFElement!X / nIeXSCALE) & ";top:" & CLng(rsCRFElement!Y / nIeYSCALE) & Chr$(34) & ">" & _
'                    "<input Disabled type=" & Chr$(34) & "file" & Chr$(34) & " TabIndex = " & Chr$(34) & rsCRFElement!FieldOrder & Chr$(34) & _
'                    " Class=""T"" name=""att" & LCase(rsDataItem!DataItemCode) & """ onblur=""RefreshForm(f.att" & LCase(rsDataItem!DataItemCode) & ");""" & _
'                    " onfocus=""setFocusColour(f.att" & LCase(rsDataItem!DataItemCode) & ");"">"
'
'                Print #nFileNumber, "</div>"
'                'If a Multimedia Control allow for more space on the left.
'
'                lImageXposition = CalulateStatusXPosition(nTextBoxCharWidth, CLng(rsCRFElement!X / nIeXSCALE), gsATTACHMENT)
'
'                'JL 07/04/00 Print out Question Status Image.
'                'Prints out status image if there is a record in the DataItemResponse table.
'
'                'JL 13/07/00. Always call SetStatusImage.
'                 Print #nFileNumber, "<div style=""Left:" & lImageXposition & ";top:" & CLng(rsCRFElement!Y / nIeYSCALE) & """> " & vbNewLine
'
'                 Print #nFileNumber, "<%= SetStatusImage(getCurrentStatus(lthisId), """ & LCase(rsDataItem!DataItemCode) & """) %>" & vbNewLine
'
'                 Print #nFileNumber, "</div> "
'
'
'                Print #nFileNumber, CreateFormObject(rsDataItem, rsCRFElement, rsStudyDefinition)
'
'            '---------------------------------------------------------------------------'
'            '                         *** RADIOS and OPTION BOXES  ***                  '
'            '---------------------------------------------------------------------------'
'            ElseIf rsCRFElement!ControlType = ControlType.OptionButtons Or rsCRFElement!ControlType = ControlType.PushButtons Then 'check gsPUSH_BUTTONS
'               'Get values for option boxs etc
'
'                sSQL = "SELECT * FROM ValueData WHERE" & _
'                    " DataItemId = " & lDataItemId & " AND" & _
'                    " ClinicalTrialId = " & vClinicalTrialId & "AND " & _
'                    " VersionId = " & vVersionId
'
'                Set rsValueData = New ADODB.Recordset
'                rsValueData.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
'    '   ATn 26/4/99
'    '   Added check for at least one value being defined for a category data item
'    'lImageXposition
'                If rsValueData.RecordCount > 0 Then
'
'                'Jl 30/05/00 Code transfered to new printCaption.
'
'                    Print #nFileNumber, "<div id=" & Chr$(34) & "L" & rsCRFElement!CRFElementId & Chr$(34) & _
'                                       " style=" & Chr$(34) & ";left:" & _
'                                       CLng(rsCRFElement!X / nIeXSCALE) & ";top:" & CLng(rsCRFElement!Y / nIeYSCALE) & ";z-index:2" & Chr$(34) & ">"
'
'                     rsValueData.MoveLast
'
'                    Print #nFileNumber, "<% redim v(" & rsValueData.RecordCount & ")"
'                    Print #nFileNumber, "redim c(" & rsValueData.RecordCount & ")"
'
'                    nRCount = 1
'
'                    nLengthOfItem = 0
'                    nMaxLengthSoFar = 0
'                    rsValueData.MoveFirst
'                    Do While Not rsValueData.EOF 'For Each Radio Item
'
'                        'JL 06/02/00 Replaced rsValueData.AbsolutePosition with nRCount for recordcount, see notes.
'                        Print #nFileNumber, "c(" & nRCount & ") =""" & rsValueData!ValueCode & """"
'                        Print #nFileNumber, "v(" & nRCount & ") =""" & rsValueData!ItemValue & """"
'                        nRCount = nRCount + 1
'                        'JL 16/06/00
'                        nLengthOfItem = Len(rsValueData!ItemValue)
'                        If nLengthOfItem > nMaxLengthSoFar Then
'                            nMaxLengthSoFar = nLengthOfItem
'                        End If
'                        rsValueData.MoveNext
'                    Loop
'
'                    'JL 16/08/00 Added extra argument to Radio for disabling control.
'                    'JL 16/08/00 Added Check for authoisation code.
'
'                        Print #nFileNumber, " Radio " & """opt" & LCase$(rsDataItem!DataItemCode) & """, """ & LCase$(rsDataItem!DataItemCode) & """, " & rsCRFElement!FieldOrder & ",getResponseValue(" & rsCRFElement!CRFElementId & ") %>" & vbNewLine
'
'
'                '   ATN 3/12/99
'                '   Javascript validation now created at form level at run-time
'                '   Print #nFileNumber, JavascriptValidation(rsDataItem!DataItemCode, rsDataItem!DataType, rsDataItem!MandatoryValidation, rsDataItem!WarningValidation)
'                    'Print #nFileNumber, GetDataValue(rsDataItem!DataItemCode)
'                    Print #nFileNumber, CreateFormObject(rsDataItem, rsCRFElement, rsStudyDefinition)
'
'                    Print #nFileNumber, "</div>"
'                    'JL 07/04/00 Print out Status Image.
'                    'JL 16/06/00 Replaced lImageXposition with call to calulateStatusXPosition
'                    'For th control width if Radio need to calulate the max with of the ItemValue (caption for each item) and add it
'                    'to the control width.
'                    lImageXposition = CalulateStatusXPosition(70, CLng(rsCRFElement!X / nIeXSCALE), gsOPTION_BUTTONS, nMaxLengthSoFar)
'
'                    Print #nFileNumber, "<div style=""Left:" & lImageXposition & ";top:" & CLng(rsCRFElement!Y / nIeYSCALE) & """> " & vbNewLine
'
'                    Print #nFileNumber, "    <%= SetStatusImage(getCurrentStatus(lthisId), """ & LCase(rsDataItem!DataItemCode) & """) %>" & vbNewLine
'
'                    Print #nFileNumber, "</div> "
'                     '07/08/00
'                        Print #nFileNumber, printCaption(rsCRFElement, rsStudyDefinition, rsDataItem, _
'                        bDisplayNumbers, bNetscape, nDefaultFontSize, lDefaultFontColour, _
'                        sDefaultFontFamily, nDefaultFontWeight, nDefaultFontStyle)
'
'
'                End If
'            '---------------------------------------------------------------------------'
'            '                               *** PICTURE  ***                            '
'            '---------------------------------------------------------------------------'
'            ElseIf rsCRFElement!ControlType = ControlType.Picture Then
'    '   ATN 4/5/99
'    '   Use virtual directory 'WebRDEDocuments
'    '   JL 16/06/00. Have to get this set up in the meantime change to Documents which actually exists.
'
'                'ic 01/12/00
'                'added condition so that if the element  cannot be found it is omitted
'                If Dir(gsDOCUMENTS_PATH & rsCRFElement!Caption) <> "" Then
'                    Set oPicture = LoadPicture(gsDOCUMENTS_PATH & rsCRFElement!Caption)
'                    ElementWidth = frmGenerateHTML.ScaleY(oPicture.Width, vbHimetric, vbPixels) * nXSCALE
'                    Set oPicture = Nothing
'
'                    If ElementWidth > 7000 And ElementWidth < 5000 Then 'Assume it is a line
'                        Dim maxImageLength As Integer               'limit it to the width of the page.
'                        nMaxImageLength = 611
'                        Print #nFileNumber, "<div id=" & Chr$(34) & "L" & rsCRFElement!CRFElementId & Chr$(34) & " style=" & Chr$(34) & _
'                                           ";left:" & CLng(rsCRFElement!X / nIeXSCALE) & ";top:" & CLng(rsCRFElement!Y / nIeYSCALE) & Chr$(34) & ">"
'
'                        'ic 12/02/01
'                        'removed gsDOCUMENTS_PATH constant as it was being written in the html
'                        Print #nFileNumber, "<img src=" & Chr$(34) & rsCRFElement!Caption & Chr$(34) & " width=" & nMaxImageLength & "  align=""top"" align=""left"">"
'                        Print #nFileNumber, "</div>"
'
'                    Else
'                        Print #nFileNumber, "<div id=" & Chr$(34) & "L" & rsCRFElement!CRFElementId & Chr$(34) & " style=" & Chr$(34) & _
'                                           ";left:" & CLng(rsCRFElement!X / nIeXSCALE) & ";top:" & CLng(rsCRFElement!Y / nIeYSCALE) & Chr$(34) & ">"
'
'                        'ic 12/02/01
'                        'removed gsDOCUMENTS_PATH constant as it was being written in the html
'                        Print #nFileNumber, "<img src=" & Chr$(34) & rsCRFElement!Caption & Chr$(34) & " align=""top"" align=""left"">"
'                        Print #nFileNumber, "</div>"
'                    End If
'
'
'
'
'                    'ic 05/12/00
'                    'copy the file to the web folder
'                    'because writing the asp is currently done for both IE and Netscape, but the
'                    'pictures will be the same for both, im only going to copy them in the IE asp function. if the 2 functions are
'                    'ever combined this code will need to be included.
'                    'it would be preferrable to add more flexible error handling here to check if the file we're about to overwrite
'                    'is currently open (eg in an image editor) but for now we'll assume we can just go ahead
'
'                    'check we havent already copied this file by looking for the filename in our copied images variable
'                    'which is just a comma delimited list of files we add to as we process this function
'                    If InStr(sCpdImages, "," & rsCRFElement!Caption) = 0 Then
'
'                        'if the file exists in the html folder already, delete it
'                        If Dir(gsWEB_HTML_LOCATION & rsCRFElement!Caption) <> "" Then Kill (gsWEB_HTML_LOCATION & rsCRFElement!Caption)
'
'                        'copy the file from source to destination
'                        FileCopy gsDOCUMENTS_PATH & rsCRFElement!Caption, gsWEB_HTML_LOCATION & rsCRFElement!Caption
'
'                        'add to our list of copied images
'                        sCpdImages = sCpdImages & "," & rsCRFElement!Caption
'                    End If
'
'                Else
'                   MsgBox "Image " & gsDOCUMENTS_PATH & rsCRFElement!Caption & " cannot be found!", vbOKOnly
'                End If
'            '---------------------------------------------------------------------------'
'            '                               *** POP-UP LIST  ***                        '
'            '---------------------------------------------------------------------------'
'            ElseIf rsCRFElement!ControlType = ControlType.PopUp Then
'                   'Get values for option boxs etc
'               ' vDataItemId2 = rsCRFElement!DataItemId
'                sSQL = "SELECT * FROM ValueData WHERE" & _
'                    " DataItemId = " & lDataItemId & " AND" & _
'                    " ClinicalTrialId = " & vClinicalTrialId & "AND " & _
'                    " VersionId = " & vVersionId
'
'                Set rsValueData = New ADODB.Recordset
'                rsValueData.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
'
'                'Jl 30/05/00 Code to print caption transfered to new function printCaption.
'                Print #nFileNumber, "<div id=" & Chr$(34) & "L" & rsCRFElement!CRFElementId & Chr$(34) & " style=" & Chr$(34) & _
'                    sPositionAttr & ";left:" & CLng(rsCRFElement!X / nIeXSCALE) & ";top:" & CLng(rsCRFElement!Y / nIeYSCALE) & ";z-index:2" & Chr$(34) & ">"
'
'
'                'Print #nFileNumber, "<select Disabled tabindex=" & Chr$(34) & rsCRFElement!FieldOrder & Chr$(34) & " Name =""cbo" _
'                '& LCase$(rsDataItem!DataItemCode) & """" & " onfocus=""setFocusColour(f.cbo" & LCase(rsDataItem!DataItemCode) & ");""" & _
'                '" onClick=""RefreshForm(f.cbo" & LCase$(rsDataItem!DataItemCode) & ")"" " & ">"
'
'                'ic 21/12/00
'                'added onblur event so that answers are evaluated when tabbing between fields as well as clicking
'                Print #nFileNumber, "<select Disabled tabindex=" & Chr$(34) & rsCRFElement!FieldOrder & Chr$(34) & " Name =""cbo" _
'                & LCase$(rsDataItem!DataItemCode) & """" & " onfocus=""setFocusColour(f.cbo" & LCase(rsDataItem!DataItemCode) & ");""" & _
'                " onClick=""RefreshForm(f.cbo" & LCase$(rsDataItem!DataItemCode) & ")"" onblur=""RefreshForm(f.cbo" & LCase$(rsDataItem!DataItemCode) & ")"" >"
'
'                'Test for at least 1 datavalue
'                If rsValueData.RecordCount > 0 Then
'                    rsValueData.MoveLast
'
'                    Print #nFileNumber, "<% redim v(" & rsValueData.RecordCount & ")"
'                    Print #nFileNumber, "redim c(" & rsValueData.RecordCount & ")"
'
'                    nLengthOfItem = 0
'                    nMaxLengthSoFar = 0
'                    nRCount = 1
'                    Dim nUpper As Integer
'                    Dim nLower As Integer
'                    Dim nCategoryValue As String
'                    Dim bMostlyUpperCase As Boolean
'                    Dim i As Integer
'                    rsValueData.MoveFirst
'                    Do While Not rsValueData.EOF
'                        '06/02/00 JL nRCount replaces (rsValueData.AbsolutePosition + 1) AbsolutePosition not working.
'                        Print #nFileNumber, "c(" & nRCount & ") =""" & rsValueData!ValueCode & """"
'                        Print #nFileNumber, "v(" & nRCount & ") =""" & rsValueData!ItemValue & """"
'                        nRCount = nRCount + 1
'                        nLengthOfItem = Len(rsValueData!ItemValue)
'                        If nLengthOfItem > nMaxLengthSoFar Then
'                            nMaxLengthSoFar = nLengthOfItem
'                            nCategoryValue = rsValueData!ItemValue
'                        End If
'                        rsValueData.MoveNext
'                    Loop
'                End If 'Test for at least 1 datavalue
'
'                bMostlyUpperCase = mostlyUpperCase(nCategoryValue)
'
'                Print #nFileNumber, "SelValues getResponseValue(lthisId) %>" & vbNewLine
'
'                Print #nFileNumber, "</select>"
'
'    '   ATN 3/12/99
'    '   Javascript validation now created at form level at run-time
'
'                Print #nFileNumber, CreateFormObject(rsDataItem, rsCRFElement, rsStudyDefinition)
'
'                Print #nFileNumber, "</div>"
'                'JL 16/06/00 Print out Status Image.
'
'                lImageXposition = CalulateStatusXPosition(nMINIMAL_POPUP_WIDTH, CLng(rsCRFElement!X / nIeXSCALE), gsPOPUP_LIST, nMaxLengthSoFar, bMostlyUpperCase)
'
'                Print #nFileNumber, "<div style=""Left:" & lImageXposition & ";top:" & CLng(rsCRFElement!Y / nIeYSCALE) & """> " & vbNewLine
'
'                Print #nFileNumber, "    <%= SetStatusImage(getCurrentStatus(lthisId), """ & LCase(rsDataItem!DataItemCode) & """) %>" & vbNewLine
'
'                Print #nFileNumber, "</div> "
'
'                 Print #nFileNumber, printCaption(rsCRFElement, rsStudyDefinition, rsDataItem, _
'                        bDisplayNumbers, bNetscape, nDefaultFontSize, lDefaultFontColour, _
'                        sDefaultFontFamily, nDefaultFontWeight, nDefaultFontStyle)
'
'             End If
'
'
'
'            'End If
'            'debug element id
'            rsCRFElement.MoveNext
'
'
'        Loop
'
'        Print #nFileNumber, "</FORM>"
'        Print #nFileNumber, "<Img src=""dummy_img.gif"">"
'        Print #nFileNumber, "</body>"
'        Print #nFileNumber, "</html>"
'        Print #nFileNumber, "<!--#include file=""CloseDataConnection.txt"" -->"
'        'MLM 18/9/00: close statement block for role function check
'        Print #nFileNumber, "<%end if%>"
'        Print #nFileNumber, "<!-- #Include file=""formCyclelist.asp"" -->"
'        Print #nFileNumber, "<!--#include file=""eventHandlers.asp""-->"
'        'MLM 26/10/00 added "Change Data" role function to arguments of GetFormArezzoDetails
'        Print #nFileNumber, "<% response.write session(""P"").GetFormArezzoDetails(session(""sRoleCode""),request.querystring(""trial"") ,request.querystring(""VersionId""), request.querystring(""trialname""), request.querystring(""site"") ,clng(request.querystring(""person"")) , cint(request.querystring(""CRFPageId"")), clng(request.querystring(""crfpagetaskid"")), lCase(request.querystring(""VisitCode"")), lCase(request.querystring(""CRFPageCode"")), session(""F5003"") ) %>"
'
'        Print #nFileNumber, "<script language =""javascript"">"
'        'Print #nFileNumber, "   setInitialCaptions();" & vbNewLine
'        Print #nFileNumber, "function fnInitialisePage()"
'        Print #nFileNumber, "{"
'        Print #nFileNumber, "   ApplyInitialFormat();"
'        Print #nFileNumber, "   DisplayForm();"
'        Print #nFileNumber, "   load_Links();"
'        Print #nFileNumber, "}"
'        Print #nFileNumber, "</script>"
'        Close #nFileNumber
'
'            FormDefinition rsCRFPages!ClinicalTrialId, rsCRFPages!VersionId, rsCRFPages
'            'End If
'            '### debug pageId
'        rsCRFPages.MoveNext
'
'    Loop
'
'    CreateHTMLComponents vClinicalTrialId, vVersionId
'
'    '   ATN 8/12/99
'    '   Added call to create netscape files here, so that its always done.
'    'CreateNetscapefiles vClinicalTrialId, vVersionId
'
'
'    'destroy recordsets
'    rsStudyDefinition.Close
'    Set rsStudyDefinition = Nothing
'    If bNoOfpages = True Then
'        rsCRFPages.Close
'        Set rsCRFPages = Nothing
'        If bNoOfelements = True Then
'            rsCRFElement.Close
'            Set rsCRFElement = Nothing
'            rsDataItem.Close
'            Set rsDataItem = Nothing
'        End If
'    End If
'
'    Exit Sub
'
'ErrHandler:
'  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "CreateHTMLFIles", "modHTML")
'        Case OnErrorAction.Ignore
'            Resume Next
'        Case OnErrorAction.Retry
'            Resume
'        Case OnErrorAction.QuitMACRO
'            Call ExitMACRO
'            Call MACROEnd
'   End Select
'
'End Sub

'------------------------------------------------------------------------------------------------------
Public Function CalulateStatusXPosition(ByVal ControlWidth As Integer, ByVal ControlLeft As Long, _
        ByVal ControlType As Integer, Optional label As Integer, Optional bMostlyUpperCase As Boolean) As Integer
'------------------------------------------------------------------------------------------------------
'Calulates the position of x/left of the status image to be displayed on the right of
'each field.
'ControlWidth is the size property of the textbox.

'ControlLeft represents the X co-ordinate of the control
'JL. Label only used for Radios and pop-up boxes.
'It is the width of the longest text string from the list of data items.
If ControlType = gsTEXT_BOX Then

Select Case ControlWidth
'For case 1 to 6, the increase in length for a textbox for each HTML
                        'ControlWidth = sMaxlength property
    Case 1
        CalulateStatusXPosition = ControlLeft + ControlWidth + 18
    Case 2
        CalulateStatusXPosition = ControlLeft + ControlWidth + 25
    Case 3
        CalulateStatusXPosition = ControlLeft + ControlWidth + 35
    Case 4 To 6
        CalulateStatusXPosition = ControlLeft + ControlWidth + 53
    Case 7 To 9
        CalulateStatusXPosition = ControlLeft + ControlWidth + 75
    Case 10 To 12
        CalulateStatusXPosition = ControlLeft + ControlWidth + 100
    Case 10 To 15
        CalulateStatusXPosition = ControlLeft + ControlWidth + 110
    Case 16 To 20
        CalulateStatusXPosition = ControlLeft + ControlWidth + 120
    Case 21 To 25
        CalulateStatusXPosition = ControlLeft + ControlWidth + 170
    Case 26 To 31
        CalulateStatusXPosition = ControlLeft + ControlWidth + 180
    Case 32 To 40
        CalulateStatusXPosition = ControlLeft + ControlWidth + 275
    Case 41 To 50
        CalulateStatusXPosition = ControlLeft + ControlWidth + 330
    Case 51 To 60
        CalulateStatusXPosition = ControlLeft + ControlWidth + 375
    Case 61 To 80
        CalulateStatusXPosition = ControlLeft + ControlWidth + 320
    Case 81 To 90
        CalulateStatusXPosition = ControlLeft + ControlWidth + 320
    Case 81 To 90
        CalulateStatusXPosition = ControlLeft + ControlWidth + 320
    Case 90 To 100
        CalulateStatusXPosition = ControlLeft + ControlWidth + 320
    Case 130 To 140
        CalulateStatusXPosition = ControlLeft + ControlWidth + 320
    Case Is > 140
        CalulateStatusXPosition = ControlLeft + ControlWidth + 350
    End Select
ElseIf ControlType = gsPOPUP_LIST Then
    'JL 16/06/00 If a POP-up list ControlWidth represents the width of the longest ItemValue (which is a string)
    'in the database for this pop-up
        If bMostlyUpperCase Then
            'Additional numbers added depends on the label/
            Select Case label  '##
            Case 0 To 2    'Up to a string of 4 characters.
                CalulateStatusXPosition = ControlWidth + ControlLeft + 10
            Case 3     'Up to a string of 4 characters.
                CalulateStatusXPosition = ControlWidth + ControlLeft + 10
            Case 4    'Up to a string of 4 characters.
                CalulateStatusXPosition = ControlWidth + ControlLeft + 13
            Case 5 To 6
                CalulateStatusXPosition = ControlWidth + ControlLeft + 55 '25
            Case 7
                CalulateStatusXPosition = ControlWidth + ControlLeft + 55 '25
            Case 8 To 10    'Up to a string of 10 characters.
                CalulateStatusXPosition = ControlWidth + ControlLeft + 65 '25
            Case 11 To 15    'Up to a string of 10 characters.
                CalulateStatusXPosition = ControlWidth + ControlLeft + 85 '55
            Case 16 To 20
                CalulateStatusXPosition = ControlWidth + ControlLeft + 170
            Case 21 To 25
                CalulateStatusXPosition = ControlWidth + ControlLeft + 190
            Case 26 To 30
                CalulateStatusXPosition = ControlWidth + ControlLeft + 230
            Case 31 To 35
                CalulateStatusXPosition = ControlWidth + ControlLeft + 270 '250
            Case 36 To 40
                CalulateStatusXPosition = ControlWidth + ControlLeft + 320
            Case 41 To 50
                CalulateStatusXPosition = ControlWidth + ControlLeft + 210
            Case 51 To 60
                CalulateStatusXPosition = ControlWidth + ControlLeft + 190
            Case Else
                CalulateStatusXPosition = ControlWidth + ControlLeft + 327
            End Select
        Else
            'Additional numbers added depends on the label/
            Select Case label
            Case 0 To 2    'Up to a string of 4 characters.
                CalulateStatusXPosition = ControlWidth + ControlLeft + 10
            Case 3
                CalulateStatusXPosition = ControlWidth + ControlLeft + 10
            Case 4
                CalulateStatusXPosition = ControlWidth + ControlLeft + 13
            Case 5
                CalulateStatusXPosition = ControlWidth + ControlLeft + 25 '25
            Case 6
                CalulateStatusXPosition = ControlWidth + ControlLeft + 30 '25
            Case 7
                CalulateStatusXPosition = ControlWidth + ControlLeft + 39 '25
            Case 8 To 10    'Up to a string of 10 characters.
                CalulateStatusXPosition = ControlWidth + ControlLeft + 65 '25
            Case 11 To 15    'Up to a string of 10 characters.
                CalulateStatusXPosition = ControlWidth + ControlLeft + 85 '55
            Case 16 To 21
                CalulateStatusXPosition = ControlWidth + ControlLeft + 115
            Case 22 To 23
                CalulateStatusXPosition = ControlWidth + ControlLeft + 130
            Case 24
                CalulateStatusXPosition = ControlWidth + ControlLeft + 132
            Case 25 To 26
                CalulateStatusXPosition = ControlWidth + ControlLeft + 130
            Case 27 To 30
                CalulateStatusXPosition = ControlWidth + ControlLeft + 154
            Case 31 To 36
                CalulateStatusXPosition = ControlWidth + ControlLeft + 195 '160
            Case 37 To 38
                CalulateStatusXPosition = ControlWidth + ControlLeft + 230 '160
            Case 39 To 40
                CalulateStatusXPosition = ControlWidth + ControlLeft + 238 '160
            Case 41 To 50
                CalulateStatusXPosition = ControlWidth + ControlLeft + 250
            Case 51 To 56
                CalulateStatusXPosition = ControlWidth + ControlLeft + 300
            Case 57 To 59
                CalulateStatusXPosition = ControlWidth + ControlLeft + 330
            Case Else
                CalulateStatusXPosition = ControlWidth + ControlLeft + 360
            End Select
        End If
ElseIf ControlType = gsOPTION_BUTTONS Then
    'JL 16/06/00 If a Radio control ControlWidth represents the width of the longest ItemValue in the database
    'for this radio.

    Select Case label
    Case 1 To 10    'Up to a string of 10 characters.
        CalulateStatusXPosition = ControlWidth + ControlLeft + label + 10
    Case 11 To 15
        CalulateStatusXPosition = ControlWidth + ControlLeft + label + 60
    Case 16 To 21
        CalulateStatusXPosition = ControlWidth + ControlLeft + label + 77
    Case 22 To 26
        CalulateStatusXPosition = ControlWidth + ControlLeft + label + 92
    Case 27 To 31
        CalulateStatusXPosition = ControlWidth + ControlLeft + label + 110
    Case 32 To 36
        CalulateStatusXPosition = ControlWidth + ControlLeft + label + 120
    Case 37 To 41
        CalulateStatusXPosition = ControlWidth + ControlLeft + label + 130
    Case Else
        CalulateStatusXPosition = ControlWidth + ControlLeft + 340
    End Select
ElseIf ControlType = gsATTACHMENT Then
    CalulateStatusXPosition = ControlLeft + ControlWidth + 250
End If


End Function
'   ATN 3/12/99
'   Javascript validation moved to WebRDE dll

'Public Function JavascriptValidation(ByVal vDataItemCode As String, _
'                                     ByVal vDataItemType As Integer, _
'                                     ByVal vMandatory As Variant, _
'                                     ByVal vWarning As Variant) As String
'
'
'Validation = "<script language =""javascript"" >" & vbNewLine
'JavascriptValidation = JavascriptValidation & LCase(vDataItemCode) & "=""" & "<%= R(""ResponseValue"") %>" & """;" & vbNewLine
'JavascriptValidation = JavascriptValidation & "function V_" & vDataItemCode & "() {" & vbNewLine
'JavascriptValidation = JavascriptValidation & LCase(vDataItemCode) & "=window.event.srcElement.value;" & vbNewLine
'
'Select Case vDataItemType
'    Case DataType.IntegerData
'        JavascriptValidation = JavascriptValidation & " if (CheckInteger(" & LCase(vDataItemCode) & ")) {" & vbNewLine
'    Case DataType.Real
'        JavascriptValidation = JavascriptValidation & " if (CheckReal(" & LCase(vDataItemCode) & ")) {" & vbNewLine
'    Case DataType.Date
'        JavascriptValidation = JavascriptValidation & " if (CheckDate(" & LCase(vDataItemCode) & ")) {" & vbNewLine
'    Case Else
'        JavascriptValidation = JavascriptValidation & " {" & vbNewLine
'End Select
'
''   ATN 5/4/99
''   Replace any new line characters before writing the validation text
'If vMandatory > "" Then
'    JavascriptValidation = JavascriptValidation & "   var MT = """ & LCase$(ReplaceCharacters(vMandatory, vbNewLine, "")) & """;" & vbNewLine
'End If
'If vWarning > "" Then
'    JavascriptValidation = JavascriptValidation & "   var WT = """ & LCase$(ReplaceCharacters(vWarning, vbNewLine, "")) & """;" & vbNewLine
'End If
'
'JavascriptValidation = JavascriptValidation & "<% "
'If vMandatory > "" Then
'    JavascriptValidation = JavascriptValidation & " M = session(""P"").GetJavascriptMandatoryValidation(""" & LCase(vDataItemCode) & """)" & vbNewLine
'ElseIf vWarning > "" Then
'    JavascriptValidation = JavascriptValidation & " M = """"" & vbNewLine
'End If
'If vWarning > "" Then
'    JavascriptValidation = JavascriptValidation & " W = session(""P"").GetJavascriptWarningValidation(""" & LCase(vDataItemCode) & """)" & vbNewLine
'ElseIf vMandatory > "" Then
'    JavascriptValidation = JavascriptValidation & " W = """"" & vbNewLine
'End If
'JavascriptValidation = JavascriptValidation & "%>" & vbNewLine
'
'If vMandatory > "" Or vWarning > "" Then
'    JavascriptValidation = JavascriptValidation & "<!-- #Include file=""Validation.asp"" -->" & vbNewLine
'End If
''jl for testing - remove this.
''JavascriptValidation = JavascriptValidation & "alert(" & LCase(vDataItemCode) & ")"
'JavascriptValidation = JavascriptValidation & " }" & vbNewLine
'JavascriptValidation = JavascriptValidation & "} </script>" & vbNewLine
'
'End Function
'Public Function JavascriptNNvalidation(ByVal vDataItemCode As String, _
'                                     ByVal vDataItemType As Integer, _
'                                     ByVal vMandatory As Variant, _
'                                     ByVal vWarning As Variant, _
'                                     ByVal vControlType As Integer) As String
'
'JavascriptNNvalidation = "<script language =""javascript"" >" & vbNewLine
'If vControlType = 2 Or vControlType = 258 Then
'    JavascriptNNvalidation = JavascriptNNvalidation & LCase(vDataItemCode) & "=""" & "<%= R2(""ResponseValue"") %>" & """;" & vbNewLine
'Else
'    JavascriptNNvalidation = JavascriptNNvalidation & LCase(vDataItemCode) & "=""" & "<%= R(""ResponseValue"") %>" & """;" & vbNewLine
'End If
'JavascriptNNvalidation = JavascriptNNvalidation & "function V_" & vDataItemCode & "() {" & vbNewLine
'If vControlType = 1 Or vControlType = 8 Then
''jl redefined the variable as 'local' with var statement in case vDataItemCode clashes with a global variable.
'    JavascriptNNvalidation = JavascriptNNvalidation & "var " & LCase(vDataItemCode) & "=this.document.forms[0].elements[0].value;" & vbNewLine
'ElseIf vControlType = 4 Then
'    JavascriptNNvalidation = JavascriptNNvalidation & "var " & LCase(vDataItemCode) & "=this.document.forms[0].ComboSelect.options[this.document.forms[0].ComboSelect.selectedIndex].value;" & vbNewLine
'ElseIf vControlType = 2 Or vControlType = 258 Then
'    JavascriptNNvalidation = JavascriptNNvalidation & "var " & LCase(vDataItemCode) & "=getRadioValue(this.document.radioForm.radioCn);" & vbNewLine
'End If
'
'Select Case vDataItemType
'    Case DataType.IntegerData
'        JavascriptNNvalidation = JavascriptNNvalidation & " if (CheckInteger(" & LCase(vDataItemCode) & ")) {" & vbNewLine
'    Case DataType.Real
'        JavascriptNNvalidation = JavascriptNNvalidation & " if (CheckReal(" & LCase(vDataItemCode) & ")) {" & vbNewLine
'    Case DataType.Date
'        JavascriptNNvalidation = JavascriptNNvalidation & " if (CheckDate(" & LCase(vDataItemCode) & ")) {" & vbNewLine
'    Case Else
'        'jl removed bracked
'        JavascriptNNvalidation = JavascriptNNvalidation & " {" & vbNewLine
'End Select
''jl for testing - remove this.
''JavascriptNNvalidation = JavascriptNNvalidation & "new Closure( V_" & vDataItemCode & ", " & vDataItemCode & ")"
''   ATN 5/4/99
''   Replace any new line characters before writing the validation text
'If vMandatory > "" Then
'    JavascriptNNvalidation = JavascriptNNvalidation & "   var MT = """ & LCase$(ReplaceCharacters(vMandatory, vbNewLine, "")) & """;" & vbNewLine
'End If
'If vWarning > "" Then
'    JavascriptNNvalidation = JavascriptNNvalidation & "   var WT = """ & LCase$(ReplaceCharacters(vWarning, vbNewLine, "")) & """;" & vbNewLine
'End If
'
'JavascriptNNvalidation = JavascriptNNvalidation & "<% "
'If vMandatory > "" Then
'    JavascriptNNvalidation = JavascriptNNvalidation & " M = session(""P"").GetJavascriptMandatoryValidation(""" & LCase(vDataItemCode) & """)" & vbNewLine
'ElseIf vWarning > "" Then
'    JavascriptNNvalidation = JavascriptNNvalidation & " M = """"" & vbNewLine
'End If
'If vWarning > "" Then
'    JavascriptNNvalidation = JavascriptNNvalidation & " W = session(""P"").GetJavascriptWarningValidation(""" & LCase(vDataItemCode) & """)" & vbNewLine
'ElseIf vMandatory > "" Then
'    JavascriptNNvalidation = JavascriptNNvalidation & " W = """"" & vbNewLine
'End If
'JavascriptNNvalidation = JavascriptNNvalidation & "%>" & vbNewLine
'
'If vMandatory > "" Or vWarning > "" Then
'    JavascriptNNvalidation = JavascriptNNvalidation & "<!-- #Include file=""Validation.asp"" -->" & vbNewLine
'End If
'''jl for testing - remove this.
''JavascriptNNvalidation = JavascriptNNvalidation & "alert(" & LCase(vDataItemCode) & ")"
''removed bracket
'JavascriptNNvalidation = JavascriptNNvalidation & " }" & vbNewLine
'JavascriptNNvalidation = JavascriptNNvalidation & "} </script>" & vbNewLine
'
'End Function

''-----------------------------------------------------------------------------------------------------'
'Public Function printFontAttributes(ByVal lFontColour As Long, nFontSize As Integer, vFontFace As Variant _
' , lDefaultFontColour As Long, nDefaultFontSize As Integer, sDefaultFontName As String, Optional ByVal vNetscapeFlag As Variant) As String
''-----------------------------------------------------------------------------------------------------'
''-----------------------------------------------------------------------------------------------------'
'Dim sFont As String
'    On Error GoTo ErrHandler
'
'    'COLOR Must go first. Javascript function setCaptionGrey relies on the position of the 1st comma to insert the grey string.
'    If lFontColour <> 0 And lFontColour <> lDefaultFontColour Then
'        sFont = sFont & " color=" & Chr$(34) & GetHexColour(lFontColour) & Chr$(34)
'    ElseIf lFontColour = 0 And lFontColour <> lDefaultFontColour Then
'        sFont = sFont & " color=" & Chr$(34) & GetHexColour(lDefaultFontColour) & Chr$(34)
'    End If
'    'SIZE
'    If vNetscapeFlag Then
'        If nFontSize <> 0 And nFontSize <> nDefaultFontSize Then
'            sFont = sFont & " size=" & Chr$(34) & getNNFontSize(nFontSize) & Chr$(34)
'        End If
'    Else
'        If nFontSize <> 0 And nFontSize <> nDefaultFontSize Then
'            sFont = sFont & " size=" & Chr$(34) & getFontSize(nFontSize) & Chr$(34)
'        ElseIf nFontSize = 0 And nFontSize <> nDefaultFontSize Then
'            sFont = sFont & " size=" & Chr$(34) & getFontSize(nDefaultFontSize) & Chr$(34)
'        ElseIf nFontSize <> 0 Then
'            sFont = sFont & " size=" & Chr$(34) & getFontSize(nFontSize) & Chr$(34)
'        End If
'    End If
'
'    'FAMILY
'    If IsNull(vFontFace) And sDefaultFontName > "" Then
'        sFont = sFont & " face=" & Chr$(34) & sDefaultFontName & Chr$(34)
'    ElseIf Not IsNull(vFontFace) And sDefaultFontName <> vFontFace Then
'        sFont = sFont & " face=" & Chr$(34) & sDefaultFontName & Chr$(34)
'    ElseIf sDefaultFontName > "" Then
'        sFont = sFont & " face=" & Chr$(34) & sDefaultFontName & Chr$(34)
'    End If
'
'    printFontAttributes = sFont & ">"
'
'Exit Function
'ErrHandler:
'  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "printFontAttributes", "modHTML.bas")
'        Case OnErrorAction.Ignore
'            Resume Next
'        Case OnErrorAction.Retry
'            Resume
'        Case OnErrorAction.QuitMACRO
'            Call ExitMACRO
'            Call MACROEnd
'   End Select
'
'End Function

'-------------------------------------------------------------------------------'
Public Sub CreateHTMLComponents(ByVal vClinicalTrialId As Long, _
                                ByVal vVersionId As Integer)
'-------------------------------------------------------------------------------'
'-------------------------------------------------------------------------------'
Dim rsStudyDocument As ADODB.Recordset
Dim sSQL As String
Dim sFileName As String
Dim nFileNumber As Integer
Dim sTempFileName As String
    
    On Error GoTo ErrHandler
    
    nFileNumber = FreeFile
    'SDM 27/01/00 SR2794
    sFileName = gsWEB_HTML_LOCATION & vClinicalTrialId & "_References.htm"
    'sfilename = gsHTML_FORMS_LOCATION & "\" & vClinicalTrialId & "_References" & ".htm"
    sSQL = "SELECT * FROM StudyDocument where ClinicalTrialId = " & vClinicalTrialId & _
                 " AND VersionId = " & vVersionId
    
    Open sFileName For Output As #nFileNumber
    
    Print #nFileNumber, " "
    
    Debug.Print "In CreateHTMLComponents " & sFileName & " opened"
    
    Set rsStudyDocument = New ADODB.Recordset
    rsStudyDocument.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    Do While Not rsStudyDocument.EOF
    
    '   ATN 4/5/99
    '   Modified reference path to use the virtual directory 'WebRDEDocuments'
         sTempFileName = "\WebRDEDocuments\" & rsStudyDocument!DocumentPath
    
        Print #nFileNumber, " <a href=" & Chr$(34) & sTempFileName & Chr$(34) & ">" & _
                rsStudyDocument!DocumentPath & _
                "</a>   "
        
        rsStudyDocument.MoveNext
    Loop
    Close #nFileNumber
    
    rsStudyDocument.Close
    Set rsStudyDocument = Nothing

Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "CreateHTMLComponents", "modHTML.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

'------------------------------------------------------------------------------------------------------
Private Function FieldAttributes(ByVal fieldno, Caption, displayNumbers, UnitOfMeasurement) As String
'------------------------------------------------------------------------------------------------------
'JL 04/08/00. Replacing with style string.
'Removed bold and italic args as they are now dealt with in printNetscapeCaption.
'------------------------------------------------------------------------------------------------------
Dim units As String
Dim justText As String
Dim fieldText As String

    On Error GoTo ErrHandler

If IsNull(Caption) Then
    Exit Function
End If
If displayNumbers = True Then
    fieldText = fieldno & ". "
Else
    fieldText = ""
End If

If fieldno <> 0 Then
    If Trim(UnitOfMeasurement) > gsEMPTY_STRING Then
        units = " (" & UnitOfMeasurement & ")"
    Else
        units = ""
    End If
End If
If fieldno = 0 Then
    Dim bComment As Boolean
    bComment = True
End If
'If bold = 1 And italic = 1 Then
    If bComment = True Then ' Then this is comment text so fieldno will not be displayed.
        'FieldAttributes = "<em><strong>" & ReplaceCharacters(Caption, vbNewLine, "<BR>") & "</strong></em>"
        FieldAttributes = Replace(Caption, vbNewLine, "<BR>")
    ElseIf fieldno > 0 And Caption > "" Then
        'FieldAttributes = "<em><strong>" & fieldText & ReplaceCharacters(Caption, vbNewLine, "<BR>") & units & "</strong></em>"
        FieldAttributes = fieldText & Replace(Caption, vbNewLine, "<BR>") & units
    End If
'ElseIf bold = 0 And italic = 1 Then
'    If bComment = True Then
'        FieldAttributes = "<em>" & ReplaceCharacters(Caption, vbNewLine, "<BR>") & units & "</em>"
'    ElseIf fieldno > 0 And (Caption = "" Or Caption = Null) Then
'        'Do nothing
'    ElseIf fieldno > 0 And Caption > "" Then
'        FieldAttributes = "<em>" & fieldText & ReplaceCharacters(Caption, vbNewLine, "<BR>") & units & "</em>"
'    End If
'ElseIf bold = 1 And italic = 0 Then
'    If bComment = True Then
'        'FieldAttributes = "<strong>" & ReplaceCharacters(Caption, vbNewLine, "<BR>") & "</strong>"
'        FieldAttributes = ReplaceCharacters(Caption, vbNewLine, "<BR>")
'    ElseIf fieldno > 0 And (Caption = "" Or Caption = Null) Then
'        'Do nothing
'    ElseIf fieldno > 0 And Caption > "" Then
'        'FieldAttributes = "<strong>" & fieldText & ReplaceCharacters(Caption, vbNewLine, "<BR>") & units & "</strong>"
'        FieldAttributes = fieldText & ReplaceCharacters(Caption, vbNewLine, "<BR>") & units
'    End If
'ElseIf bold = 0 And italic = 0 Then
'    If bComment = True Then
'        'FieldAttributes = fieldText & ReplaceCharacters(Caption, vbNewLine, "<BR>") & units  '07/08/00. JL Comment out for compile
'    ElseIf fieldno > 0 And (Caption = "" Or Caption = Null) Then
'        'do nothing
'    ElseIf fieldno > 0 And Caption > "" Then
'            'FieldAttributes = fieldText & ReplaceCharacters(Caption, vbNewLine, "<BR>") & units     '07/08/00. JL Comment out for compile
'
'    End If
'End If

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "FieldAttributes", "modHTML.bas")
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
Private Function FieldAttributes_Explorer(ByVal bold, italic, fieldno, Caption, displayNumbers, Optional ByVal UnitOfMeasurement As Variant) As String
'---------------------------------------------------------------------
'JL 04/08/00. Replacing with style string.
'Removed bold and italic args as they are now dealt with in printNetscapeCaption.
'------------------------------------------------------------------------------------------------------
Dim units As String
Dim justText As String
Dim fieldText As String

    On Error GoTo ErrHandler

If IsNull(Caption) Then
    Exit Function
End If
If displayNumbers = True Then
    fieldText = fieldno & ". "
Else
    fieldText = ""
End If

If fieldno <> 0 Then
    'ic 23/08/01 SR4415 fix
    'check if optional parameter is missing
    If IsMissing(UnitOfMeasurement) Then
        units = ""
    Else
        If Trim(UnitOfMeasurement) > gsEMPTY_STRING Then
            units = " (" & UnitOfMeasurement & ")"
        Else
            units = ""
        End If
    End If
End If
If bold = 1 And italic = 1 Then
    If fieldno = 0 Then ' Then this is comment text so fieldno will not be displayed.
        FieldAttributes_Explorer = "<em><strong>" & Replace(Caption, vbNewLine, "<BR>") & "</strong></em>"

    ElseIf fieldno > 0 And Caption > "" Then
        FieldAttributes_Explorer = "<em><strong>" & fieldText & Replace(Caption, vbNewLine, "<BR>") & units & "</strong></em>"
    End If
ElseIf bold = 0 And italic = 1 Then
    If fieldno = 0 Then
        FieldAttributes_Explorer = "<em>" & Replace(Caption, vbNewLine, "<BR>") & units & "</em>"
    ElseIf fieldno > 0 And (Caption = "" Or Caption = Null) Then
        'Do nothing
    ElseIf fieldno > 0 And Caption > "" Then
        FieldAttributes_Explorer = "<em>" & fieldText & Replace(Caption, vbNewLine, "<BR>") & units & "</em>"
    End If
ElseIf bold = 1 And italic = 0 Then
    If fieldno = 0 Then
        FieldAttributes_Explorer = "<strong>" & Replace(Caption, vbNewLine, "<BR>") & "</strong>"
    ElseIf fieldno > 0 And (Caption = "" Or Caption = Null) Then
        'Do nothing
    ElseIf fieldno > 0 And Caption > "" Then
        FieldAttributes_Explorer = "<strong>" & fieldText & Replace(Caption, vbNewLine, "<BR>") & units & "</strong>"
    End If
ElseIf bold = 0 And italic = 0 Then
    If fieldno = 0 Then
        FieldAttributes_Explorer = fieldText & Replace(Caption, vbNewLine, "<BR>") & units  '07/08/00. JL Comment out for compile
    ElseIf fieldno > 0 And (Caption = "" Or Caption = Null) Then
        'do nothing
    ElseIf fieldno > 0 And Caption > "" Then
            FieldAttributes_Explorer = fieldText & Replace(Caption, vbNewLine, "<BR>") & units     '07/08/00. JL Comment out for compile

    End If
End If

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "FieldAttributes_Explorer", "modHTML.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Function
'-------------------------------------------------------------------------------'
Private Function ReplaceAsciiCRwithHTML(ByVal Caption) As String
'-------------------------------------------------------------------------------'
'-------------------------------------------------------------------------------'
Dim msTempString As String
Dim msSingleChar As String
Dim Pos As Integer
    
    On Error GoTo ErrHandler

msTempString = ""
For Pos = 1 To Len(Caption)
    msSingleChar = Mid(Caption, Pos, 1)
    If Asc(msSingleChar) = 13 Then
        msSingleChar = "<BR>"
    End If
    msTempString = msTempString & msSingleChar
Next Pos
ReplaceAsciiCRwithHTML = msTempString

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ReplaceAsciiCRwithHTML", "modHTML.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Function

''-------------------------------------------------------------------------------'
'Private Function GetHexColour(lColour As Long) As String
''-------------------------------------------------------------------------------'
''-------------------------------------------------------------------------------'
'    Dim r As String
'    Dim g As String
'    Dim b As String
'    Dim rgb As String
'
'    On Error GoTo ErrHandler
'
'
'    rgb = Hex(2 ^ 24 + lColour)
'
'    'JL 11/08/00. Commented this out as it seems to be returning the wrong colour. No need to switch the positions of rgb anymore
'    b = Mid(rgb, 2, 2)
'    g = Mid(rgb, 4, 2)
'    r = Mid(rgb, 6, 2)
'
'    GetHexColour = r$ + g$ + b$
'
'Exit Function
'ErrHandler:
'  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "GetHexColour", "modHTML.bas")
'        Case OnErrorAction.Ignore
'            Resume Next
'        Case OnErrorAction.Retry
'            Resume
'        Case OnErrorAction.QuitMACRO
'            Call ExitMACRO
'            Call MACROEnd
'   End Select
'End Function

'-------------------------------------------------------------------------------'
Private Function RtnFontSize(ByVal nFontSize As Integer) As Integer
'-------------------------------------------------------------------------------'
'-------------------------------------------------------------------------------'
Dim nRtn As Integer

    If nFontSize < 8 And nFontSize > 0 Then
        nRtn = 1
    ElseIf nFontSize < 12 And nFontSize >= 8 Then
        nRtn = 2
    ElseIf nFontSize < 14 And nFontSize >= 12 Then
        nRtn = 4
    ElseIf nFontSize < 16 And nFontSize >= 14 Then
        nRtn = 5
    ElseIf nFontSize < 18 And nFontSize >= 16 Then
        nRtn = 5
    ElseIf nFontSize < 24 And nFontSize >= 18 Then
        nRtn = 6
    ElseIf nFontSize < 36 And nFontSize >= 24 Then
        nRtn = 7
    ElseIf nFontSize >= 36 Then
        nRtn = 8
    Else
        nRtn = 0
    End If

    RtnFontSize = nRtn
End Function

'-----------------------------------------------------------------------------------------------'
Private Function getNNFontSize(F_size) As String
'-----------------------------------------------------------------------------------------------'
'
'-----------------------------------------------------------------------------------------------'
'JL 04/08/00. Font size based on the units used within <FONT> tags.
'Replacing font tag and putting everything within the style= string, but this uses different units to that in font tag.
'JL 05/08/00. Changed this function to return New Font sizes, based on the mm unit.
'JL 10/08/00. Changed the return value from Integer to String, so the 'units' in mm can also be returned.

    On Error GoTo ErrHandler
        'mm unit
        If F_size < 8 And F_size >= 0 Then
            getNNFontSize = "3mm"                 '##
        ElseIf F_size < 10 And F_size >= 8 Then
            getNNFontSize = "4mm"
        ElseIf F_size < 12 And F_size >= 10 Then  '##
            getNNFontSize = "4mm"
        ElseIf F_size < 14 And F_size >= 12 Then
            getNNFontSize = "5mm"
        ElseIf F_size < 16 And F_size >= 14 Then
            getNNFontSize = "6mm"
        ElseIf F_size < 18 And F_size >= 15 Then   '##
            getNNFontSize = "6mm"
        ElseIf F_size < 24 And F_size >= 18 Then
            getNNFontSize = "5mm"
        ElseIf F_size < 36 And F_size >= 24 Then
            getNNFontSize = "6mm"
        ElseIf F_size >= 36 Then
            getNNFontSize = "7mm"
        Else
            getNNFontSize = "7mm"
        End If

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "getNNFontSize", "modHTML.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function

'-------------------------------------------------------------------------------'
Public Sub FormDefinition(ByVal mnClinicalTrialId As Long, _
                            ByVal mnVersionId As Integer, _
                            ByVal rsCRFPages As ADODB.Recordset)
'-------------------------------------------------------------------------------'
' REVISIONS
' DPH 30/10/2001 - Completed Form Definition (Derivation, validation)
' ASH 11/01/2002 - Re-Wrote sql for use on oracle
'-------------------------------------------------------------------------------'
Dim vFilename As String
Dim vFileNumber As Integer
vFileNumber = FreeFile
Dim vHeader As String
Dim rsCRFElement As ADODB.Recordset
Dim sSQL As String
Dim qHTML As String
Dim rsDataItem As ADODB.Recordset
Dim vFooter As String
Dim vDataItemName As String
Dim vDataItemCode As String
Dim vMandatoryValidation As String
Dim vWarningValidation As String
Dim sDerivation As String
Dim rsDataItemValidation As ADODB.Recordset

    On Error GoTo ErrHandler

'SDM 27/01/00 SR2794
vFilename = gsWEB_HTML_LOCATION & mnClinicalTrialId & "_" & rsCRFPages!CRFPageId & "d" & ".htm"
'vFilename = gsHTML_FORMS_LOCATION & "\" & mnClinicalTrialId & "_" & rsCRFPages!CRFPageId & "d" & ".htm"

Open vFilename For Output As #vFileNumber

Debug.Print "Opening " & vFilename & " in Form definition"

'Write header information for this form
vHeader = "<html>" & vbNewLine & "<head>" & vbNewLine & "<title>" & _
          "Form : " & rsCRFPages!CRFTitle & " definition" & "</title>" & vbNewLine & _
          "</head>" & vbNewLine & "<body bgcolor=" & """" & "#FFFFFF" & """" & ">" & vbNewLine & _
          "<h1>" & "Form : " & rsCRFPages!CRFTitle & " definition" & "</h1>" & vbNewLine & _
          "<div align=" & """" & "center" & """" & "><center>" & vbNewLine & _
          "<table border=" & """" & "1" & """" & " cellpadding=" & """" & "1" & """" & _
          " cellspacing=" & """" & "0" & """" & " width=" & """" & "100%" & """" & ">"
          
Debug.Print "printing header for " & rsCRFPages!CRFTitle
Print #vFileNumber, vHeader

sSQL = "SELECT * FROM CRFElement where " & _
    "CRFPageId = " & rsCRFPages!CRFPageId & _
    " AND " & "ClinicalTrialId=" & mnClinicalTrialId & _
    " AND VersionId=" & mnVersionId & " AND " & _
    "DataItemId > 0" & " ORDER BY FieldOrder"
'For each CRFElement on this form, write out HTML
Set rsCRFElement = New ADODB.Recordset
rsCRFElement.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText

Do While Not rsCRFElement.EOF
    
    'For this CRFElement get the DataItems
    sSQL = "SELECT * FROM DataItem where DataItemId=" & _
        rsCRFElement!DataItemId & " AND " & "ClinicalTrialId=" & _
        mnClinicalTrialId & " AND VersionId=" & mnVersionId
        
    Set rsDataItem = New ADODB.Recordset
    rsDataItem.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
                                        
                                        
    vDataItemName = rsDataItem!DataItemName
    vDataItemCode = rsDataItem!DataItemCode
    ' DPH 30/10/2001
    sDerivation = RemoveNull(rsDataItem!Derivation)
    'rsDataItem.Close
    
    If goUser.Database.DatabaseType = MACRODatabaseType.Oracle80 Then
        ' ATO 2/1/2002 Added SQL to take care of Oracle syntax for left joins
        ' Build List Bug report no.10 (v2.2.5)
        sSQL = "SELECT DataItemValidation.*, ValidationType.ValidationActionId " _
                & " FROM DataItemValidation,ValidationType " _
                & " WHERE ValidationType.ValidationTypeId(+) = DataItemValidation.ValidationTypeId" _
                & " AND DataItemId=" & rsCRFElement!DataItemId _
                & " AND ClinicalTrialId=" & mnClinicalTrialId _
                & " And VersionId =" & mnVersionId _
                & " ORDER BY ValidationType.ValidationActionId"
               
     Else
        ' DPH 30/10/2001 - MandatoryValidation & WarningValidation info
        sSQL = "SELECT DataItemValidation.*, ValidationType.ValidationActionId FROM DataItemValidation " & _
            " LEFT JOIN ValidationType ON DataItemValidation.ValidationTypeId = ValidationType.ValidationTypeId " & _
            "WHERE DataItemId=" & rsCRFElement!DataItemId & " AND " & "ClinicalTrialId=" & _
            mnClinicalTrialId & " AND VersionId=" & mnVersionId & _
            " ORDER BY ValidationType.ValidationActionId"
   End If
    
    Set rsDataItemValidation = New ADODB.Recordset
    rsDataItemValidation.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    vMandatoryValidation = ""
    vWarningValidation = ""
    While Not rsDataItemValidation.EOF
        Select Case rsDataItemValidation!ValidationActionId
            Case 0
                If vMandatoryValidation <> "" Then
                    vMandatoryValidation = vMandatoryValidation & "<BR>"
                End If
                vMandatoryValidation = vMandatoryValidation & RemoveNull(rsDataItemValidation!DataItemValidation)
            Case 1
                If vWarningValidation <> "" Then
                    vWarningValidation = vWarningValidation & "<BR>"
                End If
                vWarningValidation = vWarningValidation & RemoveNull(rsDataItemValidation!DataItemValidation)
            Case Else
        End Select
        rsDataItemValidation.MoveNext
    Wend
    rsDataItemValidation.Close
    Set rsDataItemValidation = Nothing
    
    If vMandatoryValidation = "" Then
        vMandatoryValidation = "&nbsp;"
    End If
    If vWarningValidation = "" Then
        vWarningValidation = "&nbsp;"
    End If
    
    'Debug.Print "getting dataItems for CRFElement =" & rsCRFElement!Caption
    'Debug.Print "dataItem = " & rsDataItem!DataItemName
    
    qHTML = "<tr>" & vbNewLine & _
            "<td align=" & """" & "center" & """" & " valign=" & "top" & _
            " width=" & """" & "5%" & """" & ">" & rsCRFElement!FieldOrder & "</td>" & vbNewLine & _
            "<td valign=" & """" & "top" & """" & " width=" & """" & "25%" & """" & "><strong>" & _
            vDataItemCode & "</strong>" & "(" & vDataItemName & ")" & "</td>" & vbNewLine & _
            "<td width=" & """" & "10%" & """" & " valign=top>Only accept if:" & vbNewLine & _
            "<p>Warn if:</p>" & vbNewLine & "<p>Derive as:</p>" & vbNewLine & _
            "<p>Only Collect if:</p></td>" & vbNewLine & _
            "<td width=" & """" & "25%" & """" & " valign=top>" & vMandatoryValidation & vbNewLine & _
            "<p>" & vWarningValidation & "</p>" & vbNewLine & _
            "<p>" & sDerivation & "</p>" & vbNewLine & _
            "<p>" & rsCRFElement!SkipCondition & "</p></td>" & vbNewLine & "</tr>"
            
    Print #vFileNumber, qHTML
    rsDataItem.Close
    Set rsDataItem = Nothing
 rsCRFElement.MoveNext
Loop

vFooter = "</table>" & vbNewLine & "</center>" & vbNewLine & "</div>" & vbNewLine & _
          "</body>" & vbNewLine & "</html>"
          
Print #vFileNumber, vFooter
Close #vFileNumber

rsCRFElement.Close
Set rsCRFElement = Nothing


Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "FormDefinition", "modHTML.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

''-------------------------------------------------------------------------------------------'
'Public Sub CreateNetscapefiles(ByVal lClinicalTrialId As Long, _
'                           ByVal nVersionId As Integer)
''-------------------------------------------------------------------------------------------'
''Creates HTML eforms displayed in the browser.
''-------------------------------------------------------------------------------------------'
''MLM 18/9/00    Check user has logged in and has correct role function.
''MLM 26/10/00   Removed []s from SQL.
''               Removed debugging border=1.
''MLM  9/11/00   Check that user has permission to view site
''ic  10/11/00   added recordset declaration to asp header
''               added  CRFElement to SELECT statement
''ic  16/11/00   changed form name to myform
''MLM 21/11/00   Also check subject locking
''JL     21/11/00    Added standard prefixes in the creation of form elements.
''ic  28/11/00   added condition for drawing image spacers
''ic  29/11/00   commented out javascript variable declaration
''               added 'round' function call to height and width measurements
''               added condition for reducing image spacer size
''ic  30/11/00   added extra condition to rsCRFPage SQL query to eliminate empty captions
''               added condition to rsCRFPage SQL query to eliminate empty captions
''ic  01/12/00   added condition so that if the element width or height is 0 (it cannot be found) it is omitted
''JL  18/12/00    Added an extra argument current user's role code to getFormArezzoDetails for
''                       disabling a question which needs authorisation.
''ic  22/12/00   added nElementHeight, nElementWidth to CellContents function call so that images can be dynamically sized
''               added declarations for X Y calculations
''               added code to change negative x y coordinates to 0
''               added condition to eliminate hidden captions from recordset
''MLM 05/03/01   Converted nCellHeight and nTmpY into Longs to cope with long forms and long empty spaces.
''-------------------------------------------------------------------------------------------'
''Exit Sub
'
'    Dim rsStudyDefinition As ADODB.Recordset
'    Dim rsCRFPage As ADODB.Recordset
'    Dim rsDataItem As ADODB.Recordset
'    Dim rsCRFElement As ADODB.Recordset
'
'    Dim lRow() As Long  'arrays to store the heights and widths of all the table cells
'    Dim lCol() As Long
'    Dim nCellUsed() As Integer 'used to track rows and cols spanned by cells
'    Dim nRow As Integer 'loop counters
'    Dim nCol As Integer
'    Dim nElement As Integer
'    Dim nRowCount As Integer
'    Dim nColCount As Integer
'    Dim nElementCount As Integer
'    Dim nElementHeight As Integer
'    Dim nElementWidth As Integer
'    Dim lCellHeight As Long 'MLM 5/3/01
'    Dim nCellWidth As Integer
'    Dim nRowSpan As Integer
'    Dim nColSpan As Integer
'
'    Dim nFileNumber As Integer
'    Dim sSQL As String
'    Dim sFileName As String
'    Dim nTempIndex As Integer
'    Dim lTemp As Long
'    Dim nCount As Integer
'
'    'ic 22/12/00
'    'added declarations for X Y calculations
'    Dim nTmpX As Integer
'    Dim lTmpY As Long 'MLM 5/3/01
'
'    On Error GoTo ErrHandler
'
'    nFileNumber = FreeFile
'
'    'retrieve the Study Definition record, which includes the default font properties and form colour
'    sSQL = "SELECT * FROM StudyDefinition" _
'        & " WHERE ClinicalTrialId = " & lClinicalTrialId _
'        & " AND VersionId = " & nVersionId
'    Set rsStudyDefinition = New ADODB.Recordset
'    rsStudyDefinition.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'    'get forms for this trial.
'    sSQL = "SELECT * FROM CRFPage " _
'        & " WHERE ClinicalTrialId = " & lClinicalTrialId _
'        & " AND VersionId = " & nVersionId
'    Set rsCRFPage = New ADODB.Recordset
'    rsCRFPage.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
''    If rsCRFPages.EOF Then
''        bNoOfpages = False
''    End If
'
'    Do While Not rsCRFPage.EOF
'        ' Start individual form -------------------------------
'
'        'open file into which the form will be written
'        sFileName = gsWEB_HTML_LOCATION & LTrim(str(rsCRFPage!ClinicalTrialId)) & "_" & _
'            LTrim(str(rsCRFPage!CRFPageId)) & "_" & goUser.Database.NameOfDatabase & "_N.asp"
'        Open sFileName For Output As #nFileNumber
'
'        'write ASP header
'        Print #nFileNumber, "<%@ LANGUAGE = VBScript %>" 'specify server-side language for ASP
'        Print #nFileNumber, "<%"
'        Print #nFileNumber, "Option Explicit"
'        Print #nFileNumber, "Dim cnnMACRO"
'        'ic 10/11/00
'        'added recordset declaration
'        Print #nFileNumber, "Dim rsSiteUser"
'        'MLM 21/11/00
'        Print #nFileNumber, "dim sErrMsg"
'
'
'        Print #nFileNumber, "Response.Expires = 0"
'        Print #nFileNumber, "%>" & vbNewLine
'
'        'check user's permissions
'        Print #nFileNumber, "<!-- #include file=""CheckLogin.asp"" -->"
'        'MLM 10/11/00 added OpenDataConnection so that site permission can be checked
'        Print #nFileNumber, "<!-- #Include file=""OpenDataConnection.txt"" -->"
'
'         'retrieve site permission
'        Print #nFileNumber, "<%set rsSiteUser = cnnMACRO.Execute(""SELECT COUNT(*) AS Permission FROM SiteUser" & _
'            " WHERE Site = '"" & Request.QueryString(""site"") & ""'" & _
'            " AND UserCode = '"" & session(""UserName"") & ""'"")"
'        Print #nFileNumber, "sErrMsg = session(""P"").OpenSubject(Request.QueryString(""trial""), Request.QueryString(""site""), Request.QueryString(""person""))"
'
'        Print #nFileNumber, "if not session(""F5002"") then%>"
'        Print #nFileNumber, "    <html><body bgcolor=white><table width=100% height=100%><tr><td align=center>"
'        Print #nFileNumber, "        You do not have permission to view subject data."
'        Print #nFileNumber, "    </td></tr></table></body></html>"
'
'        'MLM 9/11/00: check user can view this site
'        Print #nFileNumber, "<%elseif rsSiteUser(""Permission"") = 0 then%>"
'        Print #nFileNumber, "    <html><body bgcolor=white><table width=100% height=100%><tr><td align=center>"
'        Print #nFileNumber, "        You do not have permission to access <%=Request.QueryString(""site"")%>."
'        Print #nFileNumber, "    </td></tr></table></body></html>"
'
'        'MLM 21/11/00 also check file locking
'        Print #nFileNumber, "<%elseif sErrMsg > """" then%>"
'        Print #nFileNumber, "    <html><body bgcolor=white><table width=100% height=100%><tr><td align=center>"
'        Print #nFileNumber, "        <%=sErrMsg%>"
'        Print #nFileNumber, "    </td></tr></table></body></html>"
'        Print #nFileNumber, "<%else%>" & vbNewLine
'
'        'write HTML header
'        Print #nFileNumber, "<html>"
'        Print #nFileNumber, "<head><title>Form : " & rsCRFPage!CRFTitle & "</title></head>" & vbNewLine
'
''        Print #nFileNumber, "<!-- #include file=""CheckLogin.asp"" -->"
'        Print #nFileNumber, "<!-- #include file=""format.htm"" -->"
'        Print #nFileNumber, "<!-- #include file=""arezzo.htm"" -->"
'        Print #nFileNumber, "<!-- #Include file=""HandleNNDocument.asp"" -->"
'
'        'retrieve the codes of questions on this form and create a javascript variable declaration for each
'        sSQL = "SELECT DataItemCode FROM DataItem,CRFElement where " _
'            & " CRFPageId =" & rsCRFPage!CRFPageId _
'            & " AND DataItem.DataItemId = CRFElement.DataItemId " _
'            & " AND DataItem.ClinicalTrialId = CRFElement.ClinicalTrialId " _
'            & " AND DataItem.VersionId = CRFElement.VersionId " _
'            & " AND DataItem.ClinicalTrialId =" & lClinicalTrialId _
'            & " AND DataItem.VersionId=" & nVersionId
'        Set rsDataItem = New ADODB.Recordset
'        rsDataItem.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'
'        'ic 29/11/00
'        'commented out this javascript variable declaration as i dont think its needed
'        'Print #nFileNumber, "<script language =""javascript"">"
'        'Do While Not rsDataItem.EOF
'        '    Print #nFileNumber, "    var " & LCase(rsDataItem!DataItemCode) & ";"
'        '    rsDataItem.MoveNext
'        'Loop
'        'Print #nFileNumber, "</script>" & vbNewLine
'
'
'
'        'start the HTML body
'        If rsCRFPage!BackgroundColour = 0 Then
'            Print #nFileNumber, "<body bgcolor=#" & GetHexColour(rsStudyDefinition!DefaultCRFPageColour) & " onUnload=""disableLinks();"" onLoad=""fnInitialisePage();"">"
'        Else
'            Print #nFileNumber, "<body bgcolor=#" & GetHexColour(rsCRFPage!BackgroundColour) & " onUnload=""disableLinks();"" onLoad=""fnInitialisePage();"">"
'        End If
'
'
'
'        'ic 16/11/00
'        'changed form name from 'frmController' to 'myform'
'        Print #nFileNumber, "<form name=myform action=Saveform.asp method=post>"
'
'
'
'        'MLM 10/11/00 DataValues included here because it refers to form object
'        Print #nFileNumber, "<!-- #Include file=""DataValuesNN.asp"" -->"
'
'        Print #nFileNumber, "<table border=0 cellpadding=0 cellspacing=0>"
'
'        'retrieve all the form elements
'        'Captions are converted to Comments within the SQL
'        'MLM 26/10/00 Removed []s from SQL because they were giving "3604: Incorrect syntax near [" for SQL Server and Oracle errors as well.
'        'sSQL = "SELECT Y, X, DataItemId, ControlType, FieldOrder, FontColour, Caption, FontName, FontBold, FontItalic, FontSize, Hidden, Mandatory" & _
'        '    " FROM CRFElement" & _
'        '    " WHERE ClinicalTrialId = " & lClinicalTrialId & _
'        '    " AND VersionId = " & nVersionId & _
'        '    " AND CRFPageId = " & rsCRFPage!CRFPageId & _
'        '    " UNION SELECT CaptionY AS Y, CaptionX AS X, DataItemId, " & ControlType.Comment & " AS ControlType, FieldOrder, FontColour, Caption, FontName, FontBold, FontItalic, FontSize, Hidden, Mandatory" & _
'        '    " FROM CRFElement" & _
'        '    " WHERE ClinicalTrialId = " & lClinicalTrialId & _
'        '    " AND VersionId = " & nVersionId & _
'        '    " AND CRFPageId = " & rsCRFPage!CRFPageId & _
'        '    " AND DataItemId > 0"
'
'
'        'ic 10/11/00
'        'added CRFElementId to SELECT fields as it is needed in function CreateFormObject
'        'ic 30/11/00
'        'added condition to eliminate empty comments from the recordset
'        'ic 22/12/00
'        'added condition to eliminate hidden captions from the recordset
'        sSQL = "SELECT  Y, X, CRFElementId, DataItemId, ControlType, FieldOrder, FontColour, Caption, FontName, FontBold, FontItalic, FontSize, Hidden, Mandatory" & _
'            " FROM CRFElement" & _
'            " WHERE ClinicalTrialId = " & lClinicalTrialId & _
'            " AND VersionId = " & nVersionId & _
'            " AND CRFPageId = " & rsCRFPage!CRFPageId & _
'            " AND (ControlType <> " & ControlType.Comment & " OR (ControlType = " & ControlType.Comment & " AND Caption <> ''))" & _
'            " UNION SELECT  CaptionY AS Y, CaptionX AS X, CRFElementId, DataItemId, " & ControlType.Comment & " AS ControlType, FieldOrder, FontColour, Caption, FontName, FontBold, FontItalic, FontSize, Hidden, Mandatory" & _
'            " FROM CRFElement" & _
'            " WHERE ClinicalTrialId = " & lClinicalTrialId & _
'            " AND VersionId = " & nVersionId & _
'            " AND CRFPageId = " & rsCRFPage!CRFPageId & _
'            " AND DataItemId > 0" & _
'            " AND Hidden = 0"
'
'
'        'ic 30/11/00
'        'added condition so that if an elements caption is empty and we arent displaying numbers, it is eliminated from the recordset
'        If rsCRFPage!displayNumbers = 0 Then sSQL = sSQL & " AND Caption <> ''"
'
'
'        Set rsCRFElement = New ADODB.Recordset
'        rsCRFElement.Open sSQL, MacroADODBConnection, adOpenStatic, adLockReadOnly, adCmdText
'        nElementCount = rsCRFElement.RecordCount
'
'        If nElementCount > 0 Then
'            'read the coordinates of the elements into arrays
'            ReDim lRow(nElementCount)
'            ReDim lCol(nElementCount)
'
'            'ic 22/12/00
'            'changed 0 to -1 to ensure that there is always a row/column border containing image spacers
'            lRow(0) = -1 '0
'            lCol(0) = -1 '0
'
'            For nElement = 1 To nElementCount
'                lRow(nElement) = rsCRFElement!Y
'
'                'ic 22/12/00
'                'added condition so that if y coordinate is negative, it is 0
'                If lRow(nElement) < 0 Then lRow(nElement) = 0
'
'                If rsCRFElement!ControlType = ControlType.Line Then
'                    lCol(nElement) = 0
'                Else
'                    lCol(nElement) = rsCRFElement!X
'
'                    'ic 22/12/00
'                    'added condition so that if x coordinate is negative, it is 0
'                    If lCol(nElement) < 0 Then lCol(nElement) = 0
'
'                End If
'                rsCRFElement.MoveNext
'            Next nElement
'            'the Y-values will already be ordered, but repeated values need to be removed
'            nRowCount = 0
'            For nRow = 1 To nElementCount
'                If lRow(nRowCount) < lRow(nRow) Then
'                    nRowCount = nRowCount + 1
'                    lRow(nRowCount) = lRow(nRow)
'                End If
'            Next nRow
'            'X-values need to be sorted into ascending order and have repeat values removed
'            nColCount = 0
'            For nCol = 1 To nElementCount
'                If nCol < nElementCount Then
'                    nTempIndex = nCol
'                    lTemp = lCol(nCol)
'                    For nCount = nCol + 1 To nElementCount
'                        If lCol(nCount) < lCol(nTempIndex) Then
'                            nTempIndex = nCount
'                            lTemp = lCol(nCount)
'                        End If
'                    Next nCount
'                    If nTempIndex > nCol Then
'                        lCol(nTempIndex) = lCol(nCol)
'                        lCol(nCol) = lTemp
'                    End If
'                End If
'                If lCol(nColCount) < lCol(nCol) Then
'                    nColCount = nColCount + 1
'                    lCol(nColCount) = lCol(nCol)
'                End If
'            Next nCol
'            'now nRowCount and nColCount will be the number of rows and cols needed in the table
'            'any data in lRow() and lCol() above these indeces is junk
'
'            ReDim nCellUsed(nColCount)
'            rsCRFElement.MoveFirst
'
'            'loop through all cells in the table, writing the correct contents in each
'            For nRow = 0 To nRowCount
'                Print #nFileNumber, "<tr>"
'                For nCol = 0 To nColCount
'                    If nCellUsed(nCol) = 0 Then
'                        'don't bother completing the last row if all elements have been displayed already
'                        If rsCRFElement.EOF Then
'                            Exit For
'                        End If
'
'                        'ic 22/12/00
'                        'added code to change negative x y coordinates to 0
'                        lTmpY = rsCRFElement!Y
'                        nTmpX = rsCRFElement!X
'                        If lTmpY <= 0 Then lTmpY = 0
'                        If nTmpX <= 0 Then nTmpX = 0
'
'                        'this area is not already used by another cell, so we must put something in it
'
'                        'If rsCRFElement!Y <= lRow(nRow) And (rsCRFElement!X <= lCol(nCol) Or rsCRFElement!ControlType = ControlType.Line) Then
'
'                        'ic 22/12/00
'                        'changed condition to use variables assigned above
'                        If lTmpY <= lRow(nRow) And (nTmpX <= lCol(nCol) Or rsCRFElement!ControlType = ControlType.Line) Then
'                            'the cell contains an element
'                            'calculate the height and width of the cell, and the number of rows and cols that it spans...
'
'                            nElementHeight = ElementHeight(rsStudyDefinition, rsCRFElement)
'                            nRowSpan = 0
'                            Do
'                                nRowSpan = nRowSpan + 1
'                                If nRow + nRowSpan = nRowCount + 1 Then
'                                    'the bottom of the cell is at the end of the table, so it can be as tall as it needs
'                                    lCellHeight = nElementHeight
'                                Else
'                                    'the height of the cell must fit an exact number of rows
'                                    lCellHeight = lRow(nRow + nRowSpan) - lRow(nRow)
'                                End If
'                            Loop Until lCellHeight >= nElementHeight
'                            If nRowSpan > 1 Then
'                                'remember to leave following rows empty for spanned cell
'                                nCellUsed(nCol) = nRowSpan - 1
'                            End If
'
'                            nElementWidth = ElementWidth(rsStudyDefinition, rsCRFElement, rsCRFPage)
'                            nColSpan = 0
'                            Do
'                                nColSpan = nColSpan + 1
'                                If nCol + nColSpan = nColCount + 1 Then
'                                    'the right of this cell is at the right of the table, so it can be as wide as it needs
'                                    nCellWidth = nElementWidth
'                                Else
'                                    'the cell must fit into an exact number of columns
'                                    nCellWidth = lCol(nCol + nColSpan) - lCol(nCol)
'                                End If
'                                If nColSpan > 1 Then
'                                    'remember to leave following columns empty for spanned cell
'                                    nCellUsed(nCol + nColSpan - 1) = nRowSpan
'                                End If
'                            Loop Until nCellWidth >= nElementWidth
'
'
'
'                            'Print #nFileNumber, "    <td align=left valign=top" & _
'                            '    " rowspan=" & LTrim(str(nRowSpan)) & " colspan=" & LTrim(str(nColSpan)) & _
'                            '    " height=" & LTrim(str(lCellHeight / nYSCALE)) & " width=" & LTrim(str(nCellWidth / nXSCALE)) & ">";
'
'
'
'                            'ic 29/11/00
'                            'added 'round' function call to avoid writing heights and widths with excessive decimal places
'                            Print #nFileNumber, "    <td align=left valign=top" & _
'                                " rowspan=" & LTrim(str(nRowSpan)) & " colspan=" & LTrim(str(nColSpan)) & _
'                                " height=" & LTrim(str(Round(lCellHeight / nYSCALE, 1))) & " width=" & LTrim(str(Round(nCellWidth / nXSCALE, 1))) & ">";
'
'
'                            'Print #nFileNumber, CellContents(rsStudyDefinition, rsCRFElement, rsCRFPage);
'
'                            'ic 01/12/00
'                            'added condition so that if the element width or height is 0 (it cannot be found) it is omitted
'                            'ic 22/12/00
'                            'added nElementHeight, nElementWidth to function call so that images can be dynamically sized
'                            If (rsCRFElement!ControlType <> ControlType.Picture) _
'                            Or (rsCRFElement!ControlType = ControlType.Picture And nElementHeight > 0 Or nElementWidth > 0) _
'                            Then Print #nFileNumber, CellContents(rsStudyDefinition, rsCRFElement, rsCRFPage, nElementHeight, nElementWidth);
'
'
'
'                            rsCRFElement.MoveNext
'
'                        Else
'                            'the cell contains a placeholder
'                            Print #nFileNumber, "    <td>";
'
'
''                            If nCol = nColCount Then
''                                nCellWidth = 1
''                            Else
''                                nCellWidth = CInt((lCol(nCol + 1) - lCol(nCol)) / nXSCALE)
''                            End If
''                            If nRow = nRowCount Then
''                                lCellHeight = 1
''                            Else
''                                lCellHeight = CInt((lRow(nRow + 1) - lRow(nRow)) / nYSCALE)
''                            End If
''
''
''                            Print #nFileNumber, "<img src=dummy_img.gif" & _
''                            " height=" & LTrim(str(lCellHeight)) & _
''                            " width=" & LTrim(str(nCellWidth)) & ">";
'
'
'
'
'                            'ic 28/11/00
'                            'added if condition so that image spacers are only drawn in the first row and column
'                            If (nCol = 0) Or (nRow = 0) Then
'
'                                If nCol = nColCount Then
'                                    nCellWidth = 1
'                                Else
'                                    nCellWidth = CInt((lCol(nCol + 1) - lCol(nCol)) / nXSCALE)
'                                End If
'                                If nRow = nRowCount Then
'                                    lCellHeight = 1
'                                Else
'                                    lCellHeight = CInt((lRow(nRow + 1) - lRow(nRow)) / nYSCALE)
'                                End If
'
'
'                                'ic 29/11/00
'                                'added if conditions so that:
'                                '   if an image is maintaining width, its height is only 1 pixel and vice-versa
'                                '   if an image height or width is 0, it is increased to 1
'                                If nCol <> 0 Or lCellHeight = 0 Then lCellHeight = 1
'                                If nRow <> 0 Or nCellWidth = 0 Then nCellWidth = 1
'
'
'                                Print #nFileNumber, "<img src=dummy_img.gif" & _
'                                    " height=" & LTrim(str(lCellHeight)) & _
'                                    " width=" & LTrim(str(nCellWidth)) & ">";
'
'
'                            'ic 28/11/00
'                            'end of added if condition
'                            End If
'
'
'                        End If
'                        Print #nFileNumber, "</td>"
'                    Else
'                        'this area is already used by a spanned cell
'                        nCellUsed(nCol) = nCellUsed(nCol) - 1
'                    End If
'                Next nCol
'                Print #nFileNumber, "</tr>"
'            Next nRow
'        End If
'
'        'write HTML end
'        Print #nFileNumber, "</table></form></body>"
'        Print #nFileNumber, "<Img src=""dummy_img.gif"">"
'        Print #nFileNumber, "</html>"
'        rsCRFElement.Close
'        Set rsCRFElement = Nothing
'        rsDataItem.Close
'        Set rsDataItem = Nothing
'        Print #nFileNumber, "<!--#include file=""CloseDataConnection.txt"" -->"
'        Print #nFileNumber, "<!-- #Include file=""formCyclelist.asp"" -->"
'        Print #nFileNumber, "<!--#include file=""eventHandlers.asp""-->"
'
'        Print #nFileNumber, "<% response.write session(""P"").GetFormArezzoDetails(session(""sRoleCode""),request.querystring(""trial"") ,request.querystring(""VersionId""), request.querystring(""trialname""), request.querystring(""site"") ,clng(request.querystring(""person"")) , cint(request.querystring(""CRFPageId"")), clng(request.querystring(""crfpagetaskid"")), lCase(request.querystring(""VisitCode"")), lCase(request.querystring(""CRFPageCode"")), session(""F5003"") ) %>"
'
'        Print #nFileNumber, "<script language =""javascript"">"
'        Print #nFileNumber, "function fnInitialisePage()"
'        Print #nFileNumber, "{"
'        Print #nFileNumber, "   ApplyInitialFormat();"
'        Print #nFileNumber, "   DisplayForm();"
'        Print #nFileNumber, "   load_Links();"
'        Print #nFileNumber, "}"
'        Print #nFileNumber, "</script>"
'
'        ' End Individual Form ---------------------------------
'        Print #nFileNumber, "<%end if%>"    'matches role function if
'        Close #nFileNumber
'        rsCRFPage.MoveNext
'    Loop ' end CRFpage
'
'    CreateHTMLComponents lClinicalTrialId, nVersionId
'
'    'destroy recordsets
'    rsStudyDefinition.Close
'    Set rsStudyDefinition = Nothing
'    rsCRFPage.Close
'    Set rsCRFPage = Nothing
'
'    Exit Sub
'ErrHandler:
'  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "CreateNetscapeFiles", "modHTML.bas")
'        Case OnErrorAction.Ignore
'            Resume Next
'        Case OnErrorAction.Retry
'            Resume
'        Case OnErrorAction.QuitMACRO
'            Call ExitMACRO
'            Call MACROEnd
'   End Select
'End Sub
             
                       
             
'------------------------------------------------------------------------------------------------------
Private Function ElementHeight(ByRef rsStudyDefinition As ADODB.Recordset, _
                                ByRef rsCRFElement As ADODB.Recordset) As Integer
'------------------------------------------------------------------------------------------------------
'Calculates the height (in twips) of a CRFElement, based on the font size and the number of lines.
'------------------------------------------------------------------------------------------------------
'
'ic 01/12/00    added check to see if picture exists before loading it. this was causing a run time error (logged 53)
'ic  22/12/00   added code so that if a picture has been positioned partly off the top of the page, height is reduced
'------------------------------------------------------------------------------------------------------

    Const sgPIXELSPERPOINT As Single = 1.1743
    Const nYSCALE As Integer = 15
    
    Dim nFontSize As Integer
    Dim oPicture As Picture
    
    'determine the font size
    If rsCRFElement!FontSize = 0 Then
        nFontSize = rsStudyDefinition!DefaultFontSize
    Else
        nFontSize = rsCRFElement!FontSize
    End If
    
    Select Case rsCRFElement!ControlType
        Case ControlType.TextBox, ControlType.PopUp, ControlType.Calendar, ControlType.RichTextBox
            'the height of these elements is 1 row of text plus 9 pixels for control border
            ElementHeight = (nFontSize * sgPIXELSPERPOINT + 9) * nYSCALE
        Case ControlType.OptionButtons, ControlType.PushButtons
            'height depends on the number of options
            If nFontSize < 13 Then
                'the font is small enough that extra space must be allowed for the option buttons themselves
                nFontSize = 13
            End If
            ElementHeight = nFontSize * sgPIXELSPERPOINT * nYSCALE
        Case ControlType.Attachment
            'fixed height to display an <input type=file>
            ElementHeight = nYSCALE * 25
        Case ControlType.Line
            'fixed height
            ElementHeight = nYSCALE
        Case ControlType.Comment
            'height depends on number of carriage returns in text
            If IsNull(rsCRFElement!Caption) Then
                ElementHeight = 0
            Else
                ElementHeight = nFontSize * sgPIXELSPERPOINT * (CountStrsInStr(rsCRFElement!Caption, vbNewLine) + 1) * nYSCALE
            End If
        Case ControlType.Picture
        
            'ic 01/12/00
            'added check to see if picture exists before loading it. this was causing a run time error (logged 53)
            If Dir(gsDOCUMENTS_PATH & rsCRFElement!Caption) <> "" Then
            
                'the height is the height of the picture
                Set oPicture = LoadPicture(gsDOCUMENTS_PATH & rsCRFElement!Caption)
                ElementHeight = frmGenerateHTML.ScaleY(oPicture.Height, vbHimetric, vbPixels) * nYSCALE
                Set oPicture = Nothing
                
                'ic 22/12/00
                'added so that if a picture has been positioned partly off the top of the page
                'in study definition (its y coordinate is negative), its height is reduced by that amount
                If rsCRFElement!Y < 0 Then ElementHeight = ElementHeight - ((rsCRFElement!Y * -1) * nYSCALE)
            
            'ic 01/12/00
            'added condition to return 0 height if the picture is not found, it can then be omitted from the html page
            Else
                MsgBox "Image " & gsDOCUMENTS_PATH & rsCRFElement!Caption & " cannot be found!", vbOKOnly
                ElementHeight = 0
            End If
        
    End Select
    
    'ElementHeight = 100

End Function

'------------------------------------------------------------------------------------------------------
Private Function ElementWidth(ByRef rsStudyDefinition As ADODB.Recordset, _
                                ByRef rsCRFElement As ADODB.Recordset, _
                                ByRef rsCRFPage As ADODB.Recordset) As Integer
'------------------------------------------------------------------------------------------------------
'Calculates the width (in twips) of a CRFElement, based on the size of the font and the number of
'characters in the longest line.
'Allowance is also made for the Status Icon where required.
'------------------------------------------------------------------------------------------------------
'MLM 19/9/00: Check for null captions
'ic  01/12/00   added check to see if picture exists before loading it. this was causing a run time error (logged 53)
'ic  22/12/00   added code so that if a picture has been positioned partly off the side of the page, width is reduced
'------------------------------------------------------------------------------------------------------

    Const sgPIXELSPERPOINT As Single = 0.75
    Const nSTATUSICON As Integer = 400 'extra width for status icon and associated white space
    
    Dim nFontSize As Integer
    Dim rsDataItem As ADODB.Recordset
    Dim oPicture As Picture
    Dim sSQL As String
    Dim sComment As String
    Dim nPosition As Integer
    Dim nCommentLength As Integer

    'determine the font size
    If rsCRFElement!FontSize = 0 Then
        nFontSize = rsStudyDefinition!DefaultFontSize
    Else
        nFontSize = rsCRFElement!FontSize
    End If

    If rsCRFElement!DataItemId > 0 Then
        'retrieve additional details from the DataItem table
        sSQL = "SELECT * FROM DataItem" & _
            " WHERE DataItemId = " & rsCRFElement!DataItemId & _
            " AND ClinicalTrialId = " & rsStudyDefinition!ClinicalTrialId & _
            " AND VersionId = " & rsStudyDefinition!VersionId
        Set rsDataItem = New ADODB.Recordset
        rsDataItem.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    End If

    Select Case rsCRFElement!ControlType
        Case ControlType.Calendar, ControlType.RichTextBox, ControlType.TextBox
            ElementWidth = (nFontSize * rsDataItem!DataItemLength * sgPIXELSPERPOINT + 9) * nXSCALE + nSTATUSICON
        
        Case ControlType.OptionButtons, ControlType.PushButtons, ControlType.PopUp
            'leave a bit of extra space, either for the option buttons or the drop-down button
            ElementWidth = (nFontSize * rsDataItem!DataItemLength * sgPIXELSPERPOINT + 20) * nXSCALE + nSTATUSICON
        
        Case ControlType.Comment
            If IsNull(rsCRFElement!Caption) Then    'MLM 19/9/00 check for null caption
                sComment = ""
            Else
                sComment = rsCRFElement!Caption
            End If
            If rsCRFElement!DataItemId > 0 Then
                'this is really a caption.. may need to allow space for question number or units of measurement
                If rsCRFPage!displayNumbers = 1 Then
                    sComment = LTrim(str(rsCRFElement!FieldOrder)) & ". " & sComment
                End If
                If Not (IsNull(rsDataItem!UnitOfMeasurement) Or rsDataItem!UnitOfMeasurement = "") Then
                    sComment = sComment & " (" & rsDataItem!UnitOfMeasurement & ")"
                End If
            End If
            'the comment may be broken into several lines by CRLFs; work out the length of the longest line
            nPosition = 1
            nCommentLength = 0
            Do
                If InStr(nPosition, sComment, vbNewLine) = 0 Then
                    If Len(sComment) - nPosition + 1 > nCommentLength Then
                        nCommentLength = Len(sComment) - nPosition + 1
                    End If
                    Exit Do
                Else
                    If InStr(nPosition, sComment, vbNewLine) - nPosition > nCommentLength Then
                        nCommentLength = InStr(nPosition, sComment, vbNewLine) - nPosition
                    End If
                    nPosition = InStr(nPosition, sComment, vbNewLine) + 1
                End If
            Loop
            ElementWidth = nFontSize * nCommentLength * sgPIXELSPERPOINT * nXSCALE
        
        Case ControlType.Line
            ElementWidth = 8515
        
        Case ControlType.Attachment
            ElementWidth = 277 * nXSCALE
        
        Case ControlType.Picture
            
            'ic 01/12/00
            'added check to see if picture exists before loading it. this was causing a run time error (logged 53)
            If Dir(gsDOCUMENTS_PATH & rsCRFElement!Caption) <> "" Then
            
                Set oPicture = LoadPicture(gsDOCUMENTS_PATH & rsCRFElement!Caption)
                ElementWidth = frmGenerateHTML.ScaleY(oPicture.Width, vbHimetric, vbPixels) * nXSCALE
                Set oPicture = Nothing
                
                'ic 22/12/00
                'added so that if a picture has been positioned partly off the side of the page
                'in study definition (its x coordinate is negative), its width is reduced by that amount
                If rsCRFElement!X < 0 Then ElementWidth = ElementWidth - ((rsCRFElement!X * -1) * nXSCALE)
            
            'ic 01/12/00
            'added condition to return 0 width if the picture is not found, it can then be omitted from the html page
            Else
                MsgBox "Image " & gsDOCUMENTS_PATH & rsCRFElement!Caption & " cannot be found!", vbOKOnly
                ElementWidth = 0
            End If
            
            
    End Select
    
    If rsCRFElement!DataItemId > 0 Then
        rsDataItem.Close
        Set rsDataItem = Nothing
    End If
    
    'ElementWidth = 1200

End Function

''------------------------------------------------------------------------------------------------------
'Private Function CellContents(ByRef rsStudyDefinition As ADODB.Recordset, _
'                                ByRef rsCRFElement As ADODB.Recordset, _
'                                ByRef rsCRFPage As ADODB.Recordset, _
'                                ByRef nElHt As Integer, ByRef nElWt As Integer) As String
''------------------------------------------------------------------------------------------------------
''Returns the HTML required for createNetscapefiles to display one CRFElement,
''including associated Status Icon and javascript where appropriate.
''------------------------------------------------------------------------------------------------------
''MLM 19/9/00: Check for null captions
''ic  10/11/00   added call to CreateFormObject in Case statements
''ic  14/11/00   added event handler call to elements
''ic  17/11/00   added case option for multimedia element
''ic  23/11/00   added variable declaration for dropdown lists
''ic  29/11/00   added variable declaration for element width
''               added code to prevent wide textboxes
''ic  30/11/00   added code to contain elements and images in table
''ic  06/12/00   applied style to radio controls
''ic  22/12/00   changed code to write radiobuttons into asp rather than by calling 'Radio' function from asp
''               removed code that writes elements in container tables as it prevents tabbing
''               changed element width calculation for textboxes
''               added height and width parameters to image code to resize images if x,y coordinates are negative
''               added height and width arguements to CellContents function to size images
''               added condition so that if a text element is hidden, a hidden field is written
''------------------------------------------------------------------------------------------------------
'
'    Dim rsDataItem As ADODB.Recordset
'    Dim rsValueData As ADODB.Recordset
'    Dim sComment As String
'    Dim sStyle As String
'    Dim sSQL As String
'    Dim sTemp As String
'
'
'    'ic 23/11/00
'    'added code for dropdown lists
'    Dim nLengthOfItem As Integer
'    Dim nMaxLengthSoFar As Integer
'    Dim nRCount As Integer
'
'    'ic 29/11/00
'    'added code for element width
'    Dim nEWidth As Integer
'
'
'
'    'build up a string with the correct style attributes
'    If rsCRFElement!FontColour = 0 Then
'        sStyle = """color: " & GetHexColour(rsStudyDefinition!DefaultFontColour)
'    Else
'        sStyle = """color: " & GetHexColour(rsCRFElement!FontColour)
'    End If
'    If rsCRFElement!FontSize = 0 Then
'        'use the study's default font attributes
'            sStyle = sStyle & "; font-family: " & rsStudyDefinition!DefaultFontName & "; font-weight: "
'        If rsStudyDefinition!DefaultFontBold = 1 Then
'            sStyle = sStyle & "bold"
'        Else
'            sStyle = sStyle & "normal"
'        End If
'        sStyle = sStyle & "; font-style: "
'        If rsStudyDefinition!DefaultFontItalic = 1 Then
'            sStyle = sStyle & "italic"
'        Else
'            sStyle = sStyle & "normal"
'        End If
'        sStyle = sStyle & "; font-size: " & rsStudyDefinition!DefaultFontSize & "pt"""
'    Else
'        'use the specified details for the individual CRFElement
'            sStyle = sStyle & "; font-family: " & rsCRFElement!FontName & "; font-weight: "
'        If rsCRFElement!FontBold = 1 Then
'            sStyle = sStyle & "bold"
'        Else
'            sStyle = sStyle & "normal"
'        End If
'        sStyle = sStyle & "; font-style: "
'        If rsCRFElement!FontItalic = 1 Then
'            sStyle = sStyle & "italic"
'        Else
'            sStyle = sStyle & "normal"
'        End If
'        sStyle = sStyle & "; font-size: " & rsCRFElement!FontSize & "pt"""
'    End If
'    'Debug.Print sStyle
'    If rsCRFElement!DataItemId > 0 Then
'        'retrieve additional details from the DataItem table
'        sSQL = "SELECT * FROM DataItem" & _
'            " WHERE DataItemId = " & rsCRFElement!DataItemId & _
'            " AND ClinicalTrialId = " & rsStudyDefinition!ClinicalTrialId & _
'            " AND VersionId = " & rsStudyDefinition!VersionId
'        Set rsDataItem = New ADODB.Recordset
'        rsDataItem.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
'    End If
'
'    Select Case rsCRFElement!ControlType
'        Case ControlType.Line
'            CellContents = "<hr size=2 width=100% noshade>"
'
'            'ic 29/11/00
'            'testing to see the effects a horizontal rule has on the layout if it begins in the wrong column
'            'discovered this happens when empty captions are discarded on the form, fixed in CreateNetscapeFiles
'            'CellContents = "*HORIZONTAL LINE*"
'
'        Case ControlType.Comment
'            If IsNull(rsCRFElement!Caption) Then    'MLM 19/9/00 check for null caption
'                sComment = ""
'            Else
'                sComment = rsCRFElement!Caption
'            End If
'            If rsCRFElement!DataItemId > 0 Then
'                'this is really a caption.. may need to allow space for question number or units of measurement
'                If rsCRFPage!displayNumbers = 1 Then
'                    sComment = LTrim(str(rsCRFElement!FieldOrder)) & ". " & sComment
'                End If
'                If Not (IsNull(rsDataItem!UnitOfMeasurement) Or rsDataItem!UnitOfMeasurement = "") Then
'                    sComment = sComment & " (" & rsDataItem!UnitOfMeasurement & ")"
'                End If
'            End If
'            CellContents = "<span style=" & sStyle & ">" & ReplaceAsciiCRwithHTML(sComment) & "</span>"
'
'        Case ControlType.Picture
'            'CellContents = "<img src=""" & rsCRFElement!Caption & """>"
'
'            'ic 22/12/00
'            'added parameters to so that if image is partly off page, it is resized
'            CellContents = "<img src='" & rsCRFElement!Caption & "' height='" & (nElHt / nYSCALE) & "' width='" & (nElWt / nXSCALE) & "'>"
'
'        Case ControlType.Calendar, ControlType.TextBox, ControlType.RichTextBox
'            'CellContents = "<input type=text size=" & LTrim(str(rsDataItem!DataItemLength)) & " style=" & sStyle & ">" & _
'            '    "<img src=missing_status.gif width=16 height=16>"
'
'            'ic 10/11/00
'            'call to write javascript object for this dataitem
'            CellContents = CreateFormObject(rsDataItem, rsCRFElement, rsStudyDefinition)
'
'
'            'ic 22/12/00
'            'added condition so that if a text element is hidden, a hidden field is written
'            If rsCRFElement!Hidden = 0 Then
'                'ic 29/11/00
'                'added to prevent excessively wide textboxes
'                'ic 22/12/00
'                'changed nEWidth calculation for first if clause to multiply by 0.6 for more accurate sized textbox
'                If RTrim(rsDataItem!DataItemFormat) > "" Then
'                    nEWidth = Round(Len(RTrim(rsDataItem!DataItemFormat)) * 0.6)
'                ElseIf rsDataItem!DataItemLength > 0 And rsDataItem!DataItemLength < 60 Then
'                    'since the 'width' value of a textbox is less then the actual number of chars it can hold
'                    '(a textbox of width '1' can hold approxiamately 2-5 chars) im going to halve the dataitemlength
'                    'value to get a rough textbox width. note: im adding .5 so that if the dataitemlength = 1, the width
'                    'will also be 1, and not wrongly halved and rounded down to 0
'                    nEWidth = Round((rsDataItem!DataItemLength / 2) + 0.5)
'                ElseIf rsDataItem!DataItemLength >= 60 Then
'                    'prevents textbox from falling off page
'                    nEWidth = 50
'                Else
'                    'Default to a 5 character wide box
'                    nEWidth = 5
'                End If
'
'
'                'ic 14/11/00
'                'add event handler function call to cell object
'                'ic 30/11/00
'                'draw table around element and image to keep them together
'                'ic 22/12/00
'                'removed table from around element as this is preventing tabbing
'                CellContents = CellContents & "<input type=text name=txt" & LCase(rsDataItem!DataItemCode) & " size=" & nEWidth _
'                                            & " style=" & sStyle & " onfocus='DisabledHandler(this);' " & " onblur='RefreshForm(document.myform.txt" _
'                                            & LCase(rsDataItem!DataItemCode) & ");'>" _
'                                            & "<%=SetStatusImage(getCurrentStatus(" & rsCRFElement!CRFElementId & "),""txt" _
'                                            & LCase(rsDataItem!DataItemCode) & """)%>"
'            Else
'                CellContents = CellContents & "<INPUT TYPE='HIDDEN' NAME='hid" & LCase(rsDataItem!DataItemCode) & "'>"
'            End If
'
'
'        Case ControlType.PopUp
'            'sTemp = "<span style=" & sStyle & "><select>"
'
'            'ic 23/11/00
'            'added object
'            sTemp = CreateFormObject(rsDataItem, rsCRFElement, rsStudyDefinition)
'
'
'            'ic 20/11/00
'            'adding name and events to dropdown list
'            'ic 30/11/00
'            'draw table around dropdown box and image to keep them together
'            'ic 22/12/00
'            'removed table from around element as this is preventing tabbing
'            sTemp = sTemp & "<select Disabled tabindex='" & rsCRFElement!FieldOrder & "' Name ='opt" & LCase$(rsDataItem!DataItemCode) & "' " _
'                          & "onblur='RefreshForm(document.myform.opt" & LCase$(rsDataItem!DataItemCode) & ");' " _
'                          & "onfocus='DisabledHandler(this);'>"
'
'
'            sSQL = "SELECT ValueCode, ItemValue FROM ValueData" & _
'                " WHERE DataItemId = " & rsCRFElement!DataItemId & _
'                " AND ClinicalTrialId = " & rsStudyDefinition!ClinicalTrialId & _
'                " AND VersionId = " & rsStudyDefinition!VersionId & _
'                " AND Active = 1 ORDER BY ValueOrder"
'            Set rsValueData = New ADODB.Recordset
'            rsValueData.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
'
'
'            'Do While Not rsValueData.EOF
'            '    sTemp = sTemp & "<option>" & rsValueData!ItemValue & "</option>"
'            '    rsValueData.MoveNext
'            'Loop
'            'CellContents = sTemp & "</select></span><img src=missing_status.gif width=16 height=16>"
'            'rsValueData.Close
'
'
'            'ic 23/11/00
'            'added code for dropdown lists
'            With rsValueData
'                If .RecordCount > 0 Then
'                    .MoveLast
'
'                    sTemp = sTemp & "<%" _
'                                  & "redim v(" & .RecordCount & ")" & vbNewLine _
'                                  & "redim c(" & .RecordCount & ")" & vbNewLine
'
'                    nLengthOfItem = 0
'                    nMaxLengthSoFar = 0
'                    nRCount = 1
'
'                    .MoveFirst
'                    Do While Not .EOF
'                        sTemp = sTemp & "c(" & nRCount & ") =""" & !ValueCode & """" & vbNewLine _
'                                      & "v(" & nRCount & ") =""" & !ItemValue & """" & vbNewLine
'
'                        nRCount = nRCount + 1
'                        nLengthOfItem = Len(!ItemValue)
'                        If nLengthOfItem > nMaxLengthSoFar Then nMaxLengthSoFar = nLengthOfItem
'                        .MoveNext
'                    Loop
'
'                    .Close
'                End If
'            End With
'
'            Set rsValueData = Nothing
'
'            sTemp = sTemp & "SelValues getResponseValue(" & rsCRFElement!CRFElementId & ") %>" & vbNewLine _
'                          & "</select>" _
'                          & "<%=SetStatusImage(getCurrentStatus(" & rsCRFElement!CRFElementId & "),""opt" & LCase(rsDataItem!DataItemCode) & """)%>"
'
'
'            CellContents = sTemp
'
'        Case ControlType.OptionButtons, ControlType.PushButtons
'
'            'sTemp = "<table border=0 cellpadding=0 cellspacing=0><tr><td><span style=" & sStyle & ">"
'
'            sSQL = "SELECT ValueCode, ItemValue FROM ValueData" & _
'                " WHERE DataItemId = " & rsCRFElement!DataItemId & _
'                " AND ClinicalTrialId = " & rsStudyDefinition!ClinicalTrialId & _
'                " AND VersionId = " & rsStudyDefinition!VersionId & _
'                " AND Active = 1 ORDER BY ValueOrder"
'            Set rsValueData = New ADODB.Recordset
'            rsValueData.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
'
'
'            'Do While Not rsValueData.EOF
'            '    sTemp = sTemp & "<input type=radio>" & rsValueData!ItemValue & "<br>"
'            '    rsValueData.MoveNext
'            'Loop
'            'CellContents = sTemp & "</span></td><td valign=top><img src=missing_status.gif width=16 height=16></td></tr></table>"
'
'
'            'ic 23/11/00
'            'added code for radios
'            With rsValueData
'                If .RecordCount > 0 Then
''                     .MoveLast
''
''                    sTemp = "<%" _
''                          & "redim v(" & .RecordCount & ")" & vbNewLine _
''                          & "redim c(" & .RecordCount & ")" & vbNewLine
''
''                    nRCount = 1
''                    nLengthOfItem = 0
''                    nMaxLengthSoFar = 0
''
''                    .MoveFirst
''                    Do While Not .EOF
''                        sTemp = sTemp & "c(" & nRCount & ") =""" & !ValueCode & """" & vbNewLine _
''                                      & "v(" & nRCount & ") =""" & !ItemValue & """" & vbNewLine
''
''                        nRCount = nRCount + 1
''                        nLengthOfItem = Len(!ItemValue)
''                        If nLengthOfItem > nMaxLengthSoFar Then nMaxLengthSoFar = nLengthOfItem
''
''                        .MoveNext
''                    Loop
''
''
''
''                    'draw table around radio group and image to keep them together
''                    'ic 06/12/00
''                    'applied style to radio controls
''                    'ic 22/12/00
''                    'removed table from around element as this is preventing tabbing
''                    sTemp = sTemp & "%>" _
''                                  & "<span style=" & sStyle & ">" _
''                                  & "<%" _
''                                  & " Radio " & """opt" & LCase$(rsDataItem!DataItemCode) & """, """ _
''                                  & LCase$(rsDataItem!DataItemCode) & """, " & rsCRFElement!FieldOrder & ",getResponseValue(" _
''                                  & rsCRFElement!CRFElementId & ") %>" _
''                                  & "</span>" & vbNewLine _
''                                  & CreateFormObject(rsDataItem, rsCRFElement, rsStudyDefinition) & vbNewLine _
''                                  & "<%=SetStatusImage(getCurrentStatus(" & rsCRFElement!CRFElementId & "), ""opt" _
''                                  & LCase(rsDataItem!DataItemCode) & """) %>" & vbNewLine
'
'
'                    'ic 22/12/00
'                    'changed code to write radiobuttons into asp rather than by calling 'Radio' function from asp
'                    sTemp = sTemp & "<span style=" & sStyle & ">"
'
'                    'loop through radio buttons
'                    Do While Not .EOF
'                        'write attributes
'                        sTemp = sTemp & "<input type='radio' DISABLED tabindex='" & rsCRFElement!FieldOrder _
'                        & "' name='opt" & LCase$(rsDataItem!DataItemCode) & "' value='" & !ValueCode & "' "
'
'                        'write event handlers
'                        sTemp = sTemp & "onfocus='showRadioClearButton(this);return false' onClick='RefreshForm(this);'>" _
'                        & !ItemValue & "<br>"
'
'                        .MoveNext
'                    Loop
'
'                    'remove the last '<br>' so that status image will be on same line
'                    sTemp = Left(sTemp, (Len(sTemp) - 4)) & "</span>&nbsp;"
'
'                    'write setstatusimage code
'                    sTemp = sTemp & "<%=SetStatusImage(getCurrentStatus(" & rsCRFElement!CRFElementId & "), ""opt" _
'                    & LCase(rsDataItem!DataItemCode) & """)%>"
'
'                    'write radio button object
'                    sTemp = sTemp & CreateFormObject(rsDataItem, rsCRFElement, rsStudyDefinition)
'
'
'                End If
'                .Close
'            End With
'
'            Set rsValueData = Nothing
'            CellContents = sTemp
'
'        'ic 17/11/00
'        'add case statement for multimedia element
'        Case ControlType.Attachment
'            CellContents = CreateFormObject(rsDataItem, rsCRFElement, rsStudyDefinition)
'
'            CellContents = CellContents & "<input Disabled type='file' TabIndex='" & rsCRFElement!FieldOrder & "' " _
'            & "Class='T' name='att" & LCase(rsDataItem!DataItemCode) & "' onfocus='DisabledHandler(this);' onblur='RefreshForm(document.myform.att" _
'            & LCase(rsDataItem!DataItemCode) & ");'> " _
'            & "<%=SetStatusImage(getCurrentStatus(" & rsCRFElement!CRFElementId & "),""att" & LCase(rsDataItem!DataItemCode) & """)%>"
'
'    End Select
'
'
'End Function
'------------------------------------------------------------------------------------------------------
'Private Function GetDataValue(sDataItemCode As String) As String
'------------------------------------------------------------------------------------------------------
'JL 26/05/00. Changed from R(""ResponseValue"") to R(""ValueCode""). So that the javascript Validation can compare datavalues
'   can't do this using the string value (e.g. sex=0 instead of sex="Female"). Remove brackets.
'JL 27/07/00. Removed references to recordset R (defined in datavalues.asp) as no longer being used.
' Use VBscript function getValueCode instead.
'JL 25/09/00. This file is now replaced by Function CreateFormObject
'------------------------------------------------------------------------------------------------------
'Dim sDataValue As String
'
'    On Error GoTo ErrHandler
'    'JL 26/05/00. Changed from R(""ResponseValue"") to R(""ValueCode""). So that the javascript Validation can compare datavalues
'    'can't do this using the string value (e.g. sex=0 instead of sex="Female"). Remove brackets.
'    'JL 27/07/00. Removed references to recordset R (defined in datavalues.asp) as no longer being used.
'    'Use VBscript function getValueCode.
'    sDataValue = sDataValue & "<script language =""javascript"" > " & vbNewLine
'    'sDataValue = sDataValue & "<% If ((Not R.eof OR Not bEmptyRecordSet) AND hasResponseValue(lThisId)) Then %>" & vbNewLine
'    sDataValue = sDataValue & LCase(sDataItemCode) & "=""" & "<%= getValueCode(lthisId) %>" & """;" & vbNewLine
'    'sDataValue = sDataValue & "<% Else %>" & vbNewLine
'    'sDataValue = sDataValue & LCase(sDataItemCode) & "=""" & """" & vbNewLine
'    'sDataValue = sDataValue & "<% End If %>" & vbNewLine
'    sDataValue = sDataValue & "</script>" & vbNewLine
'
'    GetDataValue = sDataValue
'
'    Exit Function
'ErrHandler:
'  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "getDataValue", "modHTML.bas")
'        Case OnErrorAction.Ignore
'            Resume Next
'        Case OnErrorAction.Retry
'            Resume
'        Case OnErrorAction.QuitMACRO
'            Call ExitMACRO
'            End
'   End Select
'
'End Function

'------------------------------------------------------------------------------------------------------
Private Function CreateFormObject(ByRef rsDataItem1 As ADODB.Recordset, _
                                    ByRef rsCRFElement1 As ADODB.Recordset, _
                                    ByRef rsStudyDefinition1 As ADODB.Recordset) As String
'------------------------------------------------------------------------------------------------------
'JL Created 18/09/00.
'Initialise each form elements values.
'Form objects have the following properties :
'   .value  :- Contains what the user typed in the field or what is stored in the Database.
'   .CurrentValue
'   .PreviousValue
'   .Status
'   .PreviousStatus
'   .Disabled
'   .Dirty
'   .Format
'------------------------------------------------------------------------------------------------------
'Revisions:
'   MLM 13/10/00:   Tidied up and modified to create global JavaScript objects
'                   (as opposed to adding properties to form objects)
'ic  19/11/00       added 'o' prefix to object
'------------------------------------------------------------------------------------------------------
Dim sResult As String
Dim sQuestion As String
Dim sValue As String

    On Error GoTo ErrHandler
    
    sQuestion = LCase(rsDataItem1!DataItemCode)
    
    'sResult = "<script language=javascript>" & vbNewLine & _
    '    "var " & sQuestion & "=new Object();" & vbNewLine
    
    
    'ic 19/11/00
    'added 'o' prefix to object declaration
    sResult = "<script language=javascript>" & vbNewLine & _
        "var o" & sQuestion & "=new Object();" & vbNewLine
    
    
    'VALUE for CATEGORY types.
    If rsCRFElement1!ControlType = ControlType.OptionButtons _
        Or rsCRFElement1!ControlType = ControlType.PushButtons _
        Or rsCRFElement1!ControlType = ControlType.PopUp Then
        'Category questions use their ValueCodes as values
        sValue = "<%=getValueCode(" & rsCRFElement1!CRFElementId & ")%>"
    Else
        'Other types of questions all use ResponseValue
        sValue = "<%=getResponseValue(" & rsCRFElement1!CRFElementId & ")%>"
    End If
    
    'build up JavaScript to set the properties of the question object
    'sResult = sResult & sQuestion & ".CurrentValue='" & sValue & "';" & vbNewLine & _
    '                    sQuestion & ".PreviousValue='" & sValue & "';" & vbNewLine & _
    '                    sQuestion & ".Status=<%=getCurrentStatus(" & rsCRFElement1!CRFElementId & ")%>;" & vbNewLine & _
    '                    sQuestion & ".PreviousStatus=<%=getCurrentStatus(" & rsCRFElement1!CRFElementId & ")%>;" & vbNewLine & _
    '                    sQuestion & ".Dirty=true;" & vbNewLine & _
    '                    sQuestion & ".Disabled=true;" & vbNewLine
    
    
    'ic 19/11/00
    'added 'o' prefix to object property assignment
    sResult = sResult & "o" & sQuestion & ".CurrentValue='" & sValue & "';" & vbNewLine & _
                        "o" & sQuestion & ".PreviousValue='" & sValue & "';" & vbNewLine & _
                        "o" & sQuestion & ".Status=<%=getCurrentStatus(" & rsCRFElement1!CRFElementId & ")%>;" & vbNewLine & _
                        "o" & sQuestion & ".PreviousStatus=<%=getCurrentStatus(" & rsCRFElement1!CRFElementId & ")%>;" & vbNewLine & _
                        "o" & sQuestion & ".Dirty=true;" & vbNewLine & _
                        "o" & sQuestion & ".Disabled=true;" & vbNewLine
    
    
    'Format property
    If rsDataItem1!DataType = DataType.Date And (IsNull(rsDataItem1!DataItemFormat) Or RTrim(rsDataItem1!DataItemFormat) = "") Then
        'if the question is a date but doesn't have a format specified, use the default date format for the study definition
        'sResult = sResult & sQuestion & ".Format='" & rsStudyDefinition1!StandardDateFormat & "';" & vbNewLine
        
        'ic 19/11/00
        'added 'o' prefix to object property assignment
        sResult = sResult & "o" & sQuestion & ".Format='" & rsStudyDefinition1!StandardDateFormat & "';" & vbNewLine
    Else
        'otherwise, use the question's own format (which might be empty)
        
        'ic 19/11/00
        'added 'o' prefix to object property assignment.
        sResult = sResult & "o" & sQuestion & ".Format='" & rsDataItem1!DataItemFormat & "';" & vbNewLine
    End If
    
    sResult = sResult & "</script>" & vbNewLine
    
    CreateFormObject = sResult
    
    Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "CreateFormObject", "modHTML.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select

End Function

''------------------------------------------------------------------------------------------------------
'Public Function printCaption(ByVal rsCRFElement1 As ADODB.Recordset, ByVal rsStudyDefinition1 As ADODB.Recordset, ByVal rsDataItem1 As ADODB.Recordset _
'    , ByVal blDisplayNumbers As Boolean, ByVal bIEflag As Boolean, ByVal nDefaultFSize As Integer, ByVal lDefaultFColour As Long _
'    , ByVal sDefaultFfamily As String, ByVal nDefaultFWeight As Integer, ByVal nDefaultFStyle As Integer)
''------------------------------------------------------------------------------------------------------
''------------------------------------------------------------------------------------------------------
'Dim msCaption As String
''nIEFLAG = 0 if InterNet explorer.
'
'On Error GoTo ErrHandler
''Print caption
'        'For all fields.
'
'       If rsCRFElement1!Caption <> "" Or rsCRFElement1!Caption <> Null Then
'
'            'JL 8/02/01 Changed name of caption div tags from element id to dataItemCode to pass to JS function setCaptionGrey/Back.
'
'            'ic 21/09/2001
'            'added if condition so that div id for radio captions can be used for div that surrounds the entire radio group and
'            'caption group allowing font colours to be changed with innerhtml
'            If (rsCRFElement1!ControlType = ControlType.OptionButtons) _
'            Or (rsCRFElement1!ControlType = ControlType.PushButtons) Then
'                msCaption = "<div id=" & Chr(34) & "f_" & LCase(rsDataItem1!DataItemCode) & "_rCapDiv" & Chr(34) & _
'                               " style=" & Chr$(34) & msCaption & "left:" _
'                               & CLng(rsCRFElement1!CaptionX / nIeXSCALE) & ";top:" & CLng(rsCRFElement1!CaptionY / nIeYSCALE) & _
'                               Chr$(34) & ">" & vbCrLf
'            Else
'                msCaption = "<div id=" & Chr(34) & "f_" & LCase(rsDataItem1!DataItemCode) & "_CapDiv" & Chr(34) & _
'                               " style=" & Chr$(34) & msCaption & "left:" _
'                               & CLng(rsCRFElement1!CaptionX / nIeXSCALE) & ";top:" & CLng(rsCRFElement1!CaptionY / nIeYSCALE) & _
'                               Chr$(34) & ">" & vbCrLf
'            End If
'
'            msCaption = msCaption & "<font"
'            msCaption = msCaption & printFontAttributes(rsCRFElement1!FontColour, rsCRFElement1!FontSize, rsCRFElement1!FontName _
'                , rsStudyDefinition1!DefaultFontColour, rsStudyDefinition1!DefaultFontSize, rsStudyDefinition1!DefaultFontName, bIEflag)
'            msCaption = msCaption & FieldAttributes_Explorer(rsCRFElement1!FontBold, rsCRFElement1!FontItalic, rsCRFElement1!FieldOrder, rsCRFElement1!Caption _
'                               , blDisplayNumbers, rsDataItem1!UnitOfMeasurement)
'            msCaption = msCaption & "</font>" & vbCrLf
''            msCaption = msCaption & "</p>"
'            msCaption = msCaption & "</div>"
'
'            printCaption = msCaption
'      End If
'Exit Function
'ErrHandler:
'  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "printCaption", "modHTML.bas")
'        Case OnErrorAction.Ignore
'            Resume Next
'        Case OnErrorAction.Retry
'            Resume
'        Case OnErrorAction.QuitMACRO
'            Call ExitMACRO
'            Call MACROEnd
'   End Select
'
'End Function

''-----------------------------------------------------------------------------------------------'
'Public Function printNetscapeCaption(ByVal vCRFElement As ADODB.Recordset, ByVal vStudyDefinition As ADODB.Recordset, ByVal vDataItem As ADODB.Recordset _
'    , ByVal msPositionAttr As String, ByVal nDisplayNumbers As Boolean, ByVal nIEFLAG As Integer, ByVal nDefaultFSize As Integer, ByVal lDefaultFColour As Long _
'    , ByVal sDefaultFfamily As String, ByVal nDefaultFWeight As Integer, ByVal nDefaultFStyle As Integer)
''-----------------------------------------------------------------------------------------------'
''04/08/00. JL This function returns the styles AND sets the CaptionX and CaptionY co-ordinates.
''-----------------------------------------------------------------------------------------------'
'
'Dim msCaption As String
'
'            'Changed layer width from 400 to 700, if too short a line break is inserted in the text.
'            If vCRFElement!Caption <> "" Or vCRFElement!Caption <> Null Then
'                msCaption = "<LAYER width=""700"">" 'layer tag enclosing both textbox and label. The default font
'                                'values has to be re-printed for each layer in netscape as SS go out of scope
'                                'outside the layer its defined in.
'
'                    'div contains the style elements for the caption - sDefaultFontAttr
'                msCaption = msCaption & "<DIV id=" & Chr$(34) & "C" & vCRFElement!CRFElementId & Chr$(34) & _
'                " style=" & Chr$(34)
'                '04/08/00 jl
'
'                'SIZE must come 1st in the style string.
'                If (nDefaultFSize = vCRFElement!FontSize) Or vCRFElement!FontSize = 0 Then 'If vCRFElement!FontSize then user has not specified a font size so use the default.
'                    'Then use the default and assign the it to the style string since this will override the font tag.
'                    'If nDefaultFSize <> vCRFElement!FontSize then the user has explicitly changed the font and we dont want
'                    'to assign it to the style string since it will take presidence over the font tag.
'                    msCaption = msCaption & "font-size:" & getNNFontSize(vStudyDefinition!DefaultFontSize) & ";"
'                Else    'Not a default
'                    msCaption = msCaption & "font-size:" & getNNFontSize(vCRFElement!FontSize) & ";"
'                End If
'                'COLOUR
'                If lDefaultFColour = vCRFElement!FontColour Then
'                    msCaption = msCaption & "font-color:" & GetHexColour(lDefaultFColour) & ";"
'                ElseIf (lDefaultFColour <> vCRFElement!FontColour) And vCRFElement!FontColour <> 0 Then
'                    msCaption = msCaption & "font-color:" & GetHexColour(vCRFElement!FontColour) & ";"
'                Else
'                    msCaption = msCaption & "font-color:" & GetHexColour(lDefaultFColour) & ";"
'                End If
'                'FONT FAMILY
'                If sDefaultFfamily = vCRFElement!FontName Then
'                    msCaption = msCaption & "font-family:" & sDefaultFfamily & ";"
'                ElseIf (sDefaultFfamily <> vCRFElement!FontName) And Not IsNull(vCRFElement!FontName) Then
'                    msCaption = msCaption & "font-family:" & vCRFElement!FontName & ";"
'                Else  'User has not specified a font, use default.
'                    msCaption = msCaption & "font-family:" & sDefaultFfamily & ";"
'                End If
'                'FONT WEIGHT
'                If nDefaultFWeight = vCRFElement!FontBold Then
'                    If nDefaultFWeight = 1 Then
'                        msCaption = msCaption & "font-weight:bold;"
'                    Else
'
'                    End If
'                Else
'                    If vCRFElement!FontBold = 1 Then
'                        msCaption = msCaption & "font-weight:bold;"
'                    Else
'                        'default is normal, setting it medium causes it to be bold
'                    End If
'                End If
'                'FONT STYLE
'                If nDefaultFStyle = vCRFElement!FontItalic Then
'                    If nDefaultFStyle = 1 Then
'                        msCaption = msCaption & "font-style:italic;"
'                    Else
'                        'default is normal, setting it normal causes it to be bold
'                    End If
'                Else
'
'                    If vCRFElement!FontItalic = 1 Then
'                        msCaption = msCaption & "font-style:italic;"
'                    Else
'                        msCaption = msCaption & "font-style:normal;"
'                    End If
'                End If
'                msCaption = msCaption & msPositionAttr & ";left:" _
'                               & CLng(vCRFElement!CaptionX / nXSCALE) & ";top:" & CLng(vCRFElement!CaptionY / nYSCALE) & Chr$(34)
'
'                msCaption = msCaption & ">" & vbNewLine 'End the STYLE tags
'
'                msCaption = msCaption & FieldAttributes(vCRFElement!FieldOrder, vCRFElement!Caption _
'                                , nDisplayNumbers, vDataItem!UnitOfMeasurement)
'
'                msCaption = msCaption & "</DIV></LAYER>" & vbNewLine
'
'            End If
'
'printNetscapeCaption = msCaption
'
'Exit Function
'ErrHandler:
'  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "printNetscapeCaption", "modHTML.bas")
'        Case OnErrorAction.Ignore
'            Resume Next
'        Case OnErrorAction.Retry
'            Resume
'        Case OnErrorAction.QuitMACRO
'            Call ExitMACRO
'            Call MACROEnd
'   End Select
'
'
'End Function

''------------------------------------------------------------------------------------------------------
'Public Function getNetscapeStyles(ByVal vCRFElement As ADODB.Recordset, ByVal vStudyDefinition As ADODB.Recordset _
'    , ByVal msPositionAttr As String, ByVal nDefaultFontSize As Integer, ByVal lDefaultFColour As Long _
'    , ByVal sDefaultFfamily As String, ByVal nDefaultFWeight As Integer, ByVal nDefaultFStyle As Integer)
''------------------------------------------------------------------------------------------------------
''JL 05/08/00 This returns the styles only. Including seting the X and Y co-ordinates.
''------------------------------------------------------------------------------------------------------
'Dim sFontStyle As String
'
'                'COLOUR
'                If lDefaultFColour = vCRFElement!FontColour Then
'                    sFontStyle = sFontStyle & "font-color:" & GetHexColour(lDefaultFColour) & ";"
'                ElseIf (lDefaultFColour <> vCRFElement!FontColour) And vCRFElement!FontColour <> 0 Then
'                    sFontStyle = sFontStyle & "font-color:" & GetHexColour(vCRFElement!FontColour) & ";"
'                Else
'                    sFontStyle = sFontStyle & "font-color:" & GetHexColour(lDefaultFColour) & ";"
'                End If
'                'FONT FAMILY
'                If sDefaultFfamily = vCRFElement!FontName Then
'                    sFontStyle = sFontStyle & "font-family:" & nDefaultFWeight & ";"
'                ElseIf (nDefaultFWeight <> vCRFElement!FontName) And Not IsNull(vCRFElement!FontName) Then
'                    sFontStyle = sFontStyle & "font-family:" & vCRFElement!FontName & ";"
'                Else
'                    sFontStyle = sFontStyle & "font-family:" & sDefaultFfamily & ";"
'                End If
'                'FONT SIZE
'                If (nDefaultFontSize = vCRFElement!FontSize) Or vCRFElement!FontSize = 0 Then 'If vCRFElement!FontSize then user has not specified a font size so use the default.
'                    'Then use the default and assign the it to the style string since this will override the font tag.
'                    'If nDefaultFontSize <> vCRFElement!FontSize then the user has explicitly changed the font and we dont want
'                    'to assign it to the style string since it will take presidence over the font tag.
'                    sFontStyle = sFontStyle & "font-size:" & getNNFontSize(vStudyDefinition!DefaultFontSize) & ";"
'                Else    'Not a default
'                    sFontStyle = sFontStyle & "font-size:" & getNNFontSize(vCRFElement!FontSize) & ";"
'                End If
'
'
'                'FONT WEIGHT
'                If nDefaultFWeight = vCRFElement!FontBold Then
'                    If nDefaultFWeight = 1 Then
'                        sFontStyle = sFontStyle & "font-weight:bold;"
'                    Else
'                        'sFontStyle = sFontStyle & "font-weight:medium;"
'                    End If
'                Else
'                    If vCRFElement!FontBold = 1 Then
'                        sFontStyle = sFontStyle & "font-style:bold;"
'                    Else
'                        'sFontStyle = sFontStyle & "font-style:medium;"
'                    End If
'                End If
'                'FONT STYLE
'                If nDefaultFStyle = vCRFElement!FontItalic Then
'                    If nDefaultFStyle = 1 Then
'                        sFontStyle = sFontStyle & "font-style:italic;"
'                    Else
'                        'sFontStyle = sFontStyle & "font-style:normal;"
'                    End If
'                 Else
'                    If vCRFElement!FontItalic = 1 Then
'                        sFontStyle = sFontStyle & "font-style:italic;"
'                    Else
'                        'sFontStyle = sFontStyle & "font-style:normal;"
'                    End If
'                End If
'                sFontStyle = sFontStyle & msPositionAttr & ";left:" _
'                               & CLng(vCRFElement!X / nXSCALE) & ";top:" & CLng(vCRFElement!Y / nYSCALE) & Chr$(34)
'
'                sFontStyle = sFontStyle & ">" & vbNewLine  'End the STYLE tags
'
'                getNetscapeStyles = sFontStyle
'
'End Function

''------------------------------------------------------------------------------------------'
'Private Function JavascriptDisableCaptions_JL(ByVal rsStudyDefinition As ADODB.Recordset, ByVal lClinicalTrialId As Long, _
'                           ByVal nVersionId As Integer, ByVal lCRFPageId As Long, ByVal blDisplayNumbers As Boolean, _
'                           ByVal sUnitOfMeasurement As String, Optional ByVal bIsIE As Variant)
''------------------------------------------------------------------------------------------'
'' Javascript function to return the Caption to be greyed out according to Skip Conditions.
'' Where given the dataItemCode for a question returns its Caption and HTML formatting tags.
'' Revisions :
'' JL 08/09/00 Created.
'' JL 12/02/01 Made changes to
''------------------------------------------------------------------------------------------'
'
'Dim rsCRFElement As ADODB.Recordset
'Dim rsSkips As ADODB.Recordset 'jl 15/3/01 Added
'Dim rsSkipCaptions As ADODB.Recordset 'jl 15/3/01 Added
'Dim sSQL As String
'Dim str As String
'
'    On Error GoTo ErrHandler
'    'JL 8/02/01 Changed Recordset so we are looping through all the itemItemCodes for the CRFPage
'    'rather than the elementIds, since question captions are identified by their dataitemCode
'    'prefixed by "C".
'    str = "<script language =""javascript"">" & vbNewLine
'
'    sSQL = "SELECT SkipCondition FROM CRFElement,DataItem " & _
'                " WHERE CRFPageId =" & lCRFPageId & _
'                " AND CRFElement.ClinicalTrialId =" & lClinicalTrialId & _
'                " AND CRFElement.ClinicalTrialId = DataItem.ClinicalTrialId" & _
'                " AND CRFElement.VersionId = DataItem.VersionId" & _
'                " AND DataItem.DataItemId= CRFElement.DataItemId" & _
'                " AND CRFElement.VersionId=" & nVersionId & _
'                " AND CRFElement.SkipCondition IS NOT Null"
'
'    Set rsSkips = New ADODB.Recordset
'    rsSkips.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
'
'    'jl 15/03/01 joined tables with ClinicalTrialId and VersionId, was incorrectly returning too many rows.
'    sSQL = "SELECT * FROM CRFElement, DataItem" & _
'                " WHERE CRFPageId =" & lCRFPageId & _
'                " AND CRFElement.ClinicalTrialId =" & lClinicalTrialId & _
'                " AND CRFElement.ClinicalTrialId = DataItem.ClinicalTrialId" & _
'                " AND CRFElement.VersionId = DataItem.VersionId" & _
'                " AND DataItem.DataItemId= CRFElement.DataItemId" & _
'                " AND CRFElement.VersionId=" & nVersionId & " ORDER BY CRFElementId"
'
'    Set rsCRFElement = New ADODB.Recordset
'    rsCRFElement.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
'
'    'str = str & "function setInitialCaptions(){" & vbNewLine & vbNewLine
'    Dim n As Integer
'    Dim nNoOfCaptions
'    Dim sPreviousSkip As String
'
'    sPreviousSkip = ""
'    Do While Not rsSkips.EOF
'
'        If sPreviousSkip <> "" And sPreviousSkip <> rsSkips!SkipCondition Then
'
'        sSQL = "SELECT Caption FROM CRFElement,DataItem " & _
'            " WHERE CRFPageId =" & lCRFPageId & _
'            " AND CRFElement.SkipCondition LIKE '" & rsSkips!SkipCondition & _
'            "' AND CRFElement.ClinicalTrialId =" & lClinicalTrialId & _
'            " AND CRFElement.ClinicalTrialId = DataItem.ClinicalTrialId" & _
'            " AND CRFElement.VersionId = DataItem.VersionId" & _
'            " AND DataItem.DataItemId= CRFElement.DataItemId" & _
'            " AND CRFElement.VersionId=" & nVersionId
'
'            Set rsSkipCaptions = New ADODB.Recordset
'            rsSkipCaptions.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
'
'    If rsSkipCaptions.RecordCount > 0 Then
'        If rsSkipCaptions.RecordCount > 0 Then
'                str = str & "var a" & n & "= new Array("
'
'
'
'                nNoOfCaptions = 0
'            Do While Not rsSkipCaptions.EOF
'                nNoOfCaptions = nNoOfCaptions + 1
'
'
'
'                str = str & "'" & rsSkipCaptions!Caption & "'"
'                If nNoOfCaptions < rsSkipCaptions.RecordCount Then
'                    str = str & ","
'                Else
'                    str = str & ")" & vbNewLine 'End Array
'                End If
'                rsSkipCaptions.MoveNext
'            Loop
'                str = str & "for (var i=0; i < " & "a" & n & ".length; " & "i++)" & vbNewLine & vbTab
'                str = str & "setCaptionGrey(a" & n & "[i]); " & vbNewLine & vbNewLine
'            Else
'                str = str & "var s" & n & "=" & "'" & rsSkipCaptions!Caption & "' "
'                str = str & "setCaptionGrey(s" & n & ") " & vbNewLine & vbNewLine
'        End If
'
'
'
'    End If
'        n = n + 1
'        End If 'Test for previous skip.
'        sPreviousSkip = rsSkips!SkipCondition
'        rsSkips.MoveNext
'    Loop
'    str = str & "}" & vbNewLine  'End Function.
'
'    'str = str & vbNewLine & "setInitialCaptions()" & vbNewLine
'
'    'JL 8/02/01 Changed argument to JS function from element id to sdataItemCode.
'    str = str & "function getCaption(sDataItemCode){" & vbNewLine & _
'    "switch (sDataItemCode){" & vbNewLine
'
'    Do While Not rsCRFElement.EOF
'        ' NCJ 24/10/00 - Use RemoveNull
'        If RemoveNull(rsCRFElement!Caption) > "" Then
''        If Not IsNull(rsCRFElement!Caption) And rsCRFElement!Caption <> "" Then
'
'            If RemoveNull(rsCRFElement!SkipCondition) > "" Then
''            If Not IsNull(rsCRFElement!SkipCondition) And rsCRFElement!SkipCondition > "" Then
'
'                'JL 12/02/01. Change case comparison string from elementid to dataitemcode.
'                str = str & vbTab & "case """ & LCase(rsCRFElement!DataItemCode) & """ :" & vbNewLine & vbTab & vbTab
'
'                str = str & "return '<font"
'
'                str = str & printFontAttributes(rsCRFElement!FontColour, rsCRFElement!FontSize, rsCRFElement!FontName _
'                    , rsStudyDefinition!DefaultFontColour, rsStudyDefinition!DefaultFontSize, rsStudyDefinition!DefaultFontName, bIsIE)
'                str = str & FieldAttributes_Explorer(rsCRFElement!FontBold, rsCRFElement!FontItalic, rsCRFElement!FieldOrder, rsCRFElement!Caption _
'                                   , blDisplayNumbers, sUnitOfMeasurement)
'                str = str & "</font>'" & vbNewLine & vbTab & vbTab
'
'                str = str & "break ;" & vbNewLine
'            End If
'
'        End If
'        rsCRFElement.MoveNext
'    Loop
'
'    str = str & vbTab & "default :" & vbNewLine & vbTab & vbTab & "return """" ;" & vbNewLine
'
'    str = str & "}" & vbNewLine & _
'            "}" & vbNewLine
'
'    str = str & "</script>" & vbNewLine
'
'    JavascriptDisableCaptions = str
'
'Exit Function
'ErrHandler:
'  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "JavascriptDisableCaptions", "modHTML")
'        Case OnErrorAction.Ignore
'            Resume Next
'        Case OnErrorAction.Retry
'            Resume
'        Case OnErrorAction.QuitMACRO
'            Call ExitMACRO
'            Call MACROEnd
'   End Select
'
'
'End Function
'------------------------------------------------------------------------------------------'
Private Function JavascriptUpdateRadioValues(ByVal lClinicalTrialId As Long, _
                           ByVal nVersionId As Integer, ByVal lCRFPageId As Long)

'------------------------------------------------------------------------------------------'
' Javascript function to update the value of the radio button value when the clear radio button is pressed.
' It takes the dataitem associated with the radio button and assigns an empty string to it.
' Revisions :
' JL 11/09/00 Created.
'------------------------------------------------------------------------------------------'
Dim rsCRFElement As ADODB.Recordset
Dim sSQL As String
Dim str As String

    sSQL = "SELECT ControlType, DataItemCode, CRFElementId, DataItemCode FROM CRFElement, DataItem WHERE " _
                & "CRFElement.CRFPageId =" & lCRFPageId & _
                " AND CRFElement.ClinicalTrialId =" & lClinicalTrialId & _
                " AND CRFElement.VersionId=" & nVersionId & _
                " AND DataItem.DataItemId = CRFElement.DataItemId " & _
                " AND DataItem.ClinicalTrialId = CRFElement.ClinicalTrialId " & _
                " AND DataItem.VersionId = CRFElement.VersionId " & _
                "ORDER BY CRFElement.CRFElementId"
                
    Set rsCRFElement = New ADODB.Recordset
    rsCRFElement.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText

    str = "<script language =""javascript"">" & vbNewLine

    str = str & "function resetRadioValue(oRadioName){" & vbNewLine & _
    "switch (oRadioName){" & vbNewLine
  
    Do While Not rsCRFElement.EOF
        If rsCRFElement!ControlType = gsOPTION_BUTTONS Or rsCRFElement!ControlType = gsPUSH_BUTTONS Then
        
            str = str & vbTab & "case """ & rsCRFElement!DataItemCode & """ :" & vbNewLine

            str = str & vbTab & vbTab & vbTab & LCase(rsCRFElement!DataItemCode) & "= """" ;" & vbNewLine
            
            str = str & vbTab & vbTab & vbTab & "break ;" & vbNewLine
        
        End If
        rsCRFElement.MoveNext
    Loop

    str = str & vbTab & "default :" & vbNewLine & vbTab & vbTab & "return """" ;" & vbNewLine
    
    str = str & "}" & vbNewLine & _
            "}" & vbNewLine
            
    str = str & "</script>" & vbNewLine
    
    JavascriptUpdateRadioValues = str

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                        "JavascriptUpdateRadioValues", "modHTML")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Function

 
'------------------------------------------------------------------------------------------'
Private Sub SetElementWidth(ByVal sDataItemFormat As String, ByVal nDataItemLength As Integer, _
                    ByVal nDisplayWidth As Integer, _
                    ByRef nMaxLength As Integer, ByRef nTextBoxCharWidth As Integer)
'------------------------------------------------------------------------------------------'
' REVISIONS
' DPH 07/05/2003 - Set max display size based on element
' DPH 05/03/2007 - Bug 2545. Default display size to actual size up to max of 60
'------------------------------------------------------------------------------------------'

    ' original char calculations
    If RTrim(sDataItemFormat) > "" Then
        nTextBoxCharWidth = Len(RTrim(sDataItemFormat))
'   ATN 25/11/99
'   nMaxlength should be the same as the length of the format string
        nMaxLength = nTextBoxCharWidth
    Else
        Select Case nDataItemLength
            ' DPH 05/03/2007 - Bug 2545. Default display size to actual size up to max of 60
            Case 0 To 60
                nTextBoxCharWidth = nDataItemLength
                nMaxLength = nDataItemLength
'            Case 0 To 3
'                nTextBoxCharWidth = 3
'                nMaxLength = nDataItemLength
'            Case 4
'                nTextBoxCharWidth = 4
'                nMaxLength = nDataItemLength
'            Case 5
'                nTextBoxCharWidth = 4
'                nMaxLength = nDataItemLength
'            Case 6
'                nTextBoxCharWidth = 6
'                nMaxLength = nDataItemLength
'            Case 7 To 10
'                nTextBoxCharWidth = 9
'                nMaxLength = nDataItemLength
'            Case 11 To 20
'                nTextBoxCharWidth = 15
'                nMaxLength = nDataItemLength
'            Case 21 To 30
'                nTextBoxCharWidth = 25
'                nMaxLength = nDataItemLength
'            Case 31 To 40
'                nTextBoxCharWidth = 35
'                nMaxLength = nDataItemLength
'            Case 41 To 50
'                nTextBoxCharWidth = 45
'                nMaxLength = nDataItemLength
            Case Is > 60
                'prevents textbox from falling off page
                nTextBoxCharWidth = 60
                nMaxLength = nDataItemLength
            Case Else
                'Default to a 60 character wide box
                nTextBoxCharWidth = 60
                nMaxLength = nDataItemLength
        End Select
    End If

    If nDisplayWidth <> NULL_INTEGER And nDisplayWidth > 0 Then
        nTextBoxCharWidth = nDisplayWidth
    End If

End Sub
'-------------------------------------------------------------------------------'
Private Function mostlyUpperCase(ByVal sValueItem As String) As Boolean
'-------------------------------------------------------------------------------'
'ValueItem is the longest string value displayed in a Pop-up control.
'JL 19/12/00. Quick fix for evaluating the number of letters in Upper Case in the longest string
'contained in a Pop-up so that the position of the status icon for the question can be calulated.
'-------------------------------------------------------------------------------'

Dim sSingleChar As String
Dim Pos As Integer
Dim nLowerCount As Integer
Dim nUpperCount As Integer
Dim stempstring As String
Const sNUMBERS = "#" 'jl 20/03/01. Treat numbers as lower case chars.

    On Error GoTo ErrHandler
'Assume that the string is all in lower case

For Pos = 1 To Len(sValueItem)
    sSingleChar = Mid(sValueItem, Pos, 1)
    
    If sSingleChar = UCase(sSingleChar) And Not sSingleChar Like sNUMBERS Then 'The string is in uppercase.
        nUpperCount = nUpperCount + 1
    Else
        nLowerCount = nLowerCount + 1
    End If
   
Next Pos

If nUpperCount >= nLowerCount Then
    mostlyUpperCase = True
Else
    mostlyUpperCase = False
End If

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "mostlyUpperCase", "modHTML.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Function
''------------------------------------------------------------------------------------------'
'Private Function JavascriptDisableCaptions(ByVal rsStudyDefinition As ADODB.Recordset, ByVal lClinicalTrialId As Long, _
'                           ByVal nVersionId As Integer, ByVal lCRFPageId As Long, ByVal blDisplayNumbers As Boolean, _
'                           ByVal sUnitOfMeasurement As String, Optional ByVal bIsIE As Variant)
''------------------------------------------------------------------------------------------'
'' Javascript function to return the Caption to be greyed out according to Skip Conditions.
'' Where given the dataItemCode for a question returns its Caption and HTML formatting tags.
'' Revisions :
'' JL 08/09/00 Created.
'' JL 12/02/01 Made changes to
''------------------------------------------------------------------------------------------'
'
'Dim rsCRFElement As ADODB.Recordset
'
'Dim sSQL As String
'Dim str As String
'
'    On Error GoTo ErrHandler
'    'JL 8/02/01 Changed Recordset so we are looping through all the itemItemCodes for the CRFPage
'    'rather than the elementIds, since question captions are identified by their dataitemCode
'    'prefixed by "C".
'
'    sSQL = "SELECT * FROM CRFElement, DataItem where " _
'                & "CRFPageId =" & lCRFPageId & _
'                " AND CRFElement.ClinicalTrialId =" & lClinicalTrialId & _
'                " AND DataItem.DataItemId= CRFElement.DataItemId" & _
'                " AND CRFElement.VersionId=" & nVersionId & " ORDER BY CRFElementId"
'
'    Set rsCRFElement = New ADODB.Recordset
'    rsCRFElement.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
'
'    str = "<script language =""javascript"">" & vbNewLine
'
'    'JL 8/02/01 Changed argument to JS function from element id to sdataItemCode.
'    str = str & "function getCaption(sDataItemCode){" & vbNewLine & _
'    "switch (sDataItemCode){" & vbNewLine
'
'    Do While Not rsCRFElement.EOF
'        ' NCJ 24/10/00 - Use RemoveNull
'        If RemoveNull(rsCRFElement!Caption) > "" Then
''        If Not IsNull(rsCRFElement!Caption) And rsCRFElement!Caption <> "" Then
'
'            If RemoveNull(rsCRFElement!SkipCondition) > "" Then
''            If Not IsNull(rsCRFElement!SkipCondition) And rsCRFElement!SkipCondition > "" Then
'
'                'JL 12/02/01. Change case comparison string from elementid to dataitemcode.
'                str = str & vbTab & "case """ & LCase(rsCRFElement!DataItemCode) & """ :" & vbNewLine & vbTab & vbTab
'
'                str = str & "return '<font"
'
'                str = str & printFontAttributes(rsCRFElement!FontColour, rsCRFElement!FontSize, rsCRFElement!FontName _
'                    , rsStudyDefinition!DefaultFontColour, rsStudyDefinition!DefaultFontSize, rsStudyDefinition!DefaultFontName, bIsIE)
'                str = str & FieldAttributes_Explorer(rsCRFElement!FontBold, rsCRFElement!FontItalic, rsCRFElement!FieldOrder, rsCRFElement!Caption _
'                                   , blDisplayNumbers, sUnitOfMeasurement)
'                str = str & "</font>'" & vbNewLine & vbTab & vbTab
'
'                str = str & "break ;" & vbNewLine
'            End If
'
'        End If
'        rsCRFElement.MoveNext
'    Loop
'
'    str = str & vbTab & "default :" & vbNewLine & vbTab & vbTab & "return """" ;" & vbNewLine
'
'    str = str & "}" & vbNewLine & _
'            "}" & vbNewLine
'
'    str = str & "</script>" & vbNewLine
'
'    JavascriptDisableCaptions = str
'
'Exit Function
'ErrHandler:
'  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "JavascriptDisableCaptions", "modHTML")
'        Case OnErrorAction.Ignore
'            Resume Next
'        Case OnErrorAction.Retry
'            Resume
'        Case OnErrorAction.QuitMACRO
'            Call ExitMACRO
'            Call MACROEnd
'   End Select
'
'
'End Function

''------------------------------------------------------------------------------------------------------
'Private Function SetFontAttributes(ByVal oCRFElement As eFormElementRO, Optional ByVal bIncludeColour As Boolean) As String
''------------------------------------------------------------------------------------------------------
''Retrieve font attributes for a CRFElement
''------------------------------------------------------------------------------------------------------
'Dim sResult As String
'
'    On Error GoTo Errlabel
'
'    sResult = ""
'
'    'don't get the colour for radio buttons
'    If bIncludeColour Then
'        If oCRFElement.FontColour <> 0 Then
'            sResult = sResult & "color:#" & GetHexColour(oCRFElement.FontColour) & "; "
'        End If
'    End If
'
'    'get the font size
'    If oCRFElement.FontSize <> 0 Then
'        sResult = sResult & "FONT-SIZE: " & oCRFElement.FontSize & " pt; "
'    End If
'
'    'check if font is bold
'    If oCRFElement.FontBold Then
'        sResult = sResult & "FONT-WEIGHT: " & "bold; "
'    End If
'
'    'check if font is italic
'    If oCRFElement.FontItalic Then
'        sResult = sResult & "FONT-STYLE: italic; "
'    End If
'
'    'get font name
'    If oCRFElement.FontName <> "" Then
'        sResult = sResult & "FONT-FAMILY: " & oCRFElement.FontName & ";"
'    End If
'
'    SetFontAttributes = sResult
'
'    Exit Function
'
'Errlabel:
'    SetFontAttributes = ""
'End Function

'--------------------------------------------------------------------------------------------------
Private Function EformHasLabQuestions(ByRef oEform As eFormRO) As Boolean
'--------------------------------------------------------------------------------------------------
'   ic 03/04/2003
'   does passed eform have any lab questions on it
'--------------------------------------------------------------------------------------------------
Dim oElement As eFormElementRO
Dim bRtn As Boolean

    bRtn = False
    For Each oElement In oEform.EFormElements
        If oElement.DataType = eDataType.LabTest Then
            bRtn = True
            Exit For
        End If
    Next
    EformHasLabQuestions = bRtn
End Function

