VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DirectLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------------------------
Public Function Login(ByVal sUserName As String, ByVal sPassword As String, ByRef sUserXml As String) As Long
'--------------------------------------------------------------------------------------------------
' attempt login to MACRO & return result, with user xml if successful
'--------------------------------------------------------------------------------------------------

Dim oUser As MACROUser
Dim sSecCon As String
Dim lLoginResult As Long
Dim lUserLogin As Long
Dim bUserRole As Boolean
Dim vUserRole As Variant
Dim sMessage As String
Dim vUserDatabase As Variant
Dim oStudySite As StudySite
Dim sStudyName As String
Dim sLastStudyName As String
Dim sStudyDescription As String
Dim sSiteCode As String
Dim sSiteDescription As String
Dim sSiteLocation As String
Dim oSite As Site
Dim vRole As Variant
Dim oXMLUser As New DOMDocument40
Dim oUserNode As IXMLDOMNode
Dim oDatabasesNode As IXMLDOMNode
Dim oDatabaseNode As IXMLDOMNode
Dim oStudiesNode As IXMLDOMNode
Dim oStudyNode As IXMLDOMNode
Dim oSitesNode As IXMLDOMNode
Dim oSiteNode As IXMLDOMNode

'    On Error GoTo ErrLabel
    
    ' validate username and password
    If Not ValidateUsername(sUserName) Then
        sMessage = "invalid username"
        Login = Failed
'EXIT FUNCTION
        Exit Function
    End If
    
    If Not ValidatePassword(sPassword) Then
        sMessage = "invalid password"
        Login = Failed
'EXIT FUNCTION
        Exit Function
    End If
    
    sUserXml = ""
    
    ' create user object
    Set oUser = New MACROUser
    
    ' get security connection from settings file
    sSecCon = GetSecurityCon
    
    ' attempt login through user object login
    lLoginResult = oUser.Login(sSecCon, sUserName, sPassword, "", "MACRO API", sMessage)

    ' look at result
    Select Case lLoginResult
    
    ' If failure or account disabled then don't continue with login
    Case LoginResult.Success, LoginResult.ChangePassword, LoginResult.PasswordExpired
        
        ' schedule
        Set oUserNode = oXMLUser.createElement("user")
        
        ' attributes
        ' user name
        oUserNode.Attributes.setNamedItem SetAttributeNode(oXMLUser, "name", oUser.UserName)
        ' user full name
        oUserNode.Attributes.setNamedItem SetAttributeNode(oXMLUser, "fullname", oUser.UserNameFull)
        
        ' database node
        Set oDatabasesNode = oXMLUser.createElement("databases")

        ' loop through the available databases
        For Each vUserDatabase In oUser.UserDatabases
            
            ' set up current database so can retrieve the study / site combinations
            ' 4th param - false
            If oUser.SetCurrentDatabase(sUserName, vUserDatabase, "", True, True, sMessage) Then
                
                ' if there are user roles
                If oUser.UserRoles.Count > 0 Then
                    
                    ' database node
                    Set oDatabaseNode = oXMLUser.createElement("database")
                    ' name
                    oDatabaseNode.Attributes.setNamedItem SetAttributeNode(oXMLUser, "name", oUser.DatabaseCode)
                    
                    ' create studies node
                    Set oStudiesNode = oXMLUser.createElement("studies")
                    
                    ' reset last study name
                    sLastStudyName = ""
                    
                    ' loop through user roles
                    For Each vRole In oUser.UserRoles
                    
                        ' load study site permissions
                        oUser.SetUserRole (vRole)
                        
                        ' loop through study / sites on this macro database
                        For Each oStudySite In oUser.GetStudiesSites()
                        
                            ' get study name & description
                            sStudyName = oStudySite.StudyName
                            sStudyDescription = oUser.Studies.StudyById(oStudySite.StudyId).StudyDescription
                            
                            ' if different study name
                            If sLastStudyName <> sStudyName Then
                                                                
                                ' study node
                                Set oStudyNode = oXMLUser.createElement("study")
                                ' study name
                                oStudyNode.Attributes.setNamedItem SetAttributeNode(oXMLUser, "name", sStudyName)
                                ' study description
                                oStudyNode.Attributes.setNamedItem SetAttributeNode(oXMLUser, "description", sStudyDescription)
                                
                                ' set last study name
                                sLastStudyName = sStudyName
                                
                                ' add to studies node
                                oStudiesNode.appendChild oStudyNode
                                
                                ' sites node
                                Set oSitesNode = oXMLUser.createElement("sites")

                                ' add to study node
                                oStudyNode.appendChild oSitesNode
                            End If
                            
                            ' get site name, description & country name
                            sSiteCode = oStudySite.Site
                            Set oSite = oUser.GetAllSites(oStudySite.StudyId).Item(sSiteCode)
                            sSiteDescription = oSite.SiteDescription
                            sSiteLocation = oSite.CountryName
                            ' site node
                            Set oSiteNode = oXMLUser.createElement("site")
                            ' code
                            oSiteNode.Attributes.setNamedItem SetAttributeNode(oXMLUser, "code", sSiteCode)
                            ' description
                            oSiteNode.Attributes.setNamedItem SetAttributeNode(oXMLUser, "description", sSiteDescription)
                            ' location
                            oSiteNode.Attributes.setNamedItem SetAttributeNode(oXMLUser, "location", sSiteLocation)
                            
                            ' add to study node
                            oSitesNode.appendChild oSiteNode
                        Next
                        
                    Next
                    
                    ' add studies to database node
                    oDatabaseNode.appendChild oStudiesNode
                    
                    ' add to databases node
                    oDatabasesNode.appendChild oDatabaseNode
                End If
                
            End If
            
        Next
        
        ' add databases node to user node
        oUserNode.appendChild oDatabasesNode
        
        ' Add to document
        oXMLUser.appendChild oUserNode
        
        ' collect XML
        sUserXml = oXMLUser.xml

    End Select
 
    Login = lLoginResult
    
'    Exit Function
'ErrLabel:
'    Login = LoginResult.Failed
'    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'--------------------------------------------------------------------------------------------------
Private Function GetSecurityCon() As String
'--------------------------------------------------------------------------------------------------
' function returns the connection string for the macro security db specified in the settings file
'--------------------------------------------------------------------------------------------------
    InitialiseSettingsFile True
    GetSecurityCon = GetMACROSetting(MACRO_SETTING_SECPATH, "")
    If GetSecurityCon <> "" Then
        GetSecurityCon = DecryptString(GetSecurityCon)
    End If
End Function

'------------------------------------------------------------------------------------
Private Function SetAttributeNode(doc As MSXML2.DOMDocument40, sAttributeName As String, sAttributeValue As String) As MSXML2.IXMLDOMNode
'------------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------------

Dim oAttributeNode As MSXML2.IXMLDOMNode
    
    On Error GoTo CatchAllError
    
    Set oAttributeNode = doc.createNode(NODE_ATTRIBUTE, sAttributeName, "")
    oAttributeNode.Text = sAttributeValue
    
    Set SetAttributeNode = oAttributeNode

    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|MACRO30InterfaceVB.SetAttributeNode"
End Function

'------------------------------------------------------------------------------------
Private Function SetChildNode(doc As MSXML2.DOMDocument40, sNodeName As String, sNodeValue As String)
'------------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------------

Dim oChildNode As MSXML2.IXMLDOMNode
    
    On Error GoTo CatchAllError
    
    Set oChildNode = doc.createNode(NODE_ELEMENT, sNodeName, "")
    oChildNode.Text = sNodeValue
    
    Set SetChildNode = oChildNode

    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|MACRO30InterfaceVB.SetChildNode"
End Function


