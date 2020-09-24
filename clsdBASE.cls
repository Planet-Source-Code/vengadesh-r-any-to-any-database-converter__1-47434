VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsdBASE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''This class module will handle
''''dBase,Paradox,... to any db operations
''''class module properties
''''=======================
''''1. error code = 400000

Option Explicit

''write code for clsSuper_Connect
Implements clsSuper

Dim dbText As ADODB.Connection
Dim rsText As ADODB.Recordset

Dim catText As ADOX.Catalog
Dim tblText As ADOX.Table

Dim strCon As String            ''Connection string
Dim strSQL As String            ''SQL string
Dim strTempDbPath As String        ''contains path to temp.mdb

Dim lngTotalRecords As Long     ''Store Total Records

Dim intPrevDestDatabaseType As Integer      ''Identify previsouly created db

Private Sub Class_Initialize()

On Error GoTo ehClassInit
    
    ErrorCode = 0
    
    Set dbText = New ADODB.Connection
    Set rsText = New ADODB.Recordset

    intPrevDestDatabaseType = 0
    
    strCon = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
             "Data Source=" & App.Path & "\temp.mdb;" & _
             "Persist Security Info=False"
         
    strTempDbPath = App.Path & "\temp.mdb"
    
Exit Sub

ehClassInit:
    ErrorCode = 400000
End Sub

Private Sub Class_Terminate()

On Error Resume Next

If rsText.State = 1 Then rsText.Close
If dbText.State = 1 Then dbText.Close

Set rsText = Nothing
Set dbText = Nothing

End Sub

Public Function clsSuper_GetFields( _
            Optional ByVal strSourceDatabaseName As String, _
            Optional intSourceDatabaseType As Integer, _
            Optional ByVal strSourceDatabaseFile As String, _
            Optional ByVal strSourceDatabaseFilePath As String, _
            Optional ByVal strSourceDatabaseTableName As String, _
            Optional ByVal strSourceDatabaseUserName As String, _
            Optional ByVal strSourceDatabasePassword As String, _
            Optional ByVal strSourceDatabaseServerName As String) As ADODB.Recordset
            
''return available fields
On Error GoTo ehGetFields

ErrorCode = 0

If dbText.State <> 1 Then dbText.Open strCon

rsText.Open "SELECT * FROM " & strSourceDatabaseTableName, dbText, adOpenStatic, adLockReadOnly

If rsText.RecordCount > 0 Then
    rsText.MoveLast
    lngTotalRecords = rsText.RecordCount
Else
    lngTotalRecords = 0
End If

''return rstext recordset
Set clsSuper_GetFields = rsText

Exit Function

ehGetFields:
    ErrorCode = 400200
End Function

Public Function clsSuper_GetTotalRecords() As Long
''return total number of records
clsSuper_GetTotalRecords = lngTotalRecords
End Function

Public Function clsSuper_Convert(strFields() As String, _
                        Optional ByVal strSourceDatabaseName As String, _
                        Optional intSourceDatabaseType As Integer, _
                        Optional ByVal strSourceDatabaseFile As String, _
                        Optional ByVal strSourceDatabaseFilePath As String, _
                        Optional ByVal strSourceDatabaseTableName As String, _
                        Optional ByVal strSourceDatabaseUserName As String, _
                        Optional ByVal strSourceDatabasePassword As String, _
                        Optional ByVal strSourceDatabaseServerName As String, _
                        Optional ByVal strDestDatabaseName As String, _
                        Optional ByVal intDestDatabaseType As Integer, _
                        Optional ByVal strDestDatabaseFile As String, _
                        Optional ByVal strDestDatabaseFilePath As String, _
                        Optional ByVal strDestDatabaseUserName As String, _
                        Optional ByVal strDestDatabasePassword As String, _
                        Optional ByVal strDestDatabaseServerName As String) As Boolean
                        
On Error GoTo ehConvert

ErrorCode = 0
                        
strSQL = ""
                                          
Select Case intDestDatabaseType

''=====================================================================
    Case Is = 1     ''Access Output
    
        ''Create a blank mdb file now (dest database)
        CreateEmptyMDB strDestDatabaseFilePath & strDestDatabaseFile
        
        ''Prepare sql
        strSQL = "SELECT " & ParseStringArray(strFields) & vbCrLf & _
                 "INTO [" & Replace(strSourceDatabaseFile, ".", "_") & "]" & vbCrLf & _
                 "IN '" & strDestDatabaseFilePath & strDestDatabaseFile & "'" & vbCrLf & _
                 "FROM [" & strSourceDatabaseTableName & "]"
                        
''=====================================================================
 
     Case Is = 11, 12, 13, 21, 22, 23, 24, 25, 31, 41, 42, 43, 51
    
        ''===========================================================
        ''we need to do the following linking process
        ''only if the source is other than access
        ''===========================================================

        ''Create a temp.mdb in the app.path (if it is not exists)
        CreateEmptyMDB strTempDbPath
       
        ''Delete existing likns if any
        RemoveLink strTempDbPath
       
        ''===========================================================
        ''Create a linked table to the text file
        ''the link will be named using the source file name
        ''with '.' replaced with '_'
        ''===========================================================
         CreateLink strTempDbPath, strSourceDatabaseFilePath, strSourceDatabaseFile, Replace(strSourceDatabaseFile, ".", "_"), intSourceDatabaseType
       
        ''===========================================================
        ''Prepare sql, don't forget to change the '.' with '_' in the
        ''strSourceDatabaseFile name
        ''===========================================================
        
        strSourceDatabaseFile = Replace(strSourceDatabaseFile, ".", "_")
        
        Select Case intDestDatabaseType
        
            Case Is = 11, 12, 13, 41, 42, 43, 51 'dBASE & Text output
            
                strSQL = "SELECT " & ParseStringArray(strFields) & vbCrLf & _
                        "INTO [" & GetDatabaseType(intDestDatabaseType) & ";Database=" & strDestDatabaseFilePath & "]." & _
                        "[" & strDestDatabaseFile & "]" & vbCrLf & _
                        "FROM [" & strSourceDatabaseFile & "]"
                        
            Case Is = 21, 22, 23, 24, 25   ''Excel output
                
                strDestDatabaseFile = Mid(strDestDatabaseFile, 1, InStr(1, strDestDatabaseFile, ".") - 1)
                
                strSQL = "SELECT " & ParseStringArray(strFields) & vbCrLf & _
                        "INTO [" & GetDatabaseType(intDestDatabaseType) & ";Database=" & strDestDatabaseFilePath & strDestDatabaseFile & "]." & _
                        "[" & strDestDatabaseFile & "]" & vbCrLf & _
                        "FROM [" & strSourceDatabaseFile & "]"
                        
            Case Is = 31    'HTML output
            
                strDestDatabaseFile = Mid(strDestDatabaseFile, 1, InStr(1, strDestDatabaseFile, ".") - 1)
                strDestDatabaseFile = strDestDatabaseFile & ".htm"
                    
                strSQL = "SELECT " & ParseStringArray(strFields) & vbCrLf & _
                        "INTO [" & GetDatabaseType(intDestDatabaseType) & ";Database=" & strDestDatabaseFilePath & "]." & _
                        "[" & strDestDatabaseFile & "]" & vbCrLf & _
                        "FROM [" & strSourceDatabaseFile & "]"
        End Select
''=====================================================================
End Select

If Not dbText.State = 0 And strSQL <> "" Then
    dbText.Execute strSQL
    intPrevDestDatabaseType = intDestDatabaseType
    clsSuper_Convert = True
    Exit Function
End If

Exit Function

ehConvert:
    ErrorCode = 400300
End Function

Public Sub CreateEmptyMDB(strMdbFileName As String)

If Dir(strMdbFileName) = "" Then
    
    Set catText = New ADOX.Catalog
    
    catText.Create "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                  "Data Source=" & strMdbFileName
                  
    Set catText = Nothing
    
End If

End Sub

Public Function CreateLink(strMdbFileName As String, strLinkedDatabaseName As String, strTableName As String, strLinkName As String, intSourceDatabaseType As Integer) As Boolean

On Error GoTo ehCreateLink

Set catText = New ADOX.Catalog
Set tblText = New ADOX.Table

catText.ActiveConnection = strCon

With tblText
    Set .ParentCatalog = catText
    .Name = strLinkName
    .Properties("Jet OLEDB:Create Link") = True
    .Properties("Jet OLEDB:Cache Link Name/Password") = False
    .Properties("Jet OLEDB:Exclusive Link") = False
    .Properties("Jet OLEDB:Table Hidden In Access") = False
    
    .Properties("Jet OLEDB:Link Provider String") = GetDatabaseType(intSourceDatabaseType)
    .Properties("Jet OLEDB:Link Datasource") = strLinkedDatabaseName
    .Properties("Jet OLEDB:Remote Table Name") = strTableName
    catText.Tables.Append tblText
End With

Set catText = Nothing
Set tblText = Nothing

CreateLink = True

Exit Function

ehCreateLink:
    ''
End Function

Public Function RemoveLink(strMdbFileName As String) As Boolean ', strDbTo As String) As Boolean

On Error GoTo ehRemoveLink

Set catText = New ADOX.Catalog
Set tblText = New ADOX.Table

catText.ActiveConnection = strCon

For Each tblText In catText.Tables
    If tblText.Type = "LINK" Then
        catText.Tables.Delete tblText.Name
    End If
Next

Set catText = Nothing
Set tblText = Nothing

RemoveLink = True

Exit Function

ehRemoveLink:
    ''
End Function

Public Function clsSuper_GetTables(Optional ByVal strSourceDatabaseName As String, _
                        Optional intSourceDatabaseType As Integer, _
                        Optional ByVal strSourceDatabaseFile As String, _
                        Optional ByVal strSourceDatabaseFilePath As String, _
                        Optional ByVal strSourceDatabaseUserName As String, _
                        Optional ByVal strSourceDatabasePassword As String, _
                        Optional ByVal strSourceDatabaseServerName As String) As String()
                        
On Error GoTo ehGetTables

Dim strResult() As String
Dim i As Integer

ErrorCode = 0

''Create a temp.mdb in the app.path (if it is not exists)
CreateEmptyMDB strTempDbPath

''Delete existing likns if any
RemoveLink strTempDbPath

''===========================================================
''Create a linked table to the dbase file
''the link will be named using the source file name
''with '.' replaced with '_'
''===========================================================
CreateLink strTempDbPath, strSourceDatabaseFilePath, strSourceDatabaseFile, Replace(strSourceDatabaseFile, ".", "_"), intSourceDatabaseType

If dbText.State <> 0 Then dbText.Close
If strCon <> "" Then dbText.Open strCon

''retreive the tables from the database

Set catText = New ADOX.Catalog
Set tblText = New ADOX.Table

catText.ActiveConnection = strCon

i = 0

''determine the number of tables
''in this database so as to redim
''the strResult() variable

For Each tblText In catText.Tables
    If tblText.Type = "LINK" Then
        ''strResult(i) = tblText.Name
        i = i + 1
    End If
Next

ReDim strResult(1 To i)

i = 1

For Each tblText In catText.Tables
    If tblText.Type = "LINK" Then
        strResult(i) = tblText.Name
        i = i + 1
    End If
Next

Set catText = Nothing
Set tblText = Nothing

If UBound(strResult) > 0 Then clsSuper_GetTables = strResult()

Exit Function

ehGetTables:
    ErrorCode = 400100
End Function
