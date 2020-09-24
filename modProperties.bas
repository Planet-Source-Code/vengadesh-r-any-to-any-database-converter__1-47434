Attribute VB_Name = "modProperties"
''''This module contains general functions
''''and other udt
''''module properties
''''=================
''''1. error code = 200000

Option Explicit

Public Enum SourceDatabaseTypeConstants
    ACCESS = 1
    DBASE = 11
    EXCEL = 21
    HTML = 31
    PARADOX = 41
    TXT = 51
    
    ''non-file based databases
    ORACLE = 101
    SQLSERVER = 111
End Enum

Public Enum DestDatabaseTypeConstants

    ACCESS_2000 = 1 '1-10
    
    DBASE_III = 11  '11-20
    DBASE_IV = 12
    DBASE_5 = 13
    
    EXCEL_3 = 21    '21-30
    EXCEL_4 = 22
    EXCEL_5 = 23
    EXCEL_97 = 24
    EXCEL_8 = 25
    
    HTML = 31       '31-40
    
    PARADOX_3 = 41  '41-50
    PARADOX_4 = 42
    PARADOX_5 = 43
    
    TXT = 51        '51-60
    
    ''non-file based databases
    ORACLE = 101    '101-110
    SQLSERVER = 111 '111-120
End Enum

Public Type DatabaseProperties

    SourceDatabaseName As String
    SourceDatabaseFile As String
    SourceDatabaseFilePath As String
    SourceDatabaseUserName As String
    SourceDatabasePassword As String
    SourceDatabaseServerName As String
    SourceDatabaseTableName As String
    SourceDatabaseType As SourceDatabaseTypeConstants
    
    DestDatabaseName As String
    DestDatabaseFile As String
    DestDatabaseFilePath As String
    DestDatabaseUserName As String
    DestDatabasePassword As String
    DestDatabaseServerName As String
    DestDatabaseTableName As String
    DestDatabaseType As DestDatabaseTypeConstants
    
End Type

''used to handle error codes
Private lngErrorCode As Long

''used to load forms
Private strFormName As String

Public Property Get ErrorCode() As Long
ErrorCode = lngErrorCode
End Property

Public Property Let ErrorCode(ByVal lngValue As Long)
lngErrorCode = lngValue
End Property

Public Function GetFileName(strFile As String) As String
''this function get complete file location info and
''return file name only

On Error GoTo ehGetFileName

Dim pos As Integer

ErrorCode = 0

pos = InStrRev(strFile, "\")
GetFileName = Mid(strFile, pos + 1, Len(strFile) - pos)

Exit Function

ehGetFileName:
    ErrorCode = 200100
End Function

Public Function GetFilePath(strFile As String, bSlash As Boolean) As String
''this function get complete file location info and
''return file path only
''bSlash flag is to determine whether
''the path will contain the \ or not

On Error GoTo ehGetFilePath

Dim pos As Integer

ErrorCode = 0

pos = InStrRev(strFile, "\")

If bSlash Then
    GetFilePath = Mid(strFile, 1, pos)
Else
    GetFilePath = Mid(strFile, 1, pos - 1)
End If

Exit Function

ehGetFilePath:
    ErrorCode = 200200
End Function

Public Function IsLoaded(strFormName As String, bLoad As Boolean) As Boolean
''''This function will receive a form name
''''and zorder flag and load / zorder a form

Dim i As Integer

For i = 0 To Forms.Count - 1
    If Forms(i).Name = strFormName Then
        ''yes given form already loaded so bring it to front
        Forms(i).ZOrder
        IsLoaded = True
        Exit Function
    End If
Next
''given form not loaded, so load and bring to front
If bLoad Then
    Select Case strFormName
        Case Is = "frmConvert"
            Load frmConvert
            frmConvert.ZOrder
    End Select
End If
End Function

Public Function FindFieldType(intFieldType As Integer) As String

Select Case intFieldType
    Case Is = 20
        FindFieldType = "adBigInt"
    Case Is = 128
        FindFieldType = "adBinary"
    Case Is = 11
        FindFieldType = "adBoolean"
    Case Is = 8
        FindFieldType = "adBSTR"
    Case Is = 136
        FindFieldType = "adChapter"
    Case Is = 129
        FindFieldType = "adChar"
    Case Is = 6
        FindFieldType = "adCurrency"
    Case Is = 7
        FindFieldType = "adDate"
    Case Is = 133
        FindFieldType = "adDBDate"
    Case Is = 134
        FindFieldType = "adDBTime"
    Case Is = 135
        FindFieldType = "adDBTimeStamp"
    Case Is = 14
        FindFieldType = "adDecimal"
    Case Is = 5
        FindFieldType = "adDouble"
    Case Is = 0
        FindFieldType = "adEmpty"
    Case Is = 10
        FindFieldType = "adError"
    Case Is = 64
        FindFieldType = "adFileTime"
    Case Is = 72
        FindFieldType = "adGUID"
    Case Is = 9
        FindFieldType = "adIDispatch"
    Case Is = 3
        FindFieldType = "adInteger"
    Case Is = 13
        FindFieldType = "adIUnknown"
    Case Is = 205
        FindFieldType = "adLongVarBinary"
    Case Is = 201
        FindFieldType = "adLongVarChar"
    Case Is = 203
        FindFieldType = "adLongVarWChar"
    Case Is = 131
        FindFieldType = "adNumeric"
    Case Is = 138
        FindFieldType = "adPropVariant"
    Case Is = 4
        FindFieldType = "adSingle"
    Case Is = 2
        FindFieldType = "adSmallInt"
    Case Is = 16
        FindFieldType = "adTinyInt"
    Case Is = 21
        FindFieldType = "adUnsignedBigInt"
    Case Is = 19
        FindFieldType = "adUnsignedInt"
    Case Is = 18
        FindFieldType = "adUnsignedSmallInt"
    Case Is = 17
        FindFieldType = "adUnsignedTinyInt"
    Case Is = 132
        FindFieldType = "adUserDefined"
    Case Is = 200
        FindFieldType = "adVarChar"
    Case Is = 204
        FindFieldType = "adVarBinary"
    Case Is = 12
        FindFieldType = "adVariant"
    Case Is = 139
        FindFieldType = "adVarNumeric"
    Case Is = 202
        FindFieldType = "adVarWChar"
    Case Is = 130
        FindFieldType = "adWChar"
    Case Default
        FindFieldType = "Unknown"
End Select

End Function

Public Function GetDatabaseType(intDestDatabaseType As Integer) As String

Select Case intDestDatabaseType
           
    Case Is = 11
        GetDatabaseType = "dBASE III"
    Case Is = 12
        GetDatabaseType = "dBASE IV"
    Case Is = 13
        GetDatabaseType = "dBASE 5.0"
    Case Is = 21
        GetDatabaseType = "Excel 3.0"
    Case Is = 22
        GetDatabaseType = "Excel 4.0"
    Case Is = 23
        GetDatabaseType = "Excel 5.0"
    Case Is = 24
        GetDatabaseType = "Excel 97"
    Case Is = 25
        GetDatabaseType = "Excel 8.0"
    Case Is = 31
        GetDatabaseType = "HTML Export"
    Case Is = 41
        GetDatabaseType = "Paradox 3.x"
    Case Is = 42
        GetDatabaseType = "Paradox 4.x"
    Case Is = 43
        GetDatabaseType = "Paradox 5.x"
    Case Is = 51
        GetDatabaseType = "Text"
    
End Select

End Function

Public Function ParseStringArray(strSelectedFields() As String) As String

Dim i As Integer
Dim strResult As String

If Len(strSelectedFields(0)) > 0 Then

For i = 0 To UBound(strSelectedFields)
    strResult = strResult & "[" & strSelectedFields(i) & "],"
Next

ParseStringArray = Mid(strResult, 1, Len(strResult) - 1)

End If

End Function

