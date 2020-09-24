VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSuper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function GetTables(Optional ByVal strSourceDatabaseName As String, _
                        Optional intSourceDatabaseType As Integer, _
                        Optional ByVal strSourceDatabaseFile As String, _
                        Optional ByVal strSourceDatabaseFilePath As String, _
                        Optional ByVal strSourceDatabaseUserName As String, _
                        Optional ByVal strSourceDatabasePassword As String, _
                        Optional ByVal strSourceDatabaseServerName As String) As String()
''''Abstract Class
End Function

Public Function GetFields( _
            Optional ByVal strSourceDatabaseName As String, _
            Optional intSourceDatabaseType As Integer, _
            Optional ByVal strSourceDatabaseFile As String, _
            Optional ByVal strSourceDatabaseFilePath As String, _
            Optional ByVal strSourceDatabaseTableName As String, _
            Optional ByVal strSourceDatabaseUserName As String, _
            Optional ByVal strSourceDatabasePassword As String, _
            Optional ByVal strSourceDatabaseServerName As String) As ADODB.Recordset
''''Abstract Class
''''The return type requires modification
End Function

Public Function GetTotalRecords() As Long
''Abstract Class
End Function

Public Function Convert(strFields() As String, _
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
''Abstract Class
End Function

