
Imports System.Text.RegularExpressions
Imports ADODB

Public Class DataConnection

    'engineer (local sqlexpress)
    Public Const DB_ENGINEER_MAIN As Integer = 0
    Public Const DB_ENGINEER_PROJECT As Integer = 1

    'bakker db (bce-server)
    Public Const DB_BAKKER_MAIN As Integer = 2
    Public Const DB_BAKKER_PROJECT As Integer = 3

    'live (amsdatabase.mdb)
    Public Const DB_LIVE As Integer = 4

    'log connection (this is tha active project (engineer or bakker)) (usefull for logging in active queries)
    Public Const DB_LOG As Integer = 5


    'SQL connection string
    Private lstrCnn As String = "Provider=SQLOLEDB.1;" +
                                  "Integrated Security=SSPI;" +
                                  "Persist Security Info=False;" +
                                  "Data Source={0}; " +
                                  "Use Procedure for Prepare=0;" +
                                  "Auto Translate=True;" +
                                  "Packet Size=4096;" +
                                  "Use Encryption for Data=False;" +
                                  "Tag with column collation when possible=False;" +
                                  "Initial Catalog={1};"

    ''fallback sql connections string for bce-server
    Private lstrCnnFallback As String = "Provider=SQLOLEDB.1;" +
                                       "Uid=sa;" +
                                       "Pwd=bsei684;" +
                                       "Persist Security Info=False;" +
                                       "Data Source={0}; " +
                                       "Use Procedure for Prepare=0;" +
                                       "Auto Translate=True;" +
                                       "Packet Size=4096;" +
                                       "Use Encryption for Data=False;" +
                                       "Tag with column collation when possible=False;" +
                                       "Initial Catalog={1};"

    'ACCESS connection string
    Private lstrAccessCnn As String = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                                     "Data Source={0};"


    Private cnnArray As New ArrayList

    Sub New()
        ' max 5 different connections needed
        Me.cnnArray.Insert(0, New Connection)
        Me.cnnArray.Insert(1, New Connection)
        Me.cnnArray.Insert(2, New Connection)
        Me.cnnArray.Insert(3, New Connection)
        Me.cnnArray.Insert(4, New Connection)
        Me.cnnArray.Insert(5, New Connection)

        ' Open a Connection 
        Me.cnnArray(DataConnection.DB_ENGINEER_MAIN).ConnectionTimeout = 30
        Me.cnnArray(DataConnection.DB_ENGINEER_MAIN).CommandTimeout = 120
        Me.cnnArray(DataConnection.DB_ENGINEER_MAIN).Open(System.String.Format(Me.lstrCnn, ".\SQLEXPRESS", "BCE_MAIN"))

    End Sub
    ''' <summary>
    ''' Arraylist with connections
    ''' </summary>
    ''' <returns></returns>
    Public Function getConnections() As ArrayList
        Return Me.cnnArray
    End Function

End Class
