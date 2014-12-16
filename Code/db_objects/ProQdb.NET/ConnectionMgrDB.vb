Public Class ConnectionMgrDB

    Private m_strCurrentConnectionString As String = My.Settings.myDBConnection
    Private Const PROVIDER_DETAILS As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
    Private Const CONNECTION_USER As String = ";User Id=admin;Password=;"


    Property CurrentConnectionString() As String
        Get
            Return m_strCurrentConnectionString
        End Get
        Set(ByVal strvalue As String)
            m_strCurrentConnectionString = strvalue
        End Set
    End Property

    Function ValidConnection(Optional ByVal m_strConnection As String = "") As Boolean

        Dim MyConnection As New OleDbConnection

        Try
            If m_strConnection.Length > 0 Then
                ' Test the Passed In Connection
                MyConnection.ConnectionString = m_strConnection
            Else
                ' Test the Set Connection
                MyConnection.ConnectionString = m_strCurrentConnectionString
            End If

            MyConnection.Open()
            MyConnection.Close()
            ValidConnection = True
        Catch ex As OleDbException
            If ex.ErrorCode = -2147467259 Then
                'Database Not Found So we will return false
                ValidConnection = False
            Else
                MsgBox("Error " & ex.ErrorCode & " - " & ex.Message, MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation)
                ValidConnection = False
            End If
        Finally
            MyConnection.Dispose()
        End Try

    End Function

    Function SetConnection(ByVal strFile As String, ByRef strConnection As String) As Boolean

        Dim strNewConnection As String

        strNewConnection = PROVIDER_DETAILS & strFile & CONNECTION_USER

        If ValidConnection(strNewConnection) = False Then
            '/ Not Valid Connection So Exit with msg
            Return False
        End If

        '/ Attempt to actually Apply The Connection
        DB_DSN = strNewConnection
        G_STRDSN = strNewConnection

        If ValidConnection(DB_DSN) = True Then
            '/ Update the MySettings Property
            My.Settings.myDBConnection = strNewConnection
            My.Settings.Save()
            strConnection = strNewConnection
            Return True
        Else
            Return False
        End If

    End Function

    Function GetExpectedMdbFile() As String

        Return Right$(CurrentConnectionString, CurrentConnectionString.Length - PROVIDER_DETAILS.Length)

    End Function

End Class
