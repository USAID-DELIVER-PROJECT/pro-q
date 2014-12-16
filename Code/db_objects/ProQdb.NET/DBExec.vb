Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("DBExec_NET.DBExec")> Public Class DBExec

    Const FINAL_VERSION As Double = 1.02
    Private m_dblVersion As Double


    Public Function ReturnRS(ByRef Conn As OleDbConnection, ByRef strStoredProc As String, ByRef aParams(,) As Object) As DataSet 'adodb.Recordset

        Dim i As Integer
        Dim prmParam As OleDbParameter 'ADODB.Parameter
        Dim rst As OleDbDataReader
        Dim rstDataset As DataSet
        rstDataset = New DataSet

        'create a command object
        Dim objCommand As New OleDbCommand
        'associate the command with the conn
        objCommand.Connection = Conn
        'objCommand.Properties(7).Value = True
        'declare that it's a stored proc
        objCommand.CommandType = CommandType.StoredProcedure
        'give it a timeout (30 sec is the default...)
        objCommand.CommandTimeout = DB_COMMAND_TIMEOUT
        'name the sp
        objCommand.CommandText = strStoredProc
        'create the sp's parameter list

        For i = 0 To UBound(aParams) - 1
            prmParam = New OleDb.OleDbParameter
            prmParam.DbType = aParams(i, 1)
            prmParam.Direction = ParameterDirection.Input
            prmParam.Value = aParams(i, 0)
            objCommand.Parameters.Add(prmParam)
        Next


        'rst.CursorType = ADODB.CursorTypeEnum.adOpenStatic
        'Conn.Open()
        rst = objCommand.ExecuteReader
        rstDataset = ConvertDataReaderToDataSet(rst, strStoredProc)
        ReturnRS = rstDataset

        rst.Close()
        Conn.Close()

    End Function
    Public Function ReturnRSNoParams(ByRef strStoredProc As String, ByRef aParams(,) As Object) As DataSet 'Collection

        On Error Resume Next

        'go get the recordset
        Dim rst As DataSet 'ADODB.Recordset
        Dim i As Integer
        Dim objConn As OleDbConnection 'OleDbConnection
        Dim objAdapter As OleDbDataAdapter

        'UPGRADE_WARNING: Couldn't resolve default property of object m_strDSN. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        objConn = GetConnection(DB_DSN)
        'rst = m_objDBExec.ReturnRS(m_objConn, strStoredProc, aParams)
        objAdapter = New OleDb.OleDbDataAdapter("Select * From " & strStoredProc, objConn)
        rst = New DataSet
        objAdapter.Fill(rst, strStoredProc)
        Dim k As Integer
        k = rst.Tables(strStoredProc).Rows.Count
        'for each record
        'Dim objAggregationDB As New AggregationDB
        'With rst
        ' For i = 0 To rst.Tables(m_strTable).Rows.Count - 1 'Do Until .EOF
        ' 'put the object in the collection
        ''stuff all of the record's properties into the object
        'objAggregationDB.Load(CType((.Tables(m_strTable).Rows(i).Item("GuidID")), Guid).ToString())

        ''add the object to the collection
        'm_cAggregations.Add(objAggregationDB)
        ''UPGRADE_NOTE: Object objAggregationDB may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        'objAggregationDB = Nothing

        ''get the next record
        'Next i 'Loop
        'CloseDB(m_objConn, rst)
        'End With

        ReturnRSNoParams = rst 'm_cAggregations
        'UPGRADE_NOTE: Object rst may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        rst = Nothing
        'UPGRADE_NOTE: Object m_objConn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        objConn = Nothing
    End Function

   
    'Public Function ReturnRSfromSQL(ByRef Conn As OleDbConnection, ByRef strSQL As String) As ADODB.Recordset

    'Dim i As Short
    'Dim prmParam As adodb.Parameter

    'create a command object
    'Dim objCommand As New ADODB.Command
    'associate the command with the conn
    'objCommand.let_ActiveConnection(Conn)
    'declare that it's a stored proc
    'objCommand.CommandType = ADODB.CommandTypeEnum.adCmdText

    'give it a timeout (30 sec is the default...)
    'objCommand.CommandTimeout = DB_COMMAND_TIMEOUT
    'name the sp
    'objCommand.CommandText = strSQL

    'create the sp's parameter list
    'For i = 0 To UBound(aParams) - 1
    '    Set prmParam = New adodb.Parameter
    '    prmParam.Type = aParams(i, 1)
    '    prmParam.Direction = adParamInput
    '    prmParam.Value = aParams(i, 0)
    '    objCommand.Parameters.Append prmParam
    'Next

    'objCommand.Prepared = True

    'ReturnRSfromSQL = objCommand.Execute()
    Public Function ReturnRSfromSQL(ByRef strStoredProc As String, ByRef strName As String) As DataSet 'Collection

        On Error Resume Next

        Dim rst As DataSet
        Dim i As Integer
        Dim objConn As OleDbConnection
        Dim objAdapter As OleDbDataAdapter

        objConn = GetConnection(DB_DSN)
        objAdapter = New OleDb.OleDbDataAdapter(strStoredProc, objConn)
        rst = New DataSet
        objAdapter.Fill(rst, strName)
        Dim k As Integer
        'k = rst.Tables(strStoredProc).Rows.Count
        k = rst.Tables(strName).Rows.Count

        ReturnRSfromSQL = rst 'm_cAggregations
        rst = Nothing
        objConn = Nothing
    End Function

    Public Function ExecuteActionQuery(ByRef strStoredProc As String, ByRef aParams(,) As Object) As OleDbDataReader 'ADODB.Recordset

        Dim i As Short
        Dim prmParam As OleDbParameter 'ADODB.Parameter

        'create a command object
        Dim objCommand As New OleDbCommand 'ADODB.Command
        'associate the command with the conn
        'objCommand.let_ActiveConnection(GetConnection())
        'declare that it's a stored proc
        'objCommand.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        'give it a timeout (30 sec is the default...)
        'objCommand.CommandTimeout = DB_COMMAND_TIMEOUT
        'name the sp
        'objCommand.CommandText = strStoredProc
        'create the sp's parameter list
        'For i = 0 To UBound(aParams) - 1
        'prmParam = New ADODB.Parameter
        'prmParam.Type = aParams(i, 1)
        'prmParam.Direction = ADODB.ParameterDirectionEnum.adParamInput
        'prmParam.Value = aParams(i, 0)
        'objCommand.Parameters.Append(prmParam)
        'Next

        'associate the command with the conn
        objCommand.Connection = GetConnection(DB_DSN)
        'objCommand.Properties(7).Value = True
        'declare that it's a stored proc
        objCommand.CommandType = CommandType.StoredProcedure
        'give it a timeout (30 sec is the default...)
        objCommand.CommandTimeout = DB_COMMAND_TIMEOUT
        'name the sp
        objCommand.CommandText = strStoredProc
        'create the sp's parameter list

        For i = 0 To UBound(aParams) - 1
            prmParam = New OleDb.OleDbParameter
            prmParam.DbType = aParams(i, 1)
            prmParam.Direction = ParameterDirection.Input
            prmParam.Value = aParams(i, 0)
            objCommand.Parameters.Add(prmParam)
        Next

        ExecuteActionQuery = objCommand.ExecuteReader

    End Function
    Public Function ConvertDataReaderToDataSet(ByVal reader As OleDbDataReader, ByRef strStoredProc As String) As DataSet

        Dim dataSet As DataSet = New DataSet()
        Dim schemaTable As DataTable = reader.GetSchemaTable()
        Dim dataTable As DataTable = New DataTable()
        Dim intCounter As Integer

        dataTable.TableName = strStoredProc

        If Not schemaTable Is Nothing Then

            For intCounter = 0 To schemaTable.Rows.Count - 1
                Dim dataRow As DataRow = schemaTable.Rows(intCounter)
                Dim columnName As String = CType(dataRow("ColumnName"), String)
                Dim column As DataColumn = New DataColumn(columnName, CType(dataRow("DataType"), Type))
                dataTable.Columns.Add(column)
            Next

            dataSet.Tables.Add(dataTable)

            While reader.Read()
                Dim dataRow As DataRow = dataTable.NewRow()

                For intCounter = 0 To reader.FieldCount - 1
                    dataRow(intCounter) = reader.GetValue(intCounter)
                Next

                dataTable.Rows.Add(dataRow)
            End While
        End If

        Return dataSet

    End Function



    '+
    ' Method    : ValidateMDB
    ' Purpose   : Check to see if the connected database is upto the latest schema
    ' Returns   : Boolean - True if uptodate
    ' Created   : 2008-AUG-04, JLeiner
    '-

    Public Function ValidateMDB() As Boolean

        m_dblVersion = getMDBVersion()

        Do Until m_dblVersion >= FINAL_VERSION
            Select Case m_dblVersion
                Case Is <= 1
                    If updateV1() = False Then Exit Do
                Case 1.01
                    If updateV1x() = False Then Exit Do
                Case Else
                    Exit Do
            End Select
        Loop

        ValidateMDB = (m_dblVersion >= FINAL_VERSION)

        Exit Function


    End Function


    Private Function getMDBVersion() As Double

        Dim tableName As String = "Version"
        Dim con As OleDb.OleDbConnection
        Dim dr As OleDbDataReader
        Dim cmdGet As OleDbCommand

        con = GetConnection(DB_DSN)

        'Get database schema
        Dim dbSchema As DataTable = con.GetOleDbSchemaTable(OleDb.OleDbSchemaGuid.Tables, _
                    New Object() {Nothing, Nothing, tableName, "TABLE"})
        con.Close()

        ' If the table exists, the count = 1
        If dbSchema.Rows.Count <> 1 Then
            Dim cmd As New OleDb.OleDbCommand("CREATE TABLE [" + tableName + "] ([dblVersion] DOUBLE)", con)
            con = GetConnection(DB_DSN)
            cmd.ExecuteNonQuery()
            con.Close()
            cmd = Nothing

            cmd = New OleDb.OleDbCommand("Insert Into [" & tableName & "] Values (1.0)", con)
            con = GetConnection(DB_DSN)
            cmd.ExecuteNonQuery()
            con.Close()
            cmd = Nothing
        End If

        'Now Return the value
        Try
            con = GetConnection(DB_DSN)
            cmdGet = New OleDbCommand("select * from [" & tableName & "]", con)
            dr = cmdGet.ExecuteReader
            While dr.Read()
                getMDBVersion = dr(0)
            End While
        Catch
        End Try
        dr.Close()
        con.Close()

    End Function

    '+
    ' Method    : updateV1
    ' Purpose   : Updates the Schema for changes for the ProQ.net upgrade.
    '           : 
    ' Parameters: 
    ' Returns   : Boolean - True if successful
    ' Created   : 2007-OCT-31, JLeiner
    '-
    Private Function updateV1() As Boolean

        'Alter the storedProcedures that need to be fixed for the changes from ProQ v1

        Dim cnn As OleDb.OleDbConnection
        Dim cmd As New OleDb.OleDbCommand
        Dim fOpenTransaction As Boolean
        Dim Trans As OleDb.OleDbTransaction
        Dim strView As String

        Try
            'Open the Connection
            cnn = GetConnection(DB_DSN)
            Trans = cnn.BeginTransaction()
            cmd.Connection = cnn
            cmd.Transaction = Trans
            fOpenTransaction = True

            '--------------------------------------------------------------------------------------
            'Update the Queries By dropping then Rebuilding
            '--------------------------------------------------------------------------------------
            'NOTE:  Can not create Parameterized Query remotely.  Will create a stored Procedure
            ' Which is not visible in the Database window.
            '--------------------------------------------------------------------------------------

            ' View qlkpKitByBrand
            '------------------------------------------------------------------------
            strView = "qlkpKitByBrand"
            cmd.CommandText = "DROP View " & strView
            cmd.ExecuteNonQuery()
            cmd.CommandText = "CREATE Procedure " & strView _
                & "(strBrandID Guid) AS " _
                & "SELECT tlkKit.* " _
                & "FROM tlkKit " _
                & "WHERE  tlkKit.guidID_Brand = [strBrandID] " _
                & "ORDER BY tlkKit.intTestsPerKit;"
            cmd.ExecuteNonQuery()


            ' qlkpKitByBrandActive
            '---------------------------------------------------------------------------
            strView = "qlkpKitByBrandActive"
            cmd.CommandText = "DROP VIEW " & strView
            cmd.ExecuteNonQuery()
            cmd.CommandText = "CREATE Procedure " & strView _
                & "(strBrandID Guid) AS " _
                & "SELECT tlkKit.* " _
                & "FROM tlkKit " _
                & "WHERE tlkKit.guidID_Brand = [strBrandID] " _
                & "And tlkKit.fInactive = 0 " _
                & "ORDER BY tlkKit.intTestsPerKit;"
            cmd.ExecuteNonQuery()


            ' qselBrandInUse
            '---------------------------------------------------------------------------
            strView = "qselBrandInUse"
            cmd.CommandText = "DROP VIEW " & strView
            cmd.ExecuteNonQuery()
            cmd.CommandText = "CREATE Procedure " & strView _
                & "(strBrandID Guid) AS " _
                & "SELECT tlkBrand.* " _
                & "FROM tlkBrand INNER JOIN tblProtocolBrand ON " _
                & "tlkBrand.guidID = tblProtocolBrand.guidID_Brand " _
                & "WHERE (((tblProtocolBrand.guidID_Brand)=[strBrandID]));"
            cmd.ExecuteNonQuery()


            ' qselKitInUse
            '---------------------------------------------------------------------------
            strView = "qselKitInUse"
            cmd.CommandText = "DROP VIEW " & strView
            cmd.ExecuteNonQuery()
            cmd.CommandText = "CREATE Procedure " & strView _
                & "(strBrandID Guid) AS " _
                & "SELECT tlkKit.* " _
                & "FROM tlkKit INNER JOIN tblProtocolBrand " _
                & "ON tlkKit.guidID = tblProtocolBrand.guidID_Kit " _
                & "WHERE (((tblProtocolBrand.guidID_Kit)=[strID]));"
            cmd.ExecuteNonQuery()

            ' qselProtocolBrandsByProtocolTestID
            '---------------------------------------------------------------------------
            strView = "qselProtocolBrandsByProtocolTestID"
            cmd.CommandText = "DROP VIEW " & strView
            cmd.ExecuteNonQuery()
            cmd.CommandText = "CREATE Procedure " & strView _
                & "(strProtocolTestID Guid) AS " _
                & "SELECT tblProtocolBrand.* " _
                & "FROM(tblProtocolBrand) " _
                & "WHERE (((tblProtocolBrand.guidID_ProtocolTest)=[strProtocolTestID]));"
            cmd.ExecuteNonQuery()

            ' qselProtLevelsByProtID
            '---------------------------------------------------------------------------
            strView = "qselProtLevelsByProtID"
            cmd.CommandText = "DROP VIEW " & strView
            cmd.ExecuteNonQuery()
            cmd.CommandText = "CREATE Procedure " & strView _
                & "(strProtocolID Guid) AS " _
                & "SELECT tblProtocolLevel.* " _
                & "FROM(tblProtocolLevel) " _
                & "WHERE (((tblProtocolLevel.guidID_Protocol)=[strProtocolID]));"
            cmd.ExecuteNonQuery()

            ' qselProtPattLevByProtPattID
            '---------------------------------------------------------------------------
            strView = "qselProtPattLevByProtPattID"
            cmd.CommandText = "DROP VIEW " & strView
            cmd.ExecuteNonQuery()
            cmd.CommandText = "CREATE Procedure " & strView _
                & "(strProtocolPatternID Guid) AS " _
                & "SELECT tlkPatternLevel.* " _
                & "FROM tlkPatternLevel " _
                & "WHERE tlkPatternLevel.guidID_ProtocolPattern =[strProtocolPatternID];"
            cmd.ExecuteNonQuery()



            Trans.Commit()
            cmd.Dispose()
            cnn.Close()
            cnn.Dispose()
            fOpenTransaction = False
            UpdateDBVersion(1.01)

            updateV1 = True
        Catch EX As Exception
            MsgBox(EX.Message, MsgBoxStyle.Exclamation, "UpdateV1")

            If fOpenTransaction = True Then
                Trans.Rollback()
                fOpenTransaction = False
            End If
            updateV1 = False
        End Try

    End Function
    
    '+
    ' Method    : updateV1x
    ' Purpose   : Updates the Schema for changes for the ProQ.net upgrade Fixed error in qselBrandinUse.
    '           : 
    ' Parameters: 
    ' Returns   : Boolean - True if successful
    ' Created   : 2008-DEC-18, JLeiner
    '-
    Private Function updateV1x() As Boolean

        'Alter the storedProcedures that need to be fixed for the changes from ProQ v1

        Dim cnn As OleDb.OleDbConnection
        Dim cmd As New OleDb.OleDbCommand
        Dim fOpenTransaction As Boolean
        Dim Trans As OleDb.OleDbTransaction
        Dim strView As String

        Try
            'Open the Connection
            cnn = GetConnection(DB_DSN)
            Trans = cnn.BeginTransaction()
            cmd.Connection = cnn
            cmd.Transaction = Trans
            fOpenTransaction = True

            '--------------------------------------------------------------------------------------
            'Update the Queries By dropping then Rebuilding
            '--------------------------------------------------------------------------------------
            'NOTE:  Can not create Parameterized Query remotely.  Will create a stored Procedure
            ' Which is not visible in the Database window.
            '--------------------------------------------------------------------------------------

            ' qselBrandInUse
            '---------------------------------------------------------------------------
            strView = "qselBrandInUse"
            cmd.CommandText = "DROP VIEW " & strView
            cmd.ExecuteNonQuery()
            cmd.CommandText = "CREATE Procedure " & strView _
                & "(strBrandID Guid) AS " _
                & "SELECT tlkBrand.* " _
                & "FROM tlkBrand INNER JOIN tblProtocolBrand ON " _
                & "tlkBrand.guidID = tblProtocolBrand.guidID_Brand " _
                & "WHERE (((tblProtocolBrand.guidID_Brand)=[strBrandID]));"
            cmd.ExecuteNonQuery()


            Trans.Commit()
            cmd.Dispose()
            cnn.Close()
            cnn.Dispose()
            fOpenTransaction = False
            UpdateDBVersion(1.02)

            updateV1x = True
        Catch EX As Exception
            MsgBox(EX.Message, MsgBoxStyle.Exclamation, "UpdateV1x")

            If fOpenTransaction = True Then
                Trans.Rollback()
                fOpenTransaction = False
            End If
            updateV1x = False
        End Try

    End Function



    '+
    ' Method    : UpdateDBVersion
    ' Purpose   : Updates the Version # in the database and sets the
    '           : schemaversion property
    ' Parameters: dblVersion - the value to set the version to
    ' Returns   : -
    ' Created   : 2007-OCT-31, JLeiner
    '-
    Private Sub UpdateDBVersion(ByVal dblVersion As Double)

        Dim cnn As OleDb.OleDbConnection
        cnn = GetConnection(DB_DSN)

        Try
            Dim cmd As New OleDbCommand("Update Version Set dblVersion = " & dblVersion, cnn)
            cmd.ExecuteNonQuery()
            cmd = Nothing
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "UpdateDBVersion")
        End Try

        m_dblVersion = dblVersion
        cnn.Close()
        cnn.Dispose()

        Exit Sub

    End Sub



End Class