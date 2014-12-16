Option Strict Off
Option Explicit On
'UPGRADE_NOTE: Globals was upgraded to Globals_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
Module Globals_Renamed
	
	'the shared conn to the db
    Dim g_objDBConn As New OleDbConnection 'OleDbConnection
    'keep one of these in memory to create new IDs for us
    Public g_objGUIDGenerator As New jsiGuidGenerator.GUIDGen

    Public DB_DSN As String = My.Settings.myDBConnection
    Public G_STRDSN As String = My.Settings.myDBConnection

    '+
    'GetConnection()
    '
    'checks to see if we have a connection to the db.  if we don't then
    'it establishes one.  returns the connection (which is otherwise
    'accessible as a global, anyway)
    '
    'lbailey
    '16 june 2002
    '25 June 2002 jleiner - Make sure the connection is open
    '-
    Public Function GetConnection(Optional ByRef strDSN As String = "0") As OleDbConnection 'OleDbConnection
        'check to see whether the connection is already open
        On Error Resume Next
        Dim strFoo As String
10:     strFoo = g_objDBConn.ConnectionString
        If strFoo = "" Then
            'if not, then open it
            'g_objDBConn = New OleDbConnection
            g_objDBConn.ConnectionString = strDSN
            g_objDBConn.Open()
        End If
        'GoTo 10
        'make sure the connection is open
        If g_objDBConn.State = ConnectionState.Closed Then
            g_objDBConn.Open()
        End If

        'return the connection
        GetConnection = g_objDBConn

    End Function
	



	'+
	'OpenDB()
	'
	'opens the specfied db conn and creates a new recordset
	'
	'lbailey
	'26 june 2002
	'-
    Public Function OpenDB(ByRef objConn As OleDbConnection, ByRef objRS As DataSet, ByRef strTable As String, ByRef strSQL As String) As DataSet

        Dim adapter As OleDbDataAdapter
        'get a connection to the db
        objConn = GetConnection(DB_DSN)
        adapter = New OleDbDataAdapter(strSQL, objConn)
        objRS = New DataSet

        'load the data into memory
        adapter.Fill(objRS, strTable)
        'objRS.Open(strTable, objConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockPessimistic, ADODB.CommandTypeEnum.adCmdTable)
        OpenDB = objRS
    End Function
	
	
	
	'+
	'OpenDB()
	'
	'closes the db conn and frees the memory
	'
	'lbailey
	'26 june 2002
	'-
    Public Sub CloseDB(ByRef objConn As OleDbConnection, ByRef objRS As DataSet)

        'release the connections to the db
        'objRS.Close()
        'UPGRADE_NOTE: Object objRS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        'objConn.Close()

        objRS = Nothing
        'UPGRADE_NOTE: Object objConn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        If Not IsNothing(objConn) Then
            objConn.Close()
            objConn = Nothing
        End If
        'objConn = Nothing

    End Sub
	
	
	'+
	'InsertOrUpdate()
	'
	'looks at the flag, then either adds a new record, or filters down
	'to the current rec
	'
	'lbailey
	'26 june 2002
	'-
    Public Sub InsertOrUpdate(ByRef objRS As DataSet, ByRef fIsInsert As Boolean, ByRef strID As String)

        Dim myDataRow As DataRow
        'create the record, if necessary
        If (fIsInsert = True) Then
            myDataRow = objRS.Tables(objRS.Tables(0).TableName).NewRow
            myDataRow("guidID") = strID
            objRS.Tables(objRS.Tables(0).TableName).Rows.Add(myDataRow)
            'objRS.AddNew()
            'objRS.Fields("guidID").Value = strID
            fIsInsert = False
        Else
            'just filter the rs down to the record of interest
            'objRS.oFilter = DB_TABLE_PK & " ='" & strID & "'"
            Dim myView As DataView
            myView = objRS.Tables(0).DefaultView
            myView.RowFilter = DB_TABLE_PK & " ='" & strID & "'"
        End If

    End Sub
	
	
	
	'+
	'DeleteRecord()
	'
	'deletes the record referenced, from the rs/conn passed in
	'
	'lbailey
	'26 june 2002
	'-
    Public Sub DeleteRecord(ByRef objConn As OleDbConnection, ByRef objRS As DataSet, ByRef strTable As String, ByRef strSQL As String, ByRef strID As String)

        Dim i As Integer
        Dim strGuid As String
        Dim da As OleDbDataAdapter
        'get a connection to the db
        objConn = GetConnection(DB_DSN)
        da = New OleDbDataAdapter(strSQL, objConn)
        objRS = New DataSet
        da.Fill(objRS, strTable)
        Dim cb As New OleDb.OleDbCommandBuilder(da)

        'load the data into memory
        For i = 0 To objRS.Tables(strTable).Rows.Count - 1
            strGuid = CType((objRS.Tables(strTable).Rows(i).Item("GuidID")), Guid).ToString()
            If strGuid = strID Then
                objRS.Tables(strTable).Rows(i).Delete()
            End If
        Next
        da.Update(objRS, strTable)

    End Sub
End Module