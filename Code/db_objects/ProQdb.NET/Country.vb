Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("CountryDB_NET.CountryDB")> Public Class CountryDB
	'+
	'CountryDB.cls
	'
	'manages the country table in the db
	'
	'lbailey
	'6 june 2002
	'-
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         constants
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	Private Const DB_TABLE As String = DB_TABLE_COUNTRY
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         members
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	
	'db conns
    Private m_objConn As OleDbConnection 'OleDbConnection
    Private m_objAdapter As OleDbDataAdapter '.Connection
    Private Const m_strSQL As String = "Select * From " & DB_TABLE_COUNTRY
    Private Const m_strTable As String = DB_TABLE
	
	'data that means something
	Private m_strID As String
	Private m_strName As String
	Private m_strNotes As String
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         private methods
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	
	'+
	'Class_Initialize() (constructor)
	'
	'connect to the db and load the data into memory
	'
	'lbailey
	'16 may 2002
	'-
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		'get a connection to the db
		'Set m_objConn = GetConnection(DB_DSN)
		
		'load the data into memory
		'm_objRS.Open _
		'DB_TABLE, _
		'm_objConn, _
		'adOpenDynamic, _
		'adLockPessimistic, _
		'adCmdTable
		
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	
	
	'Class_Terminate() (destructor)
	'
	'cleans up whatever needs to be cleaned up when this object is
	'released.
	'
	'lbailey
	'6 june 2002
	'-
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'release the connections to the db
		'm_objRS.Close
		'Set m_objConn = Nothing
		
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         public methods
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	
	
	'+
	'Load()
	'
	'gets the specified record and populates this object with its values
	'
	'lbailey
	'6 june 2002
	'-
	Public Sub Load(ByRef strID As String)
		
        Dim rst As New DataSet 'ADODB.Recordset
        Dim i As Integer
        Dim strGuid As String

        m_objConn = GetConnection(DB_DSN)

        'rst.Open(m_strTable, m_objConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockPessimistic, ADODB.CommandTypeEnum.adCmdTable)
        m_objAdapter = New OleDb.OleDbDataAdapter(m_strSQL, m_objConn)
        rst = New DataSet
        m_objAdapter.Fill(rst, m_strTable)

        With rst
            'select the desired record
            '.Filter = DB_TABLE_PK & " = '" & strID & "'"
            'If Not (.BOF And .EOF) Then
            For i = 0 To .Tables(m_strTable).Rows.Count - 1
                strGuid = CType((.Tables(m_strTable).Rows(i).Item("GuidID")), Guid).ToString()
                If strGuid = strID Then
                    'copy the values into the class members
                    m_strID = CType((.Tables(m_strTable).Rows(i).Item("GuidID")), Guid).ToString() '.Fields("guidID").Value
                    m_strName = .Tables(m_strTable).Rows(i).Item("strName") '.Fields("strName").Value
                    'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                    m_strNotes = IIf(IsDBNull(.Tables(m_strTable).Rows(i).Item("memnotes")), "", .Tables(m_strTable).Rows(i).Item("memnotes")) 'IIf(IsDBNull(.Fields("MemNotes").Value), "", .Fields("MemNotes").Value)
                    Exit For
                End If
                '.Close()
            Next i
        End With
        CloseDB(m_objConn, rst)

        'UPGRADE_NOTE: Object rst may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        rst = Nothing
        'UPGRADE_NOTE: Object m_objConn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        m_objConn = Nothing

    End Sub
	
	
	
	'+
	'Update()
	'
	'updates the current record
	'
	'lbailey
	'6 june 2002
	'-
	Public Sub Update()
		
        Dim i As Integer
        Dim cb As OleDb.OleDbCommandBuilder
        Dim rst As New DataSet
        'connect to the db and get a recordset
        m_objConn = GetConnection(DB_DSN)
        m_objAdapter = New OleDb.OleDbDataAdapter(m_strSQL, m_objConn)
        m_objAdapter.Fill(rst, m_strTable)
        cb = New OleDb.OleDbCommandBuilder(m_objAdapter)

        With rst
            For i = 0 To .Tables(m_strTable).Rows.Count - 1
                Dim strGuid As String
                strGuid = CType(.Tables(m_strTable).Rows(i).Item(DB_TABLE_PK), Guid).ToString
                If strGuid = m_strID Then
                    'stuff all of the record's properties
                    .Tables(m_strTable).Rows(i).Item("strName") = m_strName
                    .Tables(m_strTable).Rows(i).Item("MemNotes") = m_strNotes
                    'write the record
                    m_objAdapter.Update(rst, m_strTable) '.Update()
                End If
            Next
        End With

        'clean up
        CloseDB(m_objConn, rst)

    End Sub



    '+
    'Create()
    '
    'Adds the New Record to the db using the info that we've got
    '
    'lbailey
    '6 june 2002
    '-
    Public Function Create() As String

        'get a connection to the db
        m_objConn = GetConnection(DB_DSN)
        Dim cb As OleDb.OleDbCommandBuilder
        Dim dsNewRow As DataRow
        Dim rst As New DataSet
        'connect to the db and get a recordset
        'OpenDB(m_objConn, m_objRS, m_strTable, m_strSQL)
        m_objConn = GetConnection(DB_DSN)
        m_objAdapter = New OleDb.OleDbDataAdapter(m_strSQL, m_objConn)
        m_objAdapter.Fill(rst, m_strTable)
        cb = New OleDb.OleDbCommandBuilder(m_objAdapter)

        'ensure that we're pointing at the record that we want
        'InsertOrUpdate(m_objRS, m_fIsNew, m_strID)

        dsNewRow = rst.Tables(m_strTable).NewRow()

        'if we don't already have an id for this, then create one
        If (Len(m_strID) < 1) Then
            'use the guid generator to create an id
            'UPGRADE_WARNING: Couldn't resolve default property of object g_objGUIDGenerator.GetGUID(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            m_strID = g_objGUIDGenerator.GetGUID()
        End If
        'stuff all of the record's properties
        dsNewRow.Item("guidID") = New Guid(m_strID)
        dsNewRow.Item("strName") = m_strName
        dsNewRow.Item("MemNotes") = m_strNotes
        rst.Tables(m_strTable).Rows.Add(dsNewRow)

        m_objAdapter.Update(rst, m_strTable)

        'clean up
        CloseDB(m_objConn, rst)

    End Function


    '+
    'Delete()
    '
    'deletes the current record
    '
    'lbailey
    '6 june 2002
    '-
    Public Sub Delete()

        Dim rst As New DataSet

        DeleteRecord(m_objConn, rst, m_strTable, m_strSQL, m_strID)
        CloseDB(m_objConn, rst)
    End Sub
	
	
	'+
	'standard manipulator functions
	'
	'lbailey
	'16 may 2002
	'-
	Public Sub SetID(ByRef strID As String)
		m_strID = strID
	End Sub
	
	Public Sub SetName(ByRef strName As String)
		m_strName = strName
	End Sub
	
	Public Sub SetNotes(ByRef strNotes As String)
		m_strNotes = strNotes
	End Sub
	
	
	
	'+
	'standard accessor functions
	'
	'lbailey
	'16 may 2002
	'-
	Public Function GetID() As String
		GetID = m_strID
	End Function
	
	Public Function GetName() As String
		GetName = m_strName
	End Function
	
	Public Function GetNotes() As String
		GetNotes = m_strNotes
	End Function
End Class