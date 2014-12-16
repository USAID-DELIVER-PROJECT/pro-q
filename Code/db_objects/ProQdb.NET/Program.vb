Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("ProgramDB_NET.ProgramDB")> Public Class ProgramDB
	' TODO: Declare local ADO Recordset object. For example:
	'the name of the table or view used by this class
    Private Const m_strTable As String = DB_TABLE_PROGRAM
	
	'whether or not the object has been constructed
    Private m_objConn As OleDbConnection 'OleDbConnection
    Private m_objAdapter As OleDbDataAdapter '.Connection
    Private Const m_strSQL As String = "Select * From " & DB_TABLE_PROGRAM

	Private m_strID As String
	Private m_strName As String
	Private m_strNotes As String
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         private methods
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	'Class_Initialize() (constructor)
	'
	'connect to the db and load the data into memory
	'
	'lbailey
	'16 may 2002
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		'get a connection to the db
		'Set m_objConn = GetConnection(DB_DSN)
		
		'load the data into memory
		'rst.Open m_strTable, m_objConn, adOpenDynamic, adLockPessimistic, adCmdTable
		
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'Load()
	'
	'gets the specified record and populates this object with its values
	Public Function Load(ByRef strID As String) As Object
		
        Dim rst As New DataSet 'ADODB.Recordset
        Dim i As Integer
        Dim strGuid As String

        m_objConn = GetConnection(DB_DSN)

        'rst.Open(m_strTable, m_objConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockPessimistic, ADODB.CommandTypeEnum.adCmdTable)
        m_objAdapter = New OleDb.OleDbDataAdapter(m_strSQL & " WHERE GuidID = {" & strID & "}", m_objConn)
        rst = New DataSet
        m_objAdapter.Fill(rst, m_strTable)


        'select the desired record
        '.Filter = DB_TABLE_PK & " = '" & strID & "'"
        'If Not (.BOF And .EOF) Then
        For i = 0 To rst.Tables(m_strTable).Rows.Count - 1
            With rst.Tables(m_strTable).Rows(i)
                strGuid = CType((.Item("GuidID")), Guid).ToString()
                If strGuid = strID Then
                    'copy the values into the class members
                    m_strID = CType((.Item("GuidID")), Guid).ToString() '.Fields("guidID").Value
                    m_strName = .Item("strName") '.Fields("strName").Value
                    'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                    m_strNotes = IIf(IsDBNull(.Item("memnotes")), "", .Item("memnotes")) 'IIf(IsDBNull(.Fields("MemNotes").Value), "", .Fields("MemNotes").Value)

                    Exit For
                End If

            End With
        Next i

        CloseDB(m_objConn, rst) '.Close()    

        'return result code
        'UPGRADE_WARNING: Couldn't resolve default property of object Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Load = S_OK

        rst = Nothing
        m_objConn = Nothing

    End Function
	
	'-------------------------------------------
	'Update - update an existing record database
	'-------------------------------------------
	
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

                    Exit For
                End If
            Next
        End With

        'clean up
        CloseDB(m_objConn, rst)


        'UPGRADE_NOTE: Object rst may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        rst = Nothing
        'UPGRADE_NOTE: Object m_objConn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        m_objConn = Nothing

    End Sub


    '+
    'Create()
    '
    'Adds the New Record to the db using the info that we've got'
    '-
    Public Function Create() As String

        'get a connection to the db
        m_objConn = GetConnection(DB_DSN)
        Dim cb As OleDb.OleDbCommandBuilder
        Dim dsNewRow As DataRow
        Dim rst As New DataSet

        'connect to the db and get a recordset
        m_objConn = GetConnection(DB_DSN)
        m_objAdapter = New OleDb.OleDbDataAdapter(m_strSQL, m_objConn)
        m_objAdapter.Fill(rst, m_strTable)
        cb = New OleDb.OleDbCommandBuilder(m_objAdapter)

        dsNewRow = rst.Tables(m_strTable).NewRow()

        'if we don't already have an id for this, then create one
        If (Len(m_strID) < 1) Then
            'use the guid generator to create an id
            'UPGRADE_WARNING: Couldn't resolve default property of object g_objGUIDGenerator.GetGUID(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            m_strID = g_objGUIDGenerator.GetGUID()
        End If
        'stuff all of the record's properties
        dsNewRow.Item("guidID") = New Guid(m_strID) '.Fields("guidID").Value = m_strID
        dsNewRow.Item("strName") = m_strName
        dsNewRow.Item("MemNotes") = m_strNotes
        rst.Tables(m_strTable).Rows.Add(dsNewRow)

        m_objAdapter.Update(rst, m_strTable)

        'clean up
        CloseDB(m_objConn, rst)

        'return the id of the created object
        Create = m_strID

        'UPGRADE_NOTE: Object rst may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        rst = Nothing
        'UPGRADE_NOTE: Object m_objConn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        m_objConn = Nothing

    End Function
    Public Sub Delete(ByVal strID As String)

        Dim rst As New DataSet

        DeleteRecord(m_objConn, rst, m_strTable, m_strSQL, strID)
        CloseDB(m_objConn, rst)

        'UPGRADE_NOTE: Object rst may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        rst = Nothing
        'UPGRADE_NOTE: Object m_objConn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        m_objConn = Nothing

    End Sub
    Public Function ReturnRS(ByRef strStoredProc As String, ByRef aParams(,) As Object) As dataset

        Dim objDbExec As New DBExec
        m_objConn = GetConnection(DB_DSN)
        ReturnRS = objDbExec.ReturnRS(m_objConn, strStoredProc, aParams)
        'UPGRADE_NOTE: Object objDbExec may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        objDbExec = Nothing
        'UPGRADE_NOTE: Object m_objConn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        m_objConn = Nothing

    End Function
   
	
	'Class_Terminate() (destructor)
	'
	'cleans up whatever needs to be cleaned up when this object is
	'released.
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'release the connections to the db
		'Set m_objConn = Nothing
		
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	
	Private Sub Class_GetDataMember(ByRef DataMember As String, ByRef Data As Object)
		' TODO:  Return the appropriate recordset based on DataMember. For example:
		
		'Select Case DataMember
		'Case ""             ' Default
		'    Set Data = Nothing
		'Case Else           ' Default
		'    Set Data = rs
		'End Select
	End Sub
	
	'+
	'standard manipulator functions
	'
	'lbailey
	'16 may 2002
	'-
	Public Function SetID(ByRef strID As String) As Short
		'set the member
		m_strID = strID
		'set the return code
		SetID = S_OK
	End Function
	Public Function SetName(ByRef strName As String) As Short
		'set the member
		m_strName = strName
		'set the return code
		SetName = S_OK
	End Function
	
	Public Function SetNotes(ByRef strNotes As String) As Short
		'set the member
		m_strNotes = strNotes
		'set the return code
		SetNotes = S_OK
	End Function
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'standard accessor functions
	
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