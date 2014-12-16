Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("MethodologyDB_NET.MethodologyDB")> Public Class MethodologyDB
	' TODO: Declare local ADO Recordset object. For example:
	'the name of the table or view used by this class
    Private Const m_strTable As String = DB_TABLE_METHODOLOGY
    Private Const m_strSQL As String = "Select * From " & DB_TABLE_METHODOLOGY

	'whether or not the object has been constructed
    Private m_objConn As OleDbConnection
    'Private m_objRS As DataSet 
	
	Private m_strID As String
	Private m_strName As String
	Private m_strNotes As String
	Private m_lngSort As Integer
	
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
		'm_objRS.Open strTable, m_objConn, adOpenDynamic, adLockPessimistic, adCmdTable
		
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	
	'+
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
		
        Dim rst As New DataSet
        Dim i As Integer
        Dim strGuid As String
        'm_objConn = GetConnection(DB_DSN)
        'rst.Open(m_strTable, m_objConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockPessimistic, ADODB.CommandTypeEnum.adCmdTable)
        rst = OpenDB(m_objConn, rst, m_strTable, m_strSQL & " WHERE GuidID = {" & strID & "}")

        'With rst
        'select the desired record
        '.Filter = DB_TABLE_PK & " = '" & strID & "'"

        'If Not (.BOF And .EOF) Then
        'copy the values into the class members
        'End If
        '.Close()
        'End With

        For i = 0 To rst.Tables(m_strTable).Rows.Count - 1
            strGuid = CType((rst.Tables(m_strTable).Rows(i).Item(DB_TABLE_PK)), Guid).ToString()
            If strGuid = strID Then
                With rst.Tables(m_strTable).Rows(i)
                    'copy the values into the class members
                    m_strID = CType((.Item("GuidID")), Guid).ToString() '.Fields("GuidID").Value
                    m_strName = .Item("strName") '.Fields("strName").Value
                    m_strNotes = IIf(IsDBNull(.Item("MemNotes")), "", .Item("MemNotes")) 'IIf(IsDBNull(.Fields("MemNotes").Value), "", .Fields("MemNotes").Value)
                    m_lngSort = .Item("lngSort") '.Fields("lngSort").Value
                End With


                Exit For
            End If
        Next

        'CloseDB(m_objConn, rst)    

        rst = Nothing
        m_objConn = Nothing
    End Sub



    '+
    'Update()
    '
    'updates the db with the current record's info
    '
    'lbailey
    '6 june 2002
    '-
    Public Sub Update()

        Dim i As Integer
        Dim cb As OleDb.OleDbCommandBuilder
        Dim objAdapter As OleDbDataAdapter
        Dim rst As New DataSet
        m_objConn = GetConnection(DB_DSN)
        'rst.Open(m_strTable, m_objConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockPessimistic, ADODB.CommandTypeEnum.adCmdTable)
        objAdapter = New OleDb.OleDbDataAdapter(m_strSQL, m_objConn)
        objAdapter.Fill(rst, m_strTable)
        cb = New OleDb.OleDbCommandBuilder(objAdapter)

        With rst
            For i = 0 To .Tables(m_strTable).Rows.Count - 1
                Dim strGuid As String
                strGuid = CType(.Tables(m_strTable).Rows(i).Item(DB_TABLE_PK), Guid).ToString
                If strGuid = m_strID Then
                    'stuff all of the record's properties
                    .Tables(m_strTable).Rows(i).Item("strName") = m_strName '.Fields("strName").Value = m_strName
                    .Tables(m_strTable).Rows(i).Item("MemNotes") = m_strNotes '.Fields("MemNotes").Value = m_strNotes
                    .Tables(m_strTable).Rows(i).Item("lngSort") = m_lngSort '.Fields("lngSort").Value = m_lngSort

                    'write the record
                    objAdapter.Update(rst, m_strTable) '.Update()
                End If
            Next
        End With

        'clean up
        CloseDB(m_objConn, rst)
    End Sub


    Public Function Create() As String
        '+
        'Create()
        '
        'Adds the New Record to the db using the info that we've got
        '
        'lbailey
        '6 june 2002
        '-

        Dim cb As OleDb.OleDbCommandBuilder
        Dim objAdapter As OleDbDataAdapter
        Dim dsNewRow As DataRow
        Dim rst As New DataSet
        m_objConn = GetConnection(DB_DSN)
        'rst.Open(m_strTable, m_objConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockPessimistic, ADODB.CommandTypeEnum.adCmdTable)
        objAdapter = New OleDb.OleDbDataAdapter(m_strSQL, m_objConn)
        objAdapter.Fill(rst, m_strTable)
        cb = New OleDb.OleDbCommandBuilder(objAdapter)

        dsNewRow = rst.Tables(m_strTable).NewRow()

        If (Len(m_strID) < 1) Then
            'use the guid generator to create an id
            m_strID = g_objGUIDGenerator.GetGUID()
        End If
        'stuff all of the record's properties
        dsNewRow.Item("guidID") = New Guid(m_strID) '.Fields("guidID").Value = m_strID
        dsNewRow.Item("strName") = m_strName '.Fields("strName").Value = m_strName
        dsNewRow.Item("MemNotes") = m_strNotes '.Fields("MemNotes").Value = m_strNotes
        dsNewRow.Item("lngSort") = m_lngSort '.Fields("lngSort").Value = m_lngSort
        rst.Tables(m_strTable).Rows.Add(dsNewRow)

        'write the record

        objAdapter.Update(rst, m_strTable)
        'return the id of the created object
        Create = m_strID
        'clean up
        CloseDB(m_objConn, rst)
    End Function

    '+
    'Delete()
    '
    'removes the current record from the db
    '
    'lbailey
    '6 june 2002
    '-
    Public Sub Delete()

        Dim rst As New DataSet
        'm_objConn = GetConnection(DB_DSN)
        'rst.Open(m_strTable, m_objConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockPessimistic, ADODB.CommandTypeEnum.adCmdTable)

        'With rst
        'select the desired record
        '.Filter = DB_TABLE_PK & " = '" & m_strID & "'"
        'delete it
        '.Delete((ADODB.AffectEnum.adAffectCurrent))
        'commit the changes
        '.UpdateBatch()

        '.Close()
        'End With
        DeleteRecord(m_objConn, rst, m_strTable, m_strSQL, m_strID)
        'UPGRADE_NOTE: Object rst may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        rst = Nothing
        'UPGRADE_NOTE: Object m_objConn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        m_objConn = Nothing
    End Sub
	
	
    Public Function ReturnRS(ByRef strStoredProc As String, ByRef aParams(,) As Object) As DataSet 'ADODB.Recordset
        Dim objDbExec As New DBExec
        m_objConn = GetConnection(DB_DSN)
        ReturnRS = objDbExec.ReturnRS(m_objConn, strStoredProc, aParams)
        'UPGRADE_NOTE: Object objDbExec may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        objDbExec = Nothing
        'UPGRADE_NOTE: Object m_objConn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        m_objConn = Nothing
    End Function
	
  
	
	
	
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
	Public Sub SetSort(ByRef lngSort As Integer)
		m_lngSort = lngSort
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
	Public Function GetSort() As String
		GetSort = CStr(m_lngSort)
	End Function
End Class