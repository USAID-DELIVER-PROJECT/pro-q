Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("QuantMethodDB_NET.QuantMethodDB")> Public Class QuantMethodDB
	' TODO: Declare local ADO Recordset object. For example:
    'the name of the table or view used by this class
    Private Const m_strTable As String = DB_TABLE_QUANTIFICATIONMETHOD
    Private Const m_strSQL As String = "Select * from " & DB_TABLE_QUANTIFICATIONMETHOD

	'whether or not the object has been constructed
    Private m_objConn As OleDbConnection 'OleDbConnection
    'Private m_objRS As DataSet 
	
	Private m_strID As String
	Private m_strQuantificationID As String
	Private m_strMethodologyID As String
	Private m_fSelectedMethod As Boolean
	
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
	
	'Class_Terminate() (destructor)
	'
	'cleans up whatever needs to be cleaned up when this object is
	'released.
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
	
	'Load()
	'
	'gets the specified record and populates this object with its values
	Public Function Load(ByRef strID As String) As Object
		
        Dim rst As New DataSet
        Dim i As Integer
        Dim strGuid As String
        'm_objConn = GetConnection(DB_DSN)
        'rst.Open(m_strTable, m_objConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockPessimistic, ADODB.CommandTypeEnum.adCmdTable)
        rst = OpenDB(m_objConn, rst, m_strTable, m_strSQL & " WHERE GuidID = {" & strID & "}")

        'select the desired record
        '.Filter = DB_TABLE_PK & " = '" & strID & "'"
        For i = 0 To rst.Tables(m_strTable).Rows.Count - 1
            strGuid = CType((rst.Tables(m_strTable).Rows(i).Item(db_table_pk)), Guid).ToString()
            If strGuid = strID Then
                With rst.Tables(m_strTable).Rows(i)
                    'copy the values into the class members
                    m_strID = CType((.Item("GuidID")), Guid).ToString() '.Fields("GuidID").Value
                    m_strQuantificationID = CType((.Item("guidID_Quantification")), Guid).ToString() '.Fields("guidID_Quantification").Value
                    m_strMethodologyID = CType((.Item("guidID_Methodology")), Guid).ToString() '.Fields("guidID_Methodology").Value
                    m_fSelectedMethod = .Item("fSelectedMethod") '.Fields("fSelectedMethod").Value
                End With

                'CloseDB(m_objConn, rst)
                Exit For
            End If
        Next

        'return result code
        'UPGRADE_WARNING: Couldn't resolve default property of object Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Load = S_OK

        'UPGRADE_NOTE: Object rst may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        rst = Nothing
        'UPGRADE_NOTE: Object m_objConn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        m_objConn = Nothing
    End Function

    '-------------------------------------------
    'Update - update an existing record database
    '-------------------------------------------

    Public Sub Update()

        Dim i As Integer
        Dim cb As OleDb.OleDbCommandBuilder
        Dim objAdapter As OleDbDataAdapter
        Dim rst As DataSet
        Dim strGuid As String
        m_objConn = GetConnection(DB_DSN)
        rst = New DataSet
        'rst.Open(m_strTable, m_objConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockPessimistic, ADODB.CommandTypeEnum.adCmdTable)
        objAdapter = New OleDb.OleDbDataAdapter(m_strSQL, m_objConn)
        objAdapter.Fill(rst, m_strTable)
        cb = New OleDb.OleDbCommandBuilder(objAdapter)

        With rst
            For i = 0 To .Tables(m_strTable).Rows.Count - 1
                strGuid = CType((.Tables(m_strTable).Rows(i).Item(DB_TABLE_PK)), Guid).ToString()
                If strGuid = m_strID Then
                    'stuff all of the record's properties
                    .Tables(m_strTable).Rows(i).Item("guidID_Quantification") = New Guid(m_strQuantificationID) '.Fields("guidID_Quantification").Value = m_strQuantificationID
                    .Tables(m_strTable).Rows(i).Item("guidID_Methodology") = New Guid(m_strMethodologyID) '.Fields("guidID_Methodology").Value = m_strMethodologyID
                    .Tables(m_strTable).Rows(i).Item("fSelectedMethod") = (m_fSelectedMethod) '.Fields("fSelectedMethod").Value = m_fSelectedMethod

                    'write the record
                    objAdapter.Update(rst, m_strTable) '.Update()

                    Exit For
                End If
            Next
        End With

        'clean up
        CloseDB(m_objConn, rst)
    End Sub


    '+
    'Create()
    '
    'Adds the New Record to the db using the info that we've got'
    '-
    Public Function Create() As String

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
            'UPGRADE_WARNING: Couldn't resolve default property of object g_objGUIDGenerator.GetGUID(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            m_strID = g_objGUIDGenerator.GetGUID()
        End If
        'stuff all of the record's properties
        dsNewRow.Item("guidID") = New Guid(m_strID) '.Fields("guidID").Value = m_strID
        dsNewRow.Item("guidID_Quantification") = New Guid(m_strQuantificationID) '.Fields("guidID_Quantification").Value = m_strQuantificationID
        dsNewRow.Item("guidID_Methodology") = New Guid(m_strMethodologyID) '.Fields("guidID_Methodology").Value = m_strMethodologyID
        dsNewRow.Item("fSelectedMethod") = m_fSelectedMethod '.Fields("fSelectedMethod").Value = m_fSelectedMethod
        rst.Tables(m_strTable).Rows.Add(dsNewRow)

        'write the record

        objAdapter.Update(rst, m_strTable)
        'return the id of the created object
        Create = m_strID
        'clean up
        CloseDB(m_objConn, rst)
    End Function
    Public Function Delete(ByVal strID As String) As Object

        Dim rst As New DataSet
        'm_objConn = GetConnection(DB_DSN)
        'rst.Open(m_strTable, m_objConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockPessimistic, ADODB.CommandTypeEnum.adCmdTable)
        DeleteRecord(m_objConn, rst, m_strTable, m_strSQL, strID)
        'With rst
        'select the desired record
        '.Filter = DB_TABLE_PK & " = '" & strID & "'"

        'delete it
        'On Error Resume Next
        '.Delete((ADODB.AffectEnum.adAffectCurrent))
        'commit the changes
        '.UpdateBatch()
        '.Close()
        'End With

        'UPGRADE_NOTE: Object rst may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        rst = Nothing
        'UPGRADE_NOTE: Object m_objConn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        m_objConn = Nothing
    End Function
    Public Function ReturnRS(ByRef strStoredProc As String, ByRef aParams(,) As Object) As DataSet 'ADODB.Recordset

        Dim objDbExec As New DBExec
        m_objConn = GetConnection(DB_DSN)
        ReturnRS = objDbExec.ReturnRS(m_objConn, strStoredProc, aParams)

        'UPGRADE_NOTE: Object objDbExec may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        objDbExec = Nothing
        'UPGRADE_NOTE: Object m_objConn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        m_objConn = Nothing
    End Function
  
	
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
	Public Function setQuantificationID(ByRef strQuantificationID As String) As Short
		'set the member
		m_strQuantificationID = strQuantificationID
		'set the return code
		setQuantificationID = S_OK
	End Function
	
	Public Function SetSelectedMethod(ByRef fSelectedMethod As Boolean) As Short
		'set the member
		m_fSelectedMethod = fSelectedMethod
		'set the return code
		SetSelectedMethod = S_OK
	End Function
	Public Function setMethodologyID(ByRef strMethodologyID As String) As Short
		'set the member
		m_strMethodologyID = strMethodologyID
		'set the return code
		setMethodologyID = S_OK
	End Function
	
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'standard accessor functions
	
	Public Function GetID() As String
		GetID = m_strID
	End Function
	
	Public Function getQuantificationID() As String
		getQuantificationID = m_strQuantificationID
	End Function
	
	Public Function GetSelectedMethod() As Boolean
		GetSelectedMethod = m_fSelectedMethod
	End Function
	
	Public Function getMethodologyID() As String
		getMethodologyID = m_strMethodologyID
	End Function
End Class