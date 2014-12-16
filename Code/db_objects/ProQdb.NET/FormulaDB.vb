Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("FormulaDB_NET.FormulaDB")> Public Class FormulaDB
	'FormulaDB.cls
	'Creates Sets the Formula Object
	'21-June-2002 jleiner
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                            constants
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Private Const DB_TABLE_FORMULA As String = "tlkFormula"
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         private members
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	'the dsn used to connect to this db/table
    Private m_strDSN As String = DB_DSN
	'the name of the table or view used by this class
    Private Const m_strTable As String = DB_TABLE_FORMULA
    Private Const m_strSQL As String = "Select * from " & DB_TABLE_FORMULA

	'whether or not the object has been constructed
    Private m_objConn As OleDbConnection
    'Private m_objRS As DataSet 

    Private m_strID As String 'the guid of the object
    Private m_strUseID As String 'The Use
    Private m_fElisa As Boolean 'Elisa or Not
    Private m_lngSection As Integer 'Formula Section
    Private m_strFormula As String 'The SQL formula for Demand

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '                         public members
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'hide everything from the user


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
    '
    'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Private Sub Class_Initialize_Renamed()
        'get a connection to the db
        'Set m_objConn = GetConnection(m_strDSN)
        'load the data into memory
        'm_objRS.Open _
        'm_strTable, _
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
    'jleiner
    '07 June 2002
    '
    'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Private Sub Class_Terminate_Renamed()
        'release the connections to the db
        '    m_objRS.Close
        '    Set m_objRS = Nothing
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
    'jleiner
    '07 June 2002
    '-
    Public Sub Load(ByRef strID As String)

        Dim rst As DataSet
        Dim i As Integer
        Dim strGuid As String
        Dim objAdapter As OleDbDataAdapter

        m_objConn = GetConnection(m_strDSN)

        'rst.Open(m_strTable, m_objConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockPessimistic, ADODB.CommandTypeEnum.adCmdTable)
        objAdapter = New OleDb.OleDbDataAdapter(m_strSQL & " WHERE GuidID = {" & strID & "}", m_objConn)
        rst = New DataSet
        objAdapter.Fill(rst, m_strTable)

        For i = 0 To rst.Tables(m_strTable).Rows.Count - 1
            'select the desired record
            '.Filter = DB_TABLE_PK & " = '" & strID & "'"
            strGuid = CType((rst.Tables(m_strTable).Rows(i).Item(DB_TABLE_PK)), Guid).ToString()
            If strGuid = strID Then
                With rst.Tables(m_strTable).Rows(i)
                    'copy the values into the class members
                    m_strID = CType((.Item("GuidID")), Guid).ToString() '.Fields("guidID").Value
                    m_strUseID = CType((.Item("guidID_Use")), Guid).ToString() '.Fields("guidID_Use").Value
                    m_lngSection = .Item("lngSection") '.Fields("lngSection").Value
                    m_fElisa = .Item("fElisa") '.Fields("fElisa").Value
                    m_strFormula = .Item("memFormula") '.Fields("memFormula").Value
                End With
                Exit For
            End If
        Next i

        CloseDB(m_objConn, rst)

    End Sub

    '+
    'Create()
    '
    'Adds the New Record to the db using the info that we've got
    '
    'lbailey
    '16 may 2002
    '-
    Public Function Create() As String

        Dim cb As OleDb.OleDbCommandBuilder
        Dim dsNewRow As DataRow
        Dim rst As New DataSet
        Dim objAdapter As OleDbDataAdapter
        m_objConn = GetConnection(m_strDSN)

        'connect to the db and get a recordset
        m_objConn = GetConnection(DB_DSN)
        objAdapter = New OleDb.OleDbDataAdapter(m_strSQL, m_objConn)
        objAdapter.Fill(rst, m_strTable)
        cb = New OleDb.OleDbCommandBuilder(objAdapter)

        dsNewRow = rst.Tables(m_strTable).NewRow()

        'if we don't already have an id for this, then create one
        If (Len(m_strID) < 1) Then
            'use the guid generator to create an id
            'UPGRADE_WARNING: Couldn't resolve default property of object g_objGUIDGenerator.GetGUID(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            m_strID = g_objGUIDGenerator.GetGUID()
        End If
        'stuff all of the record's properties
        dsNewRow.Item("guidID") = New Guid(m_strID)
        dsNewRow.Item("guidID_Use") = New Guid(m_strUseID)
        dsNewRow.Item("lngSection") = m_lngSection
        dsNewRow.Item("fElisa") = m_fElisa
        dsNewRow.Item("memFormula") = m_strFormula
        rst.Tables(m_strTable).Rows.Add(dsNewRow)

        objAdapter.Update(rst, m_strTable)

        'clean up
        CloseDB(m_objConn, rst)

        'return the id of the created object
        Create = m_strID

        'UPGRADE_NOTE: Object rst may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        rst = Nothing
        'UPGRADE_NOTE: Object m_objConn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        m_objConn = Nothing
    End Function


    '+
    'Update()
    '
    'Persists the current fields into the db
    '
    'jleiner/lbailey 31 may 2002
    '-
    Public Sub Update()

        Dim i As Integer
        Dim cb As OleDb.OleDbCommandBuilder
        Dim rst As New DataSet
        Dim objAdapter As OleDbDataAdapter

        'connect to the db and get a recordset
        m_objConn = GetConnection(DB_DSN)
        objAdapter = New OleDb.OleDbDataAdapter(m_strSQL, m_objConn)
        objAdapter.Fill(rst, m_strTable)
        cb = New OleDb.OleDbCommandBuilder(objAdapter)

        With rst
            For i = 0 To .Tables(m_strTable).Rows.Count - 1
                Dim strGuid As String
                strGuid = CType(.Tables(m_strTable).Rows(i).Item(DB_TABLE_PK), Guid).ToString
                If strGuid = m_strID Then
                    'stuff all of the record's properties
                    .Tables(m_strTable).Rows(i).Item("guidID_Use") = New Guid(m_strUseID)
                    .Tables(m_strTable).Rows(i).Item("lngSection") = m_lngSection
                    .Tables(m_strTable).Rows(i).Item("fElisa") = m_fElisa
                    .Tables(m_strTable).Rows(i).Item("memFormula") = m_strFormula
                    'write the record
                    objAdapter.Update(rst, m_strTable) '.Update()
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
    'Delete()
    '
    'Removes the specified object from the repository
    '
    'lbailey 16 may 2002
    '-
    Public Sub Delete()

        Dim rst As New DataSet

        DeleteRecord(m_objConn, rst, m_strTable, m_strSQL, m_strID)
        CloseDB(m_objConn, rst)

        'UPGRADE_NOTE: Object rst may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        rst = Nothing
        'UPGRADE_NOTE: Object m_objConn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        m_objConn = Nothing
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
	
	Public Function GetUseID() As String
		GetUseID = m_strUseID
	End Function
	
	Public Function getIsElisa() As Boolean
		getIsElisa = m_fElisa
	End Function
	Public Function GetSection() As Integer
		GetSection = m_lngSection
	End Function
	
	Public Function GetFormula() As String
		GetFormula = m_strFormula
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
	Public Sub SetUseID(ByRef strUseID As String)
		m_strUseID = strUseID
	End Sub
	Public Sub SetIsElisa(ByRef fElisa As Boolean)
		m_fElisa = fElisa
	End Sub
	Public Sub SetSection(ByRef lngSection As Integer)
		m_lngSection = lngSection
	End Sub
	
	Public Sub SetFormula(ByRef strFormula As String)
		m_strFormula = strFormula
	End Sub
End Class