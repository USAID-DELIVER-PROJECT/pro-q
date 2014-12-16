Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("FundingDB_NET.FundingDB")> Public Class FundingDB
	'FundingDB.cls
	'
	'this class manages the persistent data for the Quantification Funding
	'
	'JLeiner
	'07 June 2002
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                            constants
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         private members
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	'the dsn used to connect to this db/table
    Private m_strDSN As String = DB_DSN
	'the name of the table or view used by this class
    Private Const m_strTable As String = DB_TABLE_FUNDING
    Private Const m_strSQL As String = "Select * From " & DB_TABLE_FUNDING

	'whether or not the object has been constructed
    Private m_objConn As OleDbConnection
    'Private m_objRS As DataSet 

    Private m_strID As String 'the guid of the object
    Private m_strFundingSourceID As String 'The Funding Source
    Private m_strQuantificationID As String 'The QuantificationID
    Private m_dblValue As Double 'The Currency amount allocated to Quantification
    Private m_lngCategory As Integer 'The Category of the funding (1=Kit,2=Customs,3=storage)
    Private m_strNotes As String 'free text describing the record


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

        Dim rst As New DataSet
        Dim strGuid As String
        Dim i As Integer

        'm_objConn = GetConnection(m_strDSN)
        'rst.Open(m_strTable, m_objConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockPessimistic, ADODB.CommandTypeEnum.adCmdTable)
        rst = OpenDB(m_objConn, rst, m_strTable, m_strSQL & " WHERE GuidID = {" & strID & "}")
        'select the desired record
        '.Filter = DB_TABLE_PK & " = '" & strID & "'"
        For i = 0 To rst.Tables(m_strTable).Rows.Count - 1
            strGuid = CType((rst.Tables(m_strTable).Rows(i).Item(DB_TABLE_PK)), Guid).ToString()
            If strGuid = strID Then
                With rst.Tables(m_strTable).Rows(i)
                    'copy the values into the class members
                    m_strID = CType((.Item("GuidID")), Guid).ToString() '.Fields("guidID").Value
                    m_strFundingSourceID = CType((.Item("guidID_FundingSource")), Guid).ToString() '.Fields("guidID_FundingSource").Value
                    m_strQuantificationID = CType((.Item("guidID_Quantification")), Guid).ToString() '.Fields("guidID_Quantification").Value
                    m_dblValue = .Item("dblValue") '.Fields("dblValue").Value
                    m_lngCategory = .Item("lngCategory") '.Fields("lngCategory").Value
                End With
                Exit For
            End If
        Next

        CloseDB(m_objConn, rst)

        'UPGRADE_NOTE: Object rst may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        rst = Nothing
        'UPGRADE_NOTE: Object m_objConn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        m_objConn = Nothing
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

        'get a connection to the db
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
        dsNewRow.Item("guidID_FundingSource") = New Guid(m_strFundingSourceID)
        dsNewRow.Item("guidID_Quantification") = New Guid(m_strQuantificationID)
        dsNewRow.Item("dblValue") = m_dblValue
        dsNewRow.Item("lngCategory") = m_lngCategory
        rst.Tables(m_strTable).Rows.Add(dsNewRow)

        objAdapter.Update(rst, m_strTable)

        'clean up
        CloseDB(m_objConn, rst)

        'return the id of the created object
        Create = m_strID '.Fields("guidID").Value()
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
                    .Tables(m_strTable).Rows(i).Item("guidID") = New Guid(m_strID)
                    .Tables(m_strTable).Rows(i).Item("guidID_FundingSource") = New Guid(m_strFundingSourceID)
                    .Tables(m_strTable).Rows(i).Item("guidID_Quantification") = New Guid(m_strQuantificationID)
                    .Tables(m_strTable).Rows(i).Item("dblValue") = m_dblValue
                    .Tables(m_strTable).Rows(i).Item("lngCategory") = m_lngCategory
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
	
	Public Function GetFundingSourceID() As String
		GetFundingSourceID = m_strFundingSourceID
	End Function
	
	Public Function getQuantificationID() As String
		getQuantificationID = m_strQuantificationID
	End Function
	
	Public Function GetValue() As Double
		GetValue = m_dblValue
	End Function
	
	Public Function GetCategory() As Integer
		GetCategory = m_lngCategory
	End Function
	
	Public Function GetNotes() As String
		GetNotes = m_strNotes
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
	Public Sub SetFundingSourceID(ByRef strFundingSourceID As String)
		m_strFundingSourceID = strFundingSourceID
	End Sub
	
	Public Sub setQuantificationID(ByRef strQuantificationID As String)
		m_strQuantificationID = strQuantificationID
	End Sub
	
	Public Sub SetValue(ByRef dblValue As Double)
		m_dblValue = dblValue
	End Sub
	Public Sub SetCategory(ByRef lngCategory As Integer)
		m_lngCategory = lngCategory
	End Sub
	
	Public Sub SetNotes(ByRef strNotes As String)
		m_strNotes = strNotes
	End Sub
End Class