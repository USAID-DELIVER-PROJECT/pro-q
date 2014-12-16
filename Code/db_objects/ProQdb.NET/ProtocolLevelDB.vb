Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("ProtocolLevelDB_NET.ProtocolLevelDB")> Public Class ProtocolLevelDB
	'ProtocolDB.cls
	'
	'this class manages the persistent data for the protocols.
	'
	'lbailey
	'31 may 2002
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         private members
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	'the name of the table or view used by this class
	Private Const m_strTable As String = DB_TABLE_PROTOCOL_LEVEL
    Private Const m_strSQL As String = "Select * FROM " & DB_TABLE_PROTOCOL_LEVEL
    'whether or not the object has been constructed
    Private m_objConn As OleDbConnection
    Private m_objRS As DataSet 'ADODB.Recordset
    Dim m_objAdapter As OleDbDataAdapter
	Private m_fIsNew As Boolean
	
	'the guid of the object
	Private m_strID As String
	'the id of the parent
	Private m_strProtocolID As String
	'which level this is
	Private m_nLevel As Short
	'which level this guy is taking its percentage from
	Private m_nRefLevel As Short
	'the percentage assigned to this level
	Private m_sngPercent As Single
	Private m_intPercentStructure As Short
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         public methods
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	
	'+
	'Create()
	'
	'creates a new object, but doesn't add it to the db
	'
	'lbailey
	'16 may 2002
	'-
	Public Function Create() As String
		
		'UPGRADE_WARNING: Couldn't resolve default property of object g_objGUIDGenerator.GetGUID(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        m_strID = g_objGUIDGenerator.GetGUID()
		m_sngPercent = 0
		m_fIsNew = True
		
		Create = m_strID
		
	End Function
	
	
	
	'+
	'Load()
	'
	'gets the specified record and populates this object with its values
	'
	'lbailey
	'16 may 2002
	'-
	Public Sub Load(ByRef strID As String)
		
        Dim i As Integer
        Dim strGuid As String

        m_objRS = OpenDB(m_objConn, m_objRS, m_strTable, m_strSQL & " WHERE GuidID = {" & strID & "}")

        'select the desired record
        'm_objRS.Filter = DB_TABLE_PK & " = '" & strID & "'"
        For i = 0 To m_objRS.Tables(m_strTable).Rows.Count - 1
            strGuid = CType((m_objRS.Tables(m_strTable).Rows(i).Item("GuidID")), Guid).ToString()
            If strGuid = strID Then
                With m_objRS.Tables(m_strTable).Rows(i)
                    'copy the values into the class members
                    m_strID = CType((.Item("GuidID")), Guid).ToString() 'm_objRS.Fields("guidID").Value
                    m_strProtocolID = CType((.Item("guidID_Protocol")), Guid).ToString() 'm_objRS.Fields("guidID_Protocol").Value
                    m_nLevel = .Item("bytLevel") 'm_objRS.Fields("bytLevel").Value
                    m_nRefLevel = .Item("bytRefLevel") 'm_objRS.Fields("bytRefLevel").Value
                    m_sngPercent = .Item("sngPercent") 'm_objRS.Fields("sngPercent").Value
                    m_intPercentStructure = .Item("intPercentStructure") 'm_objRS.Fields("intPercentStructure").Value

                End With

                CloseDB(m_objConn, m_objRS)
                Exit For
            End If
        Next

    End Sub



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
        Dim dsNewRow As DataRow
        'connect to the db and get a recordset
        'OpenDB(m_objConn, m_objRS, m_strTable, m_strSQL)
        m_objConn = GetConnection(DB_DSN)
        m_objAdapter = New OleDb.OleDbDataAdapter(m_strSQL, m_objConn)
        m_objRS = New DataSet
        m_objAdapter.Fill(m_objRS, m_strTable)
        cb = New OleDb.OleDbCommandBuilder(m_objAdapter)

        'ensure that we're pointing at the record that we want
        'InsertOrUpdate(m_objRS, m_fIsNew, m_strID)

        If m_fIsNew Then

            dsNewRow = m_objRS.Tables(m_strTable).NewRow()

            'stuff all of the record's properties
            dsNewRow.Item("GuidID") = New Guid(m_strID) '.Fields("GuidID").Value = m_strID
            dsNewRow.Item("guidID_Protocol") = New Guid(m_strProtocolID) '.Fields("guidID_Protocol").Value = m_strProtocolID
            dsNewRow.Item("bytLevel") = (m_nLevel) '.Fields("bytLevel").Value = m_nLevel
            dsNewRow.Item("bytRefLevel") = (m_nRefLevel) '.Fields("bytRefLevel").Value = m_nRefLevel
            dsNewRow.Item("sngPercent") = (m_sngPercent) '.Fields("sngPercent").Value = m_sngPercent
            dsNewRow.Item("intPercentStructure") = (m_intPercentStructure) '.Fields("intPercentStructure").Value = m_intPercentStructure
            m_objRS.Tables(m_strTable).Rows.Add(dsNewRow)

            m_objAdapter.Update(m_objRS, m_strTable)
            m_fIsNew = False
        Else
            With m_objRS
                For i = 0 To m_objRS.Tables(m_strTable).Rows.Count - 1
                    If m_objRS.Tables(m_strTable).Rows(i).Item(DB_TABLE_PK).ToString = m_strID Then
                        'stuff all of the record's properties
                        .Tables(m_strTable).Rows(i).Item("GuidID") = New Guid(m_strID) '.Fields("GuidID").Value = m_strID
                        .Tables(m_strTable).Rows(i).Item("guidID_Protocol") = New Guid(m_strProtocolID) '.Fields("guidID_Protocol").Value = m_strProtocolID
                        .Tables(m_strTable).Rows(i).Item("bytLevel") = (m_nLevel) '.Fields("bytLevel").Value = m_nLevel
                        .Tables(m_strTable).Rows(i).Item("bytRefLevel") = (m_nRefLevel) '.Fields("bytRefLevel").Value = m_nRefLevel
                        .Tables(m_strTable).Rows(i).Item("sngPercent") = (m_sngPercent) '.Fields("sngPercent").Value = m_sngPercent
                        .Tables(m_strTable).Rows(i).Item("intPercentStructure") = (m_intPercentStructure) '.Fields("intPercentStructure").Value = m_intPercentStructure
                        'write the record
                        m_objAdapter.Update(m_objRS, m_strTable) '.Update()

                        Exit For
                    End If
                Next
            End With
        End If

        'clean up
        CloseDB(m_objConn, m_objRS)
    End Sub



    '+
    'Delete()
    '
    'Removes the specified object from the repository
    '
    'lbailey
    '2 june 2002
    '-
    Public Sub Delete()

        'OpenDB(m_objConn, m_objRS, m_strTable, m_strSQL)
        DeleteRecord(m_objConn, m_objRS, m_strTable, m_strSQL, m_strID)
        CloseDB(m_objConn, m_objRS)

    End Sub



    '+
    'standard accessor functions
    '
    'lbailey
    '2 june 2002
    '-
    Public Function GetID() As String
        GetID = m_strID
    End Function

    Public Function GetProtocolID() As String
        GetProtocolID = m_strProtocolID
    End Function

    Public Function GetLevel() As Short
        GetLevel = m_nLevel
    End Function

    Public Function GetRefLevel() As Short
        GetRefLevel = m_nRefLevel
    End Function

    Public Function GetPercent() As Single
        GetPercent = m_sngPercent
    End Function
    Public Function GetPercentStructure() As Short
        GetPercentStructure = m_intPercentStructure
    End Function



    '+
    'standard manipulator functions
    '
    'lbailey
    '2 june 2002
    '-
    Public Sub SetID(ByRef strID As String)
        'set the member
        m_strID = strID
    End Sub

    Public Sub SetProtocolID(ByRef strProtocolID As String)
        'set the member
        m_strProtocolID = strProtocolID
    End Sub

    Public Sub SetLevel(ByRef nLevel As Short)
        'set the member
        m_nLevel = nLevel
    End Sub

    Public Sub SetRefLevel(ByRef nRefLevel As Short)
        'set the member
        m_nRefLevel = nRefLevel
    End Sub

    Public Sub setPercent(ByRef sngPercent As Single)
        'set the member
        m_sngPercent = sngPercent
    End Sub

    Public Sub SetPercentStructure(ByRef intPercentStructure As Short)
        'set the member
        m_intPercentStructure = intPercentStructure
    End Sub
End Class