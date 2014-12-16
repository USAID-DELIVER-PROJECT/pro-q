Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("ProtocolPatternLevelDB_NET.ProtocolPatternLevelDB")> Public Class ProtocolPatternLevelDB
	'ProtocolPatternLevelDB.cls
	'
	'this class manages the lookup data for the protocol levels
	'
	'lbailey
	'2 june 2002
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         private members
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	'the name of the table or view used by this class
	Private Const m_strTable As String = DB_TABLE_PROTOCOL_PATTERN_LEVEL
    Private Const m_strSQL As String = "Select * FROM " & DB_TABLE_PROTOCOL_PATTERN_LEVEL

	'whether or not the object has been constructed
    Private m_objConn As OleDbConnection
    Private m_objRS As DataSet 'ADODB.Recordset
    Private m_objAdapter As OleDbDataAdapter
	Private m_fIsNew As Boolean
	
	'the guid of the object
	Private m_strID As String
	'the id of the protocol pattern (parent table)
	Private m_strPatternID As String
	'which level is this?
	Private m_bytLevel As Byte
	Private m_bytRefLevel As Byte
	'the friendly name
	Private m_strName As String
	'the default percentage for this level
	Private m_sngPercent As Single
	Private m_intPercentStructure As Short
	
	
	
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
	'16 may 2002
	'-
	Public Sub Load(ByVal strID As String)
		
        Dim i As Integer
        Dim strGuid As String
        'UPGRADE_WARNING: Couldn't resolve default property of object m_strTable. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        m_objRS = OpenDB(m_objConn, m_objRS, m_strTable, m_strSQL & " WHERE GuidID = {" & strID & "}")

        'select the desired record
        'm_objRS.Filter = DB_TABLE_PK & " = '" & strID & "'"
        For i = 0 To m_objRS.Tables(m_strTable).Rows.Count - 1
            strGuid = CType((m_objRS.Tables(m_strTable).Rows(i).Item("GuidID")), Guid).ToString()
            If strGuid = strID Then
                With m_objRS.Tables(m_strTable).Rows(i)
                    'copy the values into the class members
                    m_strID = CType((.Item("GuidID")), Guid).ToString() 'm_objRS.Fields("guidID").Value
                    m_strPatternID = CType((.Item("guidID_ProtocolPattern")), Guid).ToString() 'm_objRS.Fields("guidID_ProtocolPattern").Value
                    m_bytLevel = .Item("bytLevel") 'm_objRS.Fields("bytLevel").Value
                    m_bytRefLevel = .Item("bytRefLevel") 'm_objRS.Fields("bytRefLevel").Value
                    m_strName = .Item("strName") 'm_objRS.Fields("strName").Value
                    m_sngPercent = .Item("sngPercent") 'm_objRS.Fields("sngPercent").Value
                    m_intPercentStructure = .Item("intPercentStructure") 'm_objRS.Fields("intPercentStructure").Value
                End With
                CloseDB(m_objConn, m_objRS)
                Exit For
            End If
        Next i
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


        'UPGRADE_WARNING: Couldn't resolve default property of object g_objGUIDGenerator.GetGUID(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        m_strID = g_objGUIDGenerator.GetGUID()
        m_sngPercent = 0

        m_fIsNew = True

        Create = m_strID

    End Function



    '+
    'Update()
    '
    'Persists the current fields into the db
    '
    'jleiner/lbailey 31 may 2002
    '-
    Public Sub Update()

        'UPGRADE_WARNING: Couldn't resolve default property of object m_strTable. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'm_objRS = OpenDB(m_objConn, m_objRS, m_strTable, m_strSQL)

        'InsertOrUpdate(m_objRS, m_fIsNew, m_strID)

        'With m_objRS
        'stuff all of the record's properties
        '.Fields("guidID").Value = m_strID
        '.Fields("guidID_ProtocolPattern").Value = m_strPatternID
        '.Fields("bytLevel").Value = m_bytLevel
        '.Fields("bytRefLevel").Value = m_bytRefLevel
        '.Fields("strName").Value = m_strName
        '.Fields("sngPercent").Value = m_sngPercent
        '.Fields("intPercentStructure").Value = m_intPercentStructure
        'write the record
        '.Update()
        'End With

        'CloseDB(m_objConn, m_objRS)

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
            dsNewRow.Item("GuidID") = New Guid(m_strID) '.Fields("guidID").Value = m_strID
            dsNewRow.Item("guidID_ProtocolPattern") = New Guid(m_strPatternID) '.Fields("guidID_ProtocolPattern").Value = m_strPatternID
            dsNewRow.Item("bytLevel") = m_bytLevel '.Fields("bytLevel").Value = m_bytLevel
            dsNewRow.Item("bytRefLevel") = m_bytRefLevel '.Fields("bytRefLevel").Value = m_bytRefLevel
            dsNewRow.Item("strName") = m_strName '.Fields("strName").Value = m_strName
            dsNewRow.Item("sngPercent") = m_sngPercent '.Fields("sngPercent").Value = m_sngPercent
            dsNewRow.Item("intPercentStructure") = m_intPercentStructure '.Fields("intPercentStructure").Value = m_intPercentStructure
            m_objRS.Tables(m_strTable).Rows.Add(dsNewRow)

            m_objAdapter.Update(m_objRS, m_strTable)
            m_fIsNew = False
        Else
            With m_objRS
                For i = 0 To .Tables(m_strTable).Rows.Count - 1
                    Dim strGuid As String
                    strGuid = CType(.Tables(m_strTable).Rows(i).Item(DB_TABLE_PK), Guid).ToString
                    If strGuid = m_strID Then
                        'stuff all of the record's properties
                        .Tables(m_strTable).Rows(i).Item("GuidID") = New Guid(m_strID)
                        .Tables(m_strTable).Rows(i).Item("guidID_ProtocolPattern") = New Guid(m_strPatternID)
                        .Tables(m_strTable).Rows(i).Item("bytLevel") = m_bytLevel '.Fields("bytLevel").Value = m_bytLevel
                        .Tables(m_strTable).Rows(i).Item("bytRefLevel") = m_bytRefLevel '.Fields("bytRefLevel").Value = m_bytRefLevel
                        .Tables(m_strTable).Rows(i).Item("strName") = m_strName '.Fields("strName").Value = m_strName
                        .Tables(m_strTable).Rows(i).Item("sngPercent") = m_sngPercent '.Fields("sngPercent").Value = m_sngPercent
                        .Tables(m_strTable).Rows(i).Item("intPercentStructure") = m_intPercentStructure '.Fields("intPercentStructure").Value = m_intPercentStructure
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
    'Removes the current object from the repository
    '
    'lbailey 16 may 2002
    '-
    Public Sub Delete()

        'UPGRADE_WARNING: Couldn't resolve default property of object m_strTable. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'OpenDB(m_objConn, m_objRS, m_strTable, m_strSQL)
        DeleteRecord(m_objConn, m_objRS, m_strTable, m_strSQL, m_strID)
        CloseDB(m_objConn, m_objRS)

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

    Public Function GetPatternID() As String
        GetPatternID = m_strPatternID
    End Function

    Public Function GetLevel() As Byte
        GetLevel = m_bytLevel
    End Function

    Public Function GetRefLevel() As Byte
        GetRefLevel = m_bytRefLevel
    End Function


    Public Function GetName() As String
        GetName = m_strName
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
    '16 may 2002
    '-
    Public Sub SetID(ByRef strID As String)
        m_strID = strID
    End Sub

    Public Sub SetPatternID(ByRef strPatternID As String)
        m_strPatternID = strPatternID
    End Sub

    Public Sub SetLevel(ByRef bytLevel As Byte)
        m_bytLevel = bytLevel
    End Sub
    Public Sub SetRefLevel(ByRef bytRefLevel As Byte)
        m_bytRefLevel = bytRefLevel
    End Sub

    Public Sub SetName(ByRef strName As String)
        m_strName = strName
    End Sub

    Public Sub setPercent(ByRef sngPercent As Single)
        m_sngPercent = sngPercent
    End Sub

    Public Sub SetPercentStructure(ByRef intPercentStructure As Short)
        'set the member
        m_intPercentStructure = intPercentStructure
    End Sub
End Class