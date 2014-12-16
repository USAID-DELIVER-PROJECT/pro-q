Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("ProtocolPatternTestDB_NET.ProtocolPatternTestDB")> Public Class ProtocolPatternTestDB
	'ProtocolPatternTestDB.cls
	'
	'this class manages the lookup data for the protocol tests
	'
	'lbailey
	'2 june 2002
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         private members
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	'the name of the table or view used by this class
	Private Const m_strTable As String = DB_TABLE_PROTOCOL_PATTERN_TEST
    Private Const m_strSQL As String = "Select * FROM " & DB_TABLE_PROTOCOL_PATTERN_TEST

	'whether or not the object has been constructed
    Private m_objConn As OleDbConnection
    Private m_objRS As DataSet
    Private m_objAdapter As OleDbDataAdapter
	Private m_fIsNew As Boolean
	
	'the guid of the object
	Private m_strID As String
	'the friendly name
	Private m_strName As String
	'the id of the protocol pattern (parent table)
	Private m_strLevelID As String
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         public methods
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
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
	Public Sub Load(ByVal strID As String)
		
        Dim strGuid As String
        Dim i As Integer

        m_objRS = OpenDB(m_objConn, m_objRS, m_strTable, m_strSQL & " WHERE GuidID = {" & strID & "}")

        'select the desired record

        For i = 0 To m_objRS.Tables(m_strTable).Rows.Count - 1
            strGuid = CType((m_objRS.Tables(m_strTable).Rows(i).Item("GuidID")), Guid).ToString()
            If strGuid = strID Then
                With m_objRS.Tables(m_strTable).Rows(i)
                    'copy the values into the class members
                    m_strID = CType((.Item("GuidID")), Guid).ToString()
                    m_strLevelID = CType((.Item("guidID_PatternLevel")), Guid).ToString()
                    m_strName = .Item("txtName")
                End With
                CloseDB(m_objConn, m_objRS)
                Exit For
            End If
        Next i
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
            dsNewRow.Item("GuidID") = New Guid(m_strID)
            dsNewRow.Item("guidID_PatternLevel") = New Guid(m_strLevelID)
            dsNewRow.Item("txtName") = m_strName
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
                        .Tables(m_strTable).Rows(i).Item("guidID_PatternLevel") = New Guid(m_strLevelID)
                        .Tables(m_strTable).Rows(i).Item("txtName") = m_strName
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
    'lbailey 16 may 2002
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
	'16 may 2002
	'-
	Public Function GetID() As String
		GetID = m_strID
	End Function
	
	Public Function GetLevelID() As String
		GetLevelID = m_strLevelID
	End Function
	
	Public Function GetName() As String
		GetName = m_strName
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
	
	Public Sub SetLevelID(ByRef strLevelID As String)
		m_strLevelID = strLevelID
	End Sub
	
	Public Sub SetName(ByRef strName As String)
		m_strName = strName
	End Sub
End Class