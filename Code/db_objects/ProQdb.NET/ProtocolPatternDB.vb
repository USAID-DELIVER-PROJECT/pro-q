Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("ProtocolPatternDB_NET.ProtocolPatternDB")> Public Class ProtocolPatternDB
	'ProtocolPatternDB.cls
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
	Private Const m_strTable As String = DB_TABLE_PROTOCOL_PATTERN
    Private Const m_strSQL As String = "Select * FROM " & DB_TABLE_PROTOCOL_PATTERN
    'whether or not the object has been constructed
    Private m_objConn As OleDbConnection
    Private m_objRS As DataSet
    Dim m_objAdapter As OleDbDataAdapter
	Private m_fIsNew As Boolean
	
	'the guid of the object
	Private m_strID As String
	'the friendly name
	Private m_strName As String
	'free text describing the recod
	Private m_strNotes As String
	'a reference to the graphical representation of the prot
	Private m_strDiagram As String
	
	
	
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
        'get a connection to the db
        m_objConn = GetConnection(DB_DSN)

        'load the data into memory
        'rst.Open(m_strTable, m_objConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockPessimistic, ADODB.CommandTypeEnum.adCmdTable)
        m_objAdapter = New OleDb.OleDbDataAdapter(m_strSQL & " WHERE GuidID = {" & strID & "}", m_objConn)
        m_objRS = New DataSet
        m_objAdapter.Fill(m_objRS, m_strTable)

        'select the desired record
        '.Filter = "guidID = '" & strID & "'"
        For i = 0 To m_objRS.Tables(m_strTable).Rows.Count - 1
            strGuid = CType((m_objRS.Tables(m_strTable).Rows(i).Item("GuidID")), Guid).ToString()
            If strGuid = strID Then
                With m_objRS.Tables(m_strTable).Rows(i)
                    'copy the values into the class members
                    m_strID = CType((.Item("GuidID")), Guid).ToString() '.Fields("guidID").Value
                    m_strName = .Item("strName") '.Fields("strName").Value
                    m_strNotes = .Item("MemNotes") '.Fields("MemNotes").Value
                    m_strDiagram = .Item("strDiagram") '.Fields("strDiagram").Value
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
            dsNewRow.Item("strName") = m_strName
            dsNewRow.Item("MemNotes") = m_strNotes
            dsNewRow.Item("strDiagram") = m_strDiagram
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
                        .Tables(m_strTable).Rows(i).Item("GuidID") = New Guid(m_strID) '.Fields("GuidID").Value = m_strID
                        .Tables(m_strTable).Rows(i).Item("strName") = m_strName
                        .Tables(m_strTable).Rows(i).Item("MemNotes") = m_strNotes
                        .Tables(m_strTable).Rows(i).Item("strDiagram") = m_strDiagram
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
	
	Public Function GetName() As String
		GetName = m_strName
	End Function
	
	Public Function GetNotes() As String
		GetNotes = m_strNotes
	End Function
	
	Public Function GetDiagram() As String
		GetDiagram = m_strDiagram
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
	
	Public Sub SetDiagram(ByRef strDiagram As String)
		m_strDiagram = strDiagram
	End Sub
End Class