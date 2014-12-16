Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("CartonDB_NET.CartonDB")> Public Class CartonDB
	'CartonDB.cls
	'Manages the cartons in the db
	'31-May-2002 lblanken
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         private members
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	'the name of the table or view used by this class
    Private Const m_strTable As String = DB_TABLE_CARTON
    Private Const m_strSQL As String = "Select * From " & DB_TABLE_CARTON

	'whether or not the object has been constructed
    Private m_objConn As OleDbConnection
    Private m_objAdapter As OleDbDataAdapter
    Private m_objRS As dataset
	Private m_fIsNew As Boolean
	
	'the guid of the object
	Private m_strID As String
	'the brand that makes up this kit
    Private m_strKitID As String
	'the number of total? or usable? tests per kit?
	Private m_nKitsPerCarton As Object
	'carton dimensions
    Private m_dblLength As Double
    Private m_dblWidth As Double
    Private m_dblHeight As Double
    Private m_dblWeight As Double
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         public methods
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	Public Sub Load(ByVal strID As String)
		' Comments   : gets the specified record and populates this object with its values
		' Parameters : strID - GUID to load
		' Returns    :
		' Modified   : 31-May-2002 LKB
		' --------------------------------------------------------
        Dim strGuid As String
        Dim i As Integer
		'UPGRADE_WARNING: Couldn't resolve default property of object m_strTable. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        m_objRS = OpenDB(m_objConn, m_objRS, m_strTable, m_strSQL)
		
		'select the desired record
        'm_objRS.Filter = DB_TABLE_PK & " = '" & strID & "'"
        For i = 0 To m_objRS.Tables(m_strTable).Rows.Count - 1
            strGuid = CType((m_objRS.Tables(m_strTable).Rows(i).Item("GuidID")), Guid).ToString()
            If strGuid = strID Then
                With m_objRS.Tables(m_strTable).Rows(i)
                    'copy the values into the class members
                    m_strID = CType((.Item("GuidID")), Guid).ToString() '.Fields("guidID").Value
                    m_strKitID = CType((.Item("guidID_Kit")), Guid).ToString() 'm_objRS.Fields("guidID_Kit").Value
                    m_nKitsPerCarton = .Item("intKitsPerCarton") 'm_objRS.Fields("intKitsPerCarton").Value
                    m_dblWeight = .Item("dblWeight") 'm_objRS.Fields("dblWeight").Value
                    m_dblHeight = .Item("dbldepth") 'm_objRS.Fields("dbldepth").Value
                    m_dblLength = .Item("dblLength") 'm_objRS.Fields("dblLength").Value
                    m_dblWidth = .Item("dblWidth") 'm_objRS.Fields("dblWidth").Value
                End With
                Exit For
            End If
        Next i

        CloseDB(m_objConn, m_objRS)

	End Sub
	
	Public Function Create() As String
		' Comments   : Adds the New Record to the db using the info that we've got
		' Parameters :  -
		' Returns    :  -
		' Modified   : 31-May-2002 LKB
		' --------------------------------------------------------
		
		'create the id and set the defaults
		'UPGRADE_WARNING: Couldn't resolve default property of object g_objGUIDGenerator.GetGUID(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_strID = g_objGUIDGenerator.GetGUID()
		
		'remember that this is a new record
		m_fIsNew = True
		
		Create = m_strID
		
	End Function
	
	Public Sub Update()
		' Comments   : Updates the Record to the db using the info that we've got
		' Parameters : -
		' Returns    : -
		' Modified   : 31-May-2002 LKB
		' --------------------------------------------------------
		
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
            dsNewRow.Item("guidID_Kit") = New Guid(m_strKitID)
            dsNewRow.Item("intKitsPerCarton") = m_nKitsPerCarton
            dsNewRow.Item("dblWeight") = m_dblWeight
            dsNewRow.Item("dbldepth") = m_dblHeight
            dsNewRow.Item("dblLength") = m_dblLength
            dsNewRow.Item("dblWidth") = m_dblWidth
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
                        .Tables(m_strTable).Rows(i).Item("guidID_Kit") = New Guid(m_strKitID)
                        .Tables(m_strTable).Rows(i).Item("intKitsPerCarton") = m_nKitsPerCarton
                        .Tables(m_strTable).Rows(i).Item("dblWeight") = m_dblWeight
                        .Tables(m_strTable).Rows(i).Item("dbldepth") = m_dblHeight
                        .Tables(m_strTable).Rows(i).Item("dblLength") = m_dblLength
                        .Tables(m_strTable).Rows(i).Item("dblWidth") = m_dblWidth
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

    Public Sub Delete()
        ' Comments   : Removes the specified object from the repository
        ' Parameters : lngID - ID to delete
        ' Returns    :  -
        ' Modified   : 31-May-2002 LKB
        ' --------------------------------------------------------

        'UPGRADE_WARNING: Couldn't resolve default property of object m_strTable. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'OpenDB(m_objConn, m_objRS, m_strTable, m_strSQL)
        DeleteRecord(m_objConn, m_objRS, m_strTable, m_strSQL, m_strID)
        CloseDB(m_objConn, m_objRS)

    End Sub

    '-----------------------------------------------------------------
    'Standard Accessor Functions
    ' Modified   : 31-May-2002 LKB
    '-----------------------------------------------------------------

    Public Function GetID() As String
        GetID = m_strID
    End Function

    Public Function GetKitID() As String
        'UPGRADE_WARNING: Couldn't resolve default property of object m_strKitID. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        GetKitID = m_strKitID
    End Function

    Public Function GetKitsPerCarton() As Short
        'UPGRADE_WARNING: Couldn't resolve default property of object m_nKitsPerCarton. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        GetKitsPerCarton = m_nKitsPerCarton
    End Function

    Public Function GetWeight() As Double
        'UPGRADE_WARNING: Couldn't resolve default property of object m_dblWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        GetWeight = m_dblWeight
    End Function

    Public Function GetHeight() As Double
        'UPGRADE_WARNING: Couldn't resolve default property of object m_dblHeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        GetHeight = m_dblHeight
    End Function

    Public Function GetLength() As Double
        'UPGRADE_WARNING: Couldn't resolve default property of object m_dblLength. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        GetLength = m_dblLength
    End Function

    Public Function GetWidth() As Double
        'UPGRADE_WARNING: Couldn't resolve default property of object m_dblWidth. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        GetWidth = m_dblWidth
    End Function

    '-----------------------------------------------------------------
    'Standard Manipulator Functions
    ' Modified   : 31-May-2002 LKB
    '-----------------------------------------------------------------
    Public Sub SetID(ByRef strID As String)
        m_strID = strID
    End Sub

    Public Sub SetKitID(ByRef strKitID As String)
        'UPGRADE_WARNING: Couldn't resolve default property of object m_strKitID. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        m_strKitID = strKitID
    End Sub

    Public Sub SetKitsPerCarton(ByRef nKitsPerCarton As Short)
        'UPGRADE_WARNING: Couldn't resolve default property of object m_nKitsPerCarton. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        m_nKitsPerCarton = nKitsPerCarton
    End Sub

    Public Sub SetWeight(ByRef dblWeight As Double)
        'UPGRADE_WARNING: Couldn't resolve default property of object m_dblWeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        m_dblWeight = dblWeight
    End Sub

    Public Sub SetHeight(ByRef dblHeight As Double)
        'UPGRADE_WARNING: Couldn't resolve default property of object m_dblHeight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        m_dblHeight = dblHeight
    End Sub

    Public Sub SetLength(ByRef dblLength As Double)
        'UPGRADE_WARNING: Couldn't resolve default property of object m_dblLength. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        m_dblLength = dblLength
    End Sub

    Public Sub SetWidth(ByRef dblWidth As Double)
        'UPGRADE_WARNING: Couldn't resolve default property of object m_dblWidth. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        m_dblWidth = dblWidth
    End Sub
End Class