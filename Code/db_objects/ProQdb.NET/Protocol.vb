Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("ProtocolDB_NET.ProtocolDB")> Public Class ProtocolDB
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
	
	'the dsn used to connect to this db/table
    Private m_strDSN As String = DB_DSN
	'the name of the table or view used by this class
	Private Const m_strTable As String = DB_TABLE_PROTOCOL
	
	'whether or not the object has been constructed
    Private m_objConn As OleDbConnection
    Private m_objRS As DataSet
	Private m_fIsNew As Boolean
    Private m_objAdapter As OleDbDataAdapter '.Connection
    Private Const m_strSQL As String = "Select * From " & DB_TABLE_PROTOCOL

    'the guid of the object
	Private m_strID As String
	'the guid of the object
	Private m_strPatternID As String
	'the friendly name
	Private m_strName As String
	'free text describing the recod
	Private m_strNotes As String
	'a reference to the graphical representation of the prot
	Private m_strDiagram As String
	'the demand (the number of people's blood samples to be tested)
	Private m_lSamples As Integer
	
	
	
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
                    m_strID = CType((.Item("GuidID")), Guid).ToString() 'm_strID = m_objRS.Fields("guidID").Value
                    m_strPatternID = CType((.Item("guidID_Pattern")), Guid).ToString() 'm_strPatternID = m_objRS.Fields("guidID_Pattern").Value
                    m_strName = .Item("strName")   'm_strName = m_objRS.Fields("strName").Value
                    m_strNotes = IIf(IsDBNull(.Item("MemNotes")), "", .Item("MemNotes"))   'm_strNotes = IIf(IsDbNull(m_objRS.Fields("MemNotes").Value), "", m_objRS.Fields("MemNotes").Value)
                    m_strDiagram = .Item("strDiagram")   'm_strDiagram = m_objRS.Fields("strDiagram").Value
                    m_lSamples = .Item("lngSamples")   'm_lSamples = m_objRS.Fields("lngSamples").Value
                End With
                Exit For
            End If
        Next

        'remember that this record already exists
		m_fIsNew = False
		
		CloseDB(m_objConn, m_objRS)
		
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
		
		'create the id and set the defaults
		'UPGRADE_WARNING: Couldn't resolve default property of object g_objGUIDGenerator.GetGUID(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_strID = g_objGUIDGenerator.GetGUID()
		m_strName = "none"
		m_lSamples = 0
		
		'remember that this is a new record
		m_fIsNew = True
		
		Create = m_strID
		
	End Function
	
	
	
	'+
	'Update()
	'
	'Persists the current fields into the db, creating a new record, if
	'necessary
	'
	'jleiner/lbailey 31 may 2002
	'-
	Public Sub Update()
        Dim i As Integer
        Dim strGuid As String
        Dim dsNewRow As DataRow
        'connect to the db and get a recordset
        'OpenDB(m_objConn, m_objRS, m_strTable, m_strSQL)
        m_objConn = GetConnection(DB_DSN)
        m_objAdapter = New OleDb.OleDbDataAdapter(m_strSQL, m_objConn)
        m_objRS = New DataSet
        m_objAdapter.Fill(m_objRS, m_strTable)
        Dim cb As New OleDb.OleDbCommandBuilder(m_objAdapter)


        'ensure that we're pointing at the record that we want
        'InsertOrUpdate(m_objRS, m_fIsNew, m_strID)
        If m_fIsNew Then

            dsNewRow = m_objRS.Tables(m_strTable).NewRow()

            'stuff all of the record's properties
            dsNewRow.Item("guidID") = New Guid(m_strID) '.Fields("GuidID").Value = m_strID
            dsNewRow.Item("guidID_Pattern") = New Guid(m_strPatternID) '.Fields("strName").Value = m_strName
            dsNewRow.Item("strName") = m_strName '.Fields("strSupplier").Value = m_strSupplier
            dsNewRow.Item("MemNotes") = m_strNotes '.Fields("lngType").Value = m_lngType
            dsNewRow.Item("strDiagram") = m_strDiagram '.Fields("intShelfLife").Value = m_intShelfLife
            dsNewRow.Item("lngSamples") = m_lSamples '.Fields("strStorageTemp").Value = m_strStorageTemp
            m_objRS.Tables(m_strTable).Rows.Add(dsNewRow)

            m_objAdapter.Update(m_objRS, m_strTable)
            m_fIsNew = False
        Else
            For i = 0 To m_objRS.Tables(m_strTable).Rows.Count - 1
                strGuid = CType((m_objRS.Tables(m_strTable).Rows(i).Item("GuidID")), Guid).ToString()
                If strGuid = m_strID Then
                    With m_objRS
                        .Tables(m_strTable).Rows(i).Item("guidID") = New Guid(m_strID) '.Fields("guidID").Value = m_strID
                        .Tables(m_strTable).Rows(i).Item("guidID_Pattern") = New Guid(m_strPatternID) '.Fields("guidID_Pattern").Value = m_strPatternID
                        .Tables(m_strTable).Rows(i).Item("strName") = m_strName '.Fields("strName").Value = m_strName
                        .Tables(m_strTable).Rows(i).Item("MemNotes") = m_strNotes '.Fields("MemNotes").Value = m_strNotes
                        .Tables(m_strTable).Rows(i).Item("strDiagram") = m_strDiagram '.Fields("strDiagram").Value = m_strDiagram
                        .Tables(m_strTable).Rows(i).Item("lngSamples") = m_lSamples '.Fields("lngSamples").Value = m_lSamples

                        m_objAdapter.Update(m_objRS, m_strTable)
                        'With m_objRS
                        'stuff all of the record's properties
                        Exit For
                    End With
                End If
            Next
        End If


        'write the record
        '.Update()

        'End With

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
	
	Public Function GetPatternID() As String
		GetPatternID = m_strPatternID
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
	
	Public Function GetSamples() As Integer
		GetSamples = m_lSamples
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
	
	Public Sub SetName(ByRef strName As String)
		m_strName = strName
	End Sub
	
	Public Sub SetNotes(ByRef strNotes As String)
		m_strNotes = strNotes
	End Sub
	
	Public Sub SetDiagram(ByRef strDiagram As String)
		m_strDiagram = strDiagram
	End Sub
	
    Public Sub SetSamples(ByRef lSamples As Integer)
        m_lSamples = lSamples
    End Sub
End Class