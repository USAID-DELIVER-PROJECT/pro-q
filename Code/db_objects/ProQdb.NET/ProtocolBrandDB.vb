Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("ProtocolBrandDB_NET.ProtocolBrandDB")> Public Class ProtocolBrandDB
	'ProtocolBrandDB.cls
	'
	'this class manages the persistent data for the protocols.
	'
	'lbailey
	'2 june 2002
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         private members
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	'the name of the table or view used by this class
	Private Const m_strTable As String = DB_TABLE_PROTOCOL_BRAND
    Private Const m_strSQL As String = "Select * FROM " & DB_TABLE_PROTOCOL_BRAND
    'whether or not the object has been constructed
    Private m_objConn As OleDbConnection
    Private m_objRS As DataSet
    Dim m_objAdapter As OleDbDataAdapter
	Private m_fIsNew As Boolean
	
	'the guid of the object
	Private m_strID As String
	'the friendly name
	Private m_strName As String
	'free text
	Private m_strNotes As String
	'the brand used
	Private m_strBrandID As String
	'the test that this belongs to
	Private m_strTestID As String
	'the kit that has been associtated with this brand
	Private m_strKitID As String
	'the number of tests of this brand
	Private m_lCount As Object
	'the percentage of the tests that this brand is used
	Private m_sngPercent As Object
	'whether the mfr/supplier information should be displayed
	Private m_bGeneric As Object
	Private m_strGenericCode As String
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         public methods
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	
	'+
	'Create()
	'
	'creates a new object, but doesn't write it to the db
	'
	'lbailey
	'16 may 2002
	'-
	Public Function Create() As String
		
		'UPGRADE_WARNING: Couldn't resolve default property of object g_objGUIDGenerator.GetGUID(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_strID = g_objGUIDGenerator.GetGUID()
		'UPGRADE_WARNING: Couldn't resolve default property of object m_lCount. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_lCount = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object m_sngPercent. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_sngPercent = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object m_bGeneric. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_bGeneric = False
		
		m_fIsNew = True
		
		'return the id of the created object
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
	Public Function Load(ByVal strID As String) As Object
		
        Dim i As Integer
        Dim strGuid As String

        m_objRS = OpenDB(m_objConn, m_objRS, m_strTable, m_strSQL & " WHERE GuidID = {" & strID & "}")

        'select the desired record
        'm_objRS.Filter = DB_TABLE_PK & " = '" & strID & "'"
        For i = 0 To m_objRS.Tables(m_strTable).Rows.Count - 1
            strGuid = CType((m_objRS.Tables(m_strTable).Rows(i).Item(DB_TABLE_PK)), Guid).ToString()
            If strGuid = strID Then
                With m_objRS.Tables(m_strTable).Rows(i)
                    'copy the values into the class members
                    m_strID = CType((.Item("GuidID")), Guid).ToString() 'm_objRS.Fields("guidID").Value
                    m_strName = .Item("txtName") 'm_objRS.Fields("txtName").Value
                    m_strNotes = IIf(IsDBNull(.Item("memNotes")), "", .Item("memNotes"))
                    m_strBrandID = CType((.Item("guidID_Brand")), Guid).ToString() 'm_objRS.Fields("guidID_Brand").Value
                    m_strTestID = CType((.Item("guidID_ProtocolTest")), Guid).ToString() 'm_objRS.Fields("guidID_ProtocolTest").Value
                    m_strKitID = CType((.Item("guidID_Kit")), Guid).ToString() 'm_objRS.Fields("guidID_Kit").Value
                    m_lCount = .Item("lngCount") 'm_objRS.Fields("lngCount").Value
                    m_sngPercent = .Item("sngPercent") 'm_objRS.Fields("sngPercent").Value
                    m_bGeneric = .Item("fGeneric") 'm_objRS.Fields("fGeneric").Value
                    m_strGenericCode = IIf(IsDBNull(.Item("strGenericCode")), "", .Item("strGenericCode")) 'm_objRS.Fields("strGenericCode").Value
                End With

                Exit For
            End If
        Next i

        CloseDB(m_objConn, m_objRS)
        m_fIsNew = False

        'return result code
        'UPGRADE_WARNING: Couldn't resolve default property of object Load. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Load = S_OK

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

        If m_fIsNew Then

            dsNewRow = m_objRS.Tables(m_strTable).NewRow()

            'stuff all of the record's properties
            dsNewRow.Item("guidID") = New Guid(m_strID)
            dsNewRow.Item("txtName") = m_strName
            dsNewRow.Item("memNotes") = m_strNotes
            dsNewRow.Item("guidID_Brand") = New Guid(m_strBrandID)
            dsNewRow.Item("guidID_ProtocolTest") = New Guid(m_strTestID)
            dsNewRow.Item("guidID_Kit") = New Guid(m_strKitID)
            dsNewRow.Item("lngCount") = m_lCount
            dsNewRow.Item("sngPercent") = m_sngPercent
            dsNewRow.Item("fGeneric") = m_bGeneric
            dsNewRow.Item("strGenericCode") = IIf(IsDBNull(m_strGenericCode), "", m_strGenericCode)
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
                        .Tables(m_strTable).Rows(i).Item("guidID") = New Guid(m_strID)
                        .Tables(m_strTable).Rows(i).Item("txtName") = m_strName
                        .Tables(m_strTable).Rows(i).Item("memNotes") = m_strNotes
                        .Tables(m_strTable).Rows(i).Item("guidID_Brand") = New Guid(m_strBrandID)
                        .Tables(m_strTable).Rows(i).Item("guidID_ProtocolTest") = New Guid(m_strTestID)
                        .Tables(m_strTable).Rows(i).Item("guidID_Kit") = New Guid(m_strKitID)
                        .Tables(m_strTable).Rows(i).Item("lngCount") = m_lCount
                        .Tables(m_strTable).Rows(i).Item("sngPercent") = m_sngPercent
                        .Tables(m_strTable).Rows(i).Item("fGeneric") = m_bGeneric
                        .Tables(m_strTable).Rows(i).Item("strGenericCode") = m_strGenericCode
                        'write the record
                        m_objAdapter.Update(m_objRS, m_strTable) '.Update()
                        Exit For
                    End If
                Next
            End With
        End If

        'clean up
        m_fIsNew = False

        CloseDB(m_objConn, m_objRS)

    End Sub



    '+
    'Delete()
    '
    'Removes this object from the db
    '
    'lbailey 5 june 2002
    '-
    Public Sub Delete()

        If Not m_fIsNew Then
            'OpenDB(m_objConn, m_objRS, m_strTable, m_strSQL)
            DeleteRecord(m_objConn, m_objRS, m_strTable, m_strSQL, m_strID)
            CloseDB(m_objConn, m_objRS)
        End If

    End Sub
	
	
	
	'+
	'standard accessor functions
	'
	'lbailey
	'5 june 2002
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
	
	Public Function GetBrandID() As String
		GetBrandID = m_strBrandID
	End Function
	
	Public Function GetTestID() As String
		GetTestID = m_strTestID
	End Function
	
	Public Function GetKitID() As String
		GetKitID = m_strKitID
	End Function
	
	Public Function GetCount() As Integer
		'UPGRADE_WARNING: Couldn't resolve default property of object m_lCount. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetCount = m_lCount
	End Function
	
	Public Function GetPercent() As Single
		'UPGRADE_WARNING: Couldn't resolve default property of object m_sngPercent. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetPercent = m_sngPercent
	End Function
	
	Public Function GetGeneric() As Boolean
		'UPGRADE_WARNING: Couldn't resolve default property of object m_bGeneric. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetGeneric = m_bGeneric
	End Function
	Public Function GetGenericCode() As String
		GetGenericCode = m_strGenericCode
	End Function
	
	
	
	'+
	'standard manipulator functions
	'
	'lbailey
	'5 june 2002
	'-
	Public Sub SetName(ByRef strName As String)
		m_strName = strName
	End Sub
	
	Public Sub SetNotes(ByRef strNotes As String)
		m_strNotes = strNotes
	End Sub
	
	Public Sub SetBrandID(ByRef strBrandID As String)
		m_strBrandID = strBrandID
	End Sub
	
	Public Sub SetTestID(ByRef strTestID As String)
		m_strTestID = strTestID
	End Sub
	
	Public Sub SetKitID(ByRef strKitID As String)
		m_strKitID = strKitID
	End Sub
	
	Public Sub SetCount(ByRef lCount As Integer)
		'UPGRADE_WARNING: Couldn't resolve default property of object m_lCount. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_lCount = lCount
	End Sub
	
	Public Sub setPercent(ByRef sngPercent As Single)
		'UPGRADE_WARNING: Couldn't resolve default property of object m_sngPercent. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_sngPercent = sngPercent
	End Sub
	
	Public Sub SetGeneric(ByRef bGeneric As Boolean)
		'UPGRADE_WARNING: Couldn't resolve default property of object m_bGeneric. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_bGeneric = bGeneric
	End Sub
	
	Public Sub SetGenericCode(ByRef strCode As String)
		m_strGenericCode = strCode
	End Sub
End Class