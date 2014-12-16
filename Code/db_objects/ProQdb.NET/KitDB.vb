Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("KitDB_NET.KitDB")> Public Class KitDB
	'KitDB.cls
	'
	'Manages the kits in the db
	'
	'lbailey
	'1 june 2002
	'1 July 2002 - jleiner - added Dimensions fields for calculating volume
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         private members
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	'the name of the table or view used by this class
	Private Const m_strTable As String = DB_TABLE_KIT
    Private Const m_strSQL As String = "Select * FROM " & DB_TABLE_KIT

    'whether or not the object has been constructed
    Private m_objConn As OleDbConnection
    Private m_objRS As DataSet 'ADODB.Recordset
    Private m_objAdapter As OleDbDataAdapter
	Private m_fIsNew As Boolean
	
	'the guid of the object
	Private m_strID As String
	'the brand that makes up this kit
	Private m_strBrandID As String
	'the number of total? or usable? tests per kit?
	Private m_nTestsPerKit As Short
	
	'in dollars
	Private m_dblKitCost As Double
	
	'length x width x height
	Private m_strKitDimension As Object
	
	'in lbs? or kg?
	Private m_dblWeight As Double
	
	' Dimensions
	Private m_dblHeight As Double
	Private m_dblWidth As Double
	Private m_dblLength As Double
	Private m_fIsInactive As Boolean
	
	
	
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
	'1 june 2002
	'-
	Public Sub Load(ByVal strID As String)
		
        Dim strGuid As String
        Dim i As Integer
        'UPGRADE_WARNING: Couldn't resolve default property of object m_strTable. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        m_objRS = OpenDB(m_objConn, m_objRS, m_strTable, m_strSQL & " WHERE GuidID = {" & strID & "}")

        'select the desired record
        'm_objRS.Filter = DB_TABLE_PK & " = '" & strID & "'"
        For i = 0 To m_objRS.Tables(m_strTable).Rows.Count - 1
            strGuid = CType((m_objRS.Tables(m_strTable).Rows(i).Item("GuidID")), Guid).ToString()
            If strGuid = strID Then
                With m_objRS.Tables(m_strTable).Rows(i)
                    'copy the values into the class members
                    m_strID = CType((.Item("GuidID")), Guid).ToString() '.Fields("guidID").Value
                    m_strBrandID = CType((.Item("guidID_Brand")), Guid).ToString() 'm_objRS.Fields("guidID_Brand").Value
                    m_nTestsPerKit = .Item("intTestsPerKit") 'm_objRS.Fields("intTestsPerKit").Value
                    m_dblKitCost = .Item("dblKitCost") 'm_objRS.Fields("dblKitCost").Value
                    m_strKitDimension = .Item("strKitDimension") 'm_objRS.Fields("strKitDimension").Value
                    m_dblWeight = .Item("dblWeight") 'm_objRS.Fields("dblWeight").Value

                    m_dblHeight = .Item("dblHeight") 'm_objRS.Fields("dblHeight").Value
                    m_dblWidth = .Item("dblWidth") 'm_objRS.Fields("dblWidth").Value
                    m_dblLength = .Item("dblLength") 'm_objRS.Fields("dblLength").Value

                    m_fIsInactive = .Item("fInactive") 'm_objRS.Fields("fInactive").Value
                End With
                Exit For
            End If
        Next i

        CloseDB(m_objConn, m_objRS)

    End Sub



    '+
    'Create()
    '
    'creates a new object and gives it an id
    '
    'lbailey
    '1 june 2002
    '-
    Public Function Create() As String

        'create the id and set the defaults
        'UPGRADE_WARNING: Couldn't resolve default property of object g_objGUIDGenerator.GetGUID(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        m_strID = g_objGUIDGenerator.GetGUID()

        'remember that this is a new record
        m_fIsNew = True

        Create = m_strID

    End Function



    '+
    'Update()
    '
    'pushes the current data into the repository
    '
    'lbailey
    '1 june 2002
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
            dsNewRow.Item("guidID_Brand") = New Guid(m_strBrandID)
            dsNewRow.Item("intTestsPerKit") = m_nTestsPerKit
            dsNewRow.Item("dblKitCost") = m_dblKitCost
            dsNewRow.Item("strKitDimension") = IIf(m_strKitDimension = "", System.DBNull.Value, m_strKitDimension)
            dsNewRow.Item("dblWeight") = m_dblWeight
            dsNewRow.Item("dblHeight") = m_dblHeight
            dsNewRow.Item("dblWidth") = m_dblWidth
            dsNewRow.Item("dblLength") = m_dblLength
            dsNewRow.Item("fInactive") = m_fIsInactive
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
                        .Tables(m_strTable).Rows(i).Item("guidID_Brand") = New Guid(m_strBrandID)
                        .Tables(m_strTable).Rows(i).Item("intTestsPerKit") = m_nTestsPerKit
                        .Tables(m_strTable).Rows(i).Item("dblKitCost") = m_dblKitCost
                        .Tables(m_strTable).Rows(i).Item("strKitDimension") = IIf(m_strKitDimension = "", System.DBNull.Value, m_strKitDimension)
                        .Tables(m_strTable).Rows(i).Item("dblWeight") = m_dblWeight
                        .Tables(m_strTable).Rows(i).Item("dblHeight") = m_dblHeight
                        .Tables(m_strTable).Rows(i).Item("dblWidth") = m_dblWidth
                        .Tables(m_strTable).Rows(i).Item("dblLength") = m_dblLength
                        .Tables(m_strTable).Rows(i).Item("fInactive") = m_fIsInactive
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
    '1 june 2002
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
	'1 june 2002
	'-
	Public Function GetID() As String
		GetID = m_strID
	End Function
	
	Public Function GetBrandID() As String
		GetBrandID = m_strBrandID
	End Function
	
	Public Function GetTestsPerKit() As Short
		GetTestsPerKit = m_nTestsPerKit
	End Function
	
	Public Function GetKitCost() As Double
		GetKitCost = m_dblKitCost
	End Function
	
	Public Function GetKitDimension() As String
		'UPGRADE_WARNING: Couldn't resolve default property of object m_strKitDimension. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetKitDimension = m_strKitDimension
	End Function
	
	Public Function GetWeight() As Double
		GetWeight = m_dblWeight
	End Function
	Public Function GetHeight() As Double
		GetHeight = m_dblHeight
	End Function
	Public Function GetWidth() As Double
		GetWidth = m_dblWidth
	End Function
	Public Function GetLength() As Double
		GetLength = m_dblLength
	End Function
	'!fInactive
	Public Function GetInactive() As Boolean
		GetInactive = m_fIsInactive
	End Function
	
	
	
	
	'+
	'standard manipulator procedures
	'
	'lbailey
	'1 june 2002
	'-
	Public Sub SetID(ByRef strID As String)
		m_strID = strID
	End Sub
	
	Public Sub SetBrandID(ByRef strBrandID As String)
		m_strBrandID = strBrandID
	End Sub
	
	Public Sub SetTestsPerKit(ByRef nTestsPerKit As Short)
		m_nTestsPerKit = nTestsPerKit
	End Sub
	
	Public Sub SetKitCost(ByRef dblKitCost As Double)
		m_dblKitCost = dblKitCost
	End Sub
	
	Public Sub SetKitDimension(ByRef strKitDimension As String)
		'UPGRADE_WARNING: Couldn't resolve default property of object m_strKitDimension. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_strKitDimension = strKitDimension
	End Sub
	
	Public Sub SetWeight(ByRef dblWeight As Double)
		m_dblWeight = dblWeight
	End Sub
	
	Public Sub SetHeight(ByRef dblHeight As Double)
		m_dblHeight = dblHeight
	End Sub
	Public Sub SetWidth(ByRef dblWidth As Double)
		m_dblWidth = dblWidth
	End Sub
	Public Sub SetLength(ByRef dblLength As Double)
		m_dblLength = dblLength
	End Sub
	
	Public Sub setInactive(ByRef fInactive As Boolean)
		m_fIsInactive = fInactive
	End Sub
End Class