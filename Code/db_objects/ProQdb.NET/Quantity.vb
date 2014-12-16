Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("QuantityDB_NET.QuantityDB")> Public Class QuantityDB
	'Quantity.cls
	'
	'manages the quantities in the db.  the quantities are sets of
	'selected brands along.  each
	'
	'lbailey
	'6 june 2002
	'-
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         private members
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	
	'the name of the table or view used by this class
	Private Const m_strTable As String = DB_TABLE_QUANTITY
    Private Const m_strSQL As String = "Select * FROM " & DB_TABLE_QUANTITY

	'whether or not the object has been constructed
    Private m_objConn As OleDbConnection
    Private m_objRS As DataSet
    Private m_objAdapter As OleDbDataAdapter
	Private m_fIsNew As Boolean
	
	'the guid of the object
	Private m_strID As String
	
	Private m_strSelectedBrandID As String
	Private m_lCount As Integer
	Private m_intCategory As Short
	
	
	
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
	'1 june 2002
	'-
	Public Function Create() As String
		
		'UPGRADE_WARNING: Couldn't resolve default property of object g_objGUIDGenerator.GetGUID(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_strID = g_objGUIDGenerator.GetGUID()
		m_lCount = 0
		
		m_fIsNew = True
		
		Create = m_strID
		
	End Function
	
	
	
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

        m_objRS = OpenDB(m_objConn, m_objRS, m_strTable, m_strSQL & " WHERE GuidID = {" & strID & "}")

        'select the desired record

        For i = 0 To m_objRS.Tables(m_strTable).Rows.Count - 1
            strGuid = CType((m_objRS.Tables(m_strTable).Rows(i).Item("GuidID")), Guid).ToString()
            If strGuid = strID Then
                With m_objRS.Tables(m_strTable).Rows(i)
                    'copy the values into the class members
                    m_strID = CType((.Item("GuidID")), Guid).ToString()
                    m_strSelectedBrandID = CType((.Item("guidID_SelectedBrand")), Guid).ToString()
                    m_lCount = .Item("lCount")
                    m_intCategory = .Item("intCategory")
                End With
                CloseDB(m_objConn, m_objRS)
                Exit For
            End If
        Next i

    End Sub



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
            dsNewRow.Item("guidID_SelectedBrand") = New Guid(m_strSelectedBrandID)
            dsNewRow.Item("lCount") = m_lCount
            dsNewRow.Item("intCategory") = m_intCategory
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
                        .Tables(m_strTable).Rows(i).Item("guidID_SelectedBrand") = New Guid(m_strSelectedBrandID)
                        .Tables(m_strTable).Rows(i).Item("lCount") = m_lCount
                        .Tables(m_strTable).Rows(i).Item("intCategory") = m_intCategory
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
	'8 june 2002
	'-
	Public Function GetID() As String
		GetID = m_strID
	End Function
	
	Public Function GetSelectedBrandID() As String
		GetSelectedBrandID = m_strSelectedBrandID
	End Function
	
	Public Function GetCount() As Integer
		GetCount = m_lCount
	End Function
	
	Public Function GetCategory() As Short
		GetCategory = m_intCategory
	End Function
	
	
	
	'+
	'standard manipulator procedures
	'
	'lbailey
	'8 june 2002
	'-
	Public Sub SetID(ByRef strID As String)
		m_strID = strID
	End Sub
	
	Public Sub SetSelectedBrandID(ByRef strID As String)
		m_strSelectedBrandID = strID
	End Sub
	
	Public Sub SetCount(ByRef lCount As Integer)
		m_lCount = lCount
	End Sub
	
	Public Sub SetCategory(ByRef intCategory As Short)
		m_intCategory = intCategory
	End Sub
End Class