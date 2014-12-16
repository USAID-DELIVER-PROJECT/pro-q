Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("SelectedBrandDB_NET.SelectedBrandDB")> Public Class SelectedBrandDB
	'+
	'SelectedBrandDB.cls
	'
	'each record in this table represents a distinct brand associated
	'with a methodology, protocol, or section.
	'
	'lbailey
	'8 june 2002
	'-
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         private methods
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	'the name of the table or view used by this class
	Private Const m_strTable As String = DB_TABLE_SELECTED_BRAND
    Private Const m_strSQL As String = "Select * FROM " & DB_TABLE_SELECTED_BRAND

	'whether or not the object has been constructed
    Private m_objConn As OleDbConnection
    Private m_objRS As DataSet
    Private m_objAdapter As OleDbDataAdapter
	Private m_fIsNew As Boolean
	
	'data members
	Private m_strID As String
	Private m_strBrandID As String
	Private m_strKitID As String
	Private m_intTestsPerKit As Short
	Private m_dblKitCost As Double
	Private m_strQuantificationID As String
	Private m_lCount As Integer
	Private m_bGeneric As Boolean
	Private m_strGenericCode As String
	
	
	'+
	'Create()
	'
	'Adds the New Record to the db using the info that we've got
	'
	'lbailey
	'8 june 2002
	'-
	Public Function Create() As String
		
		'create the id and set the defaults
		'UPGRADE_WARNING: Couldn't resolve default property of object g_objGUIDGenerator.GetGUID(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_strID = g_objGUIDGenerator.GetGUID()
		
		'remember that this is a new record
		m_bGeneric = False
		m_fIsNew = True
		
		Create = m_strID
		
	End Function
	
	
	
	'+
	'Load()
	'
	'gets the specified record and populates this object with its values
	'
	'lbailey
	'8 june 2002
	'-
	Public Sub Load(ByRef strID As String)
		
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
                    m_strBrandID = CType((.Item("guidID_Brand")), Guid).ToString() 'm_objRS.Fields("guidID_Brand").Value
                    m_strQuantificationID = CType((.Item("guidID_Quantification")), Guid).ToString() 'm_objRS.Fields("guidID_Quantification").Value
                    m_strKitID = CType((.Item("guidID_Kit")), Guid).ToString() 'm_objRS.Fields("guidID_Kit").Value
                    m_intTestsPerKit = .Item("intTestsPerKit") 'm_objRS.Fields("intTestsPerKit").Value
                    m_dblKitCost = .Item("dblKitCost") 'm_objRS.Fields("dblKitCost").Value
                    m_lCount = .Item("lngCount") 'm_objRS.Fields("lngCount").Value
                    m_bGeneric = .Item("fGeneric") 'm_objRS.Fields("fGeneric").Value
                    m_strGenericCode = .Item("strGenericCode") 'm_objRS.Fields("strGenericCode").Value
                End With
                CloseDB(m_objConn, m_objRS)
                Exit For
            End If
        Next i

    End Sub



    '+
    'Update()
    '
    'puts the current object into the db (saves it)
    '
    'lbailey
    '8 june 2002
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
            dsNewRow.Item("guidID_Kit") = New Guid(m_strKitID)
            dsNewRow.Item("guidID_Quantification") = New Guid(m_strQuantificationID)
            dsNewRow.Item("intTestsPerKit") = m_intTestsPerKit
            dsNewRow.Item("fGeneric") = m_bGeneric
            dsNewRow.Item("dblKitCost") = m_dblKitCost
            dsNewRow.Item("lngCount") = m_lCount
            dsNewRow.Item("strGenericCode") = m_strGenericCode
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
                        .Tables(m_strTable).Rows(i).Item("guidID_Kit") = New Guid(m_strKitID)
                        .Tables(m_strTable).Rows(i).Item("guidID_Quantification") = New Guid(m_strQuantificationID)
                        .Tables(m_strTable).Rows(i).Item("intTestsPerKit") = m_intTestsPerKit
                        .Tables(m_strTable).Rows(i).Item("fGeneric") = m_bGeneric
                        .Tables(m_strTable).Rows(i).Item("dblKitCost") = m_dblKitCost
                        .Tables(m_strTable).Rows(i).Item("lngCount") = m_lCount
                        .Tables(m_strTable).Rows(i).Item("strGenericCode") = m_strGenericCode                        'write the record
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
    'removes the current record from the table
    '
    'lbailey
    '8 june 2002
    '-
    Public Sub Delete()

        'OpenDB(m_objConn, m_objRS, m_strTable, m_strSQL)
        DeleteRecord(m_objConn, m_objRS, m_strTable, m_strSQL, m_strID)
        CloseDB(m_objConn, m_objRS)

    End Sub
	
	
	
	
	'+
	'standard manipulator functions
	'
	'lbailey
	'16 may 2002
	'-
	Public Sub SetID(ByRef strID As String)
		m_strID = strID
	End Sub
	
	Public Sub SetBrandID(ByRef strBrandID As String)
		m_strBrandID = strBrandID
	End Sub
	Public Sub SetKitID(ByRef strKitID As String)
		m_strKitID = strKitID
	End Sub
	Public Sub SetGeneric(ByRef bGeneric As Boolean)
		m_bGeneric = bGeneric
	End Sub
	Public Sub SetGenericCode(ByRef strCode As String)
		m_strGenericCode = strCode
	End Sub
	Public Sub setQuantificationID(ByRef strID As String)
		m_strQuantificationID = strID
	End Sub
	
	Public Sub SetTestsPerKit(ByRef intTestsPerKit As Short)
		m_intTestsPerKit = intTestsPerKit
	End Sub
	
	Public Sub SetKitCost(ByRef dblKitCost As Double)
		m_dblKitCost = dblKitCost
	End Sub
	
	Public Sub SetCount(ByRef lCount As Integer)
		m_lCount = lCount
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
	
	Public Function GetBrandID() As String
		GetBrandID = m_strBrandID
	End Function
	
	Public Function GetKitID() As String
		GetKitID = m_strKitID
	End Function
	Public Function getQuantificationID() As String
		getQuantificationID = m_strQuantificationID
	End Function
	Public Function GetGeneric() As Boolean
		GetGeneric = m_bGeneric
	End Function
	Public Function GetGenericCode() As String
		GetGenericCode = m_strGenericCode
	End Function
	Public Function GetTestsPerKit() As String
		GetTestsPerKit = CStr(m_intTestsPerKit)
	End Function
	
	Public Function GetKitCost() As Double
		GetKitCost = m_dblKitCost
	End Function
	
	Public Function GetCount() As Integer
		GetCount = m_lCount
	End Function
End Class