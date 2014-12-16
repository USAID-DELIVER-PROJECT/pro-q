Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("QuantityBrandDB_NET.QuantityBrandDB")> Public Class QuantityBrandDB
	'+
	'SelectedBrandDB.cls
	'
	'each record in this table represents a distinct brand associated
	'with a
	
	
	
	'data source
    Private m_strDSN As String = DB_DSN
	'the name of the table or view used by this class
    Private Const m_strTable As String = DB_TABLE_SELECTED_BRAND
    Private Const m_strSQL As String = "Select * From " & DB_TABLE_SELECTED_BRAND
    'whether or not the object has been constructed
    Private m_objConn As OleDbConnection
    Private rst As DataSet

    Private m_strID As String
    Private m_strBrandID As String
    Private m_strKitID As String
    Private m_intTestsPerKit As Short
    Private m_dblKitCost As Double
    Private m_strQuantificationID As String
    Private m_lCount As Integer



    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '                         private methods
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


    'Class_Initialize() (constructor)
    '
    'connect to the db and load the data into memory
    '
    'lbailey
    '16 may 2002
    'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Private Sub Class_Initialize_Renamed()
        'get a connection to the db
        'Set m_objConn = GetConnection()

        'load the data into memory
        'rst.Open m_strTable, m_objConn, adOpenDynamic, adLockPessimistic, adCmdTable

    End Sub
    Public Sub New()
        MyBase.New()
        Class_Initialize_Renamed()
    End Sub

    'Load()
    '
    'gets the specified record and populates this object with its values
    '
    Public Sub Load(ByRef strID As String)

        Dim i As Integer
        Dim strGuid As String
        Dim rst As DataSet
        Dim objAdapter As OleDbDataAdapter
        'get a connection to the db
        m_objConn = GetConnection(m_strDSN)

        'load the data into memory
        'rst.Open(m_strTable, m_objConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockPessimistic, ADODB.CommandTypeEnum.adCmdTable)
        objAdapter = New OleDb.OleDbDataAdapter(m_strSQL & " WHERE GuidID = {" & strID & "}", m_objConn)
        rst = New DataSet
        objAdapter.Fill(rst, m_strTable)

        'select the desired record
        '.Filter = "guidID = '" & strID & "'"
        For i = 0 To rst.Tables(m_strTable).Rows.Count - 1
            strGuid = CType((rst.Tables(m_strTable).Rows(i).Item(db_table_pk)), Guid).ToString()
            If strGuid = strID Then
                With rst.Tables(m_strTable).Rows(i)
                    'copy the values into the class members
                    m_strID = CType((.Item("GuidID")), Guid).ToString() '.Fields("guidID").Value
                    m_strBrandID = CType((.Item("guidID_Brand")), Guid).ToString() '.Fields("guidID_Brand").Value
                    m_strKitID = CType((.Item("guidID_Kit")), Guid).ToString() '.Fields("guidID_Kit").Value
                    m_intTestsPerKit = .Item("intTestsPerKit") '.Fields("intTestsPerKit").Value
                    m_dblKitCost = .Item("dblKitCost") '.Fields("dblKitCost").Value
                End With
                CloseDB(m_objConn, rst)
                Exit For
            End If
        Next i

        'UPGRADE_NOTE: Object rst may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        rst = Nothing
        'UPGRADE_NOTE: Object m_objConn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        m_objConn = Nothing

    End Sub


    Public Sub Update()

        Dim i As Integer
        Dim cb As OleDb.OleDbCommandBuilder
        Dim rst As New DataSet
        Dim objAdapter As OleDbDataAdapter
        'connect to the db and get a recordset
        m_objConn = GetConnection(DB_DSN)
        objAdapter = New OleDb.OleDbDataAdapter(m_strSQL, m_objConn)
        objAdapter.Fill(rst, m_strTable)
        cb = New OleDb.OleDbCommandBuilder(objAdapter)

        With rst
            For i = 0 To .Tables(m_strTable).Rows.Count - 1
                Dim strGuid As String
                strGuid = CType(.Tables(m_strTable).Rows(i).Item(DB_TABLE_PK), Guid).ToString
                If strGuid = m_strID Then
                    'stuff all of the record's properties
                    .Tables(m_strTable).Rows(i).Item("guidID_Brand") = New Guid(m_strBrandID)
                    .Tables(m_strTable).Rows(i).Item("guidID_Kit") = New Guid(m_strKitID)
                    .Tables(m_strTable).Rows(i).Item("intTestsPerKit") = m_intTestsPerKit
                    .Tables(m_strTable).Rows(i).Item("dblKitCost") = m_dblKitCost
                    'write the record
                    objAdapter.Update(rst, m_strTable) '.Update()

                    Exit For
                End If
            Next
        End With

        'clean up
        CloseDB(m_objConn, rst)
        'UPGRADE_NOTE: Object rst may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        rst = Nothing
        'UPGRADE_NOTE: Object m_objConn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        m_objConn = Nothing

    End Sub


    '+
    'Create()
    '
    'Adds the New Record to the db using the info that we've got
    '
    '-
    Public Function Create() As String

        '    'here's what we're a'gonna do
        '    rst.AddNew
        '
        '    With rst
        '      'if we don't already have an id for this, then create one
        '      If (Len(m_strID) < 1) Then
        '          'use the guid generator to create an id
        '          Dim objGUID As New JSIGUIDGen.GUIDGen
        '           m_strID = objGUID.GetGUID()
        '      End If
        '
        '      'stuff all of the record's properties
        '      !guidID = m_strID
        '      !guidID_Brand = m_strBrandID
        '      !guidID_Kit = m_strKitID
        '      !intTestsPerKit = m_intTestsPerKit
        '      !dblKitCost = m_dblKitCost
        '
        '      'write the record
        '      .Update
        '
        '    End With
        '
        '    'return the id of the created object
        '    Create = rst!guidID

        'UPGRADE_WARNING: Couldn't resolve default property of object g_objGUIDGenerator.GetGUID(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        m_strID = g_objGUIDGenerator.GetGUID()
        Create = m_strID

    End Function


    Public Sub Delete()

        Dim rst As New DataSet

        DeleteRecord(m_objConn, rst, m_strTable, m_strSQL, m_strID)
        CloseDB(m_objConn, rst)


        'UPGRADE_NOTE: Object rst may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        rst = Nothing
        'UPGRADE_NOTE: Object m_objConn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        m_objConn = Nothing

    End Sub
	
	
	'Class_Terminate() (destructor)
	'
	'cleans up whatever needs to be cleaned up when this object is
	'released.
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'release the connections to the db
		'rst.Close
		'm_objConn.Close
		
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
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
	
	Public Sub setBrandID(ByRef strBrandID As String)
		m_strBrandID = strBrandID
	End Sub
	
	Public Sub SetKitID(ByRef strKitID As String)
		m_strKitID = strKitID
	End Sub
	
	Public Sub SetTestsPerKit(ByRef intTestsPerKit As Short)
		m_intTestsPerKit = intTestsPerKit
	End Sub
	
	Public Sub SetKitCost(ByRef dblKitCost As Double)
		m_dblKitCost = dblKitCost
	End Sub
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'standard accessor functions
	
	Public Function GetID() As String
		GetID = m_strID
	End Function
	
	Public Function getBrandID() As String
		getBrandID = m_strBrandID
	End Function
	
	Public Function GetKitID() As String
		GetKitID = m_strKitID
	End Function
	
	Public Function GetTestsPerKit() As String
		GetTestsPerKit = CStr(m_intTestsPerKit)
	End Function
	
	Public Function GetKitCost() As Integer
		GetKitCost = m_dblKitCost
	End Function
End Class