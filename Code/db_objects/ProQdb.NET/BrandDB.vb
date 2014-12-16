Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("BrandDB_NET.BrandDB")> Public Class BrandDB
	'BrandDB.cls
	'
	'Manages the brands in the db
	'
	'lbailey
	'1 june 2002
	
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         private members
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	'the name of the table or view used by this class
	Private Const m_strTable As String = DB_TABLE_BRAND
    Private Const m_strSQL As String = "Select * From " & DB_TABLE_BRAND

    'whether or not the object has been constructed
    Private m_objConn As OleDbConnection 'OleDbConnection
    Private m_objRS As DataSet 'ADODB.Recordset
    Private m_objAdapter As OleDbDataAdapter
    Private m_myView As DataView
    Private m_fIsNew As Boolean

    'the guid of the object
    Private m_strID As String
    'the friendly name
    Private m_strName As String
    'the manufacturer?  distributor?  donor?
    Private m_strSupplier As String
    'elisa, rapid, blot, etc
    Private m_lngType As Integer
    'how many months this will last
    Private m_intShelfLife As Short
    'what temp this should be stored at
    Private m_strStorageTemp As String
    'whether this must be refrigerated
    Private m_fColdStorage As Boolean
    'concomitant accessories
    Private m_strMaterialsRequired As String
    'free text
    Private m_strNotes As String
    'whether the test requires lab facilities (affects svc cap)
    Private m_fIsLab As Boolean
    ' Is the product 'deleted'
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

        Dim i As Integer
        Dim strGuid As String

        m_objRS = OpenDB(m_objConn, m_objRS, m_strTable, m_strSQL)

        'select the desired record
        'm_objRS.Filter = DB_TABLE_PK & " = '" & strID & "'"
        For i = 0 To m_objRS.Tables(m_strTable).Rows.Count - 1
            strGuid = CType((m_objRS.Tables(m_strTable).Rows(i).Item("GuidID")), Guid).ToString()
            If strGuid = strID Then
                With m_objRS.Tables(m_strTable).Rows(i)
                    'copy the values into the class members
                    m_strID = strGuid '.Fields("GuidID").Value
                    m_strName = .Item("strName") '.Fields("strName").Value
                    m_strSupplier = .Item("strSupplier") '.Fields("strSupplier").Value
                    m_lngType = .Item("lngType") '.Fields("lngType").Value
                    m_intShelfLife = .Item("intShelfLife") '.Fields("intShelfLife").Value
                    m_strStorageTemp = .Item("strStorageTemp") '.Fields("strStorageTemp").Value
                    m_fColdStorage = .Item("fColdStorage") '.Fields("fColdStorage").Value
                    m_strMaterialsRequired = IIf(IsDBNull(.Item("memMaterialsRequired")), "", .Item("memMaterialsRequired")) 'IIf(IsDBNull(.Fields("memMaterialsRequired").Value), "", .Fields("memMaterialsRequired").Value)
                    m_strNotes = IIf(IsDBNull(.Item("memNotes")), "", .Item("memNotes")) 'IIf(IsDBNull(.Fields("memNotes").Value), "", .Fields("memNotes").Value)
                    m_fIsLab = .Item("fLabTest") '.Fields("fLabTest").Value
                    m_fIsInactive = .Item("fInactive") '.Fields("fInactive").Value
                End With
                Exit For
            End If
        Next
        CloseDB(m_objConn, m_objRS)

    End Sub



    '+
    'Create()
    '
    'Adds the New Record to the db using the info that we've got
    '
    'lbailey
    '1 june 2002
    '-
    Public Function Create() As String

        'create the id and set the defaults
        'UPGRADE_WARNING: Couldn't resolve default property of object g_objGUIDGenerator.GetGUID(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        m_strID = g_objGUIDGenerator.GetGUID()
        m_strName = "New Brand"
        m_strSupplier = "New Supplier"

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

            '/ OLD METHOD
            'With RST
            '.Fields("GuidID").Value = m_strID
            '.Fields("strName").Value = m_strName
            '.Fields("strSupplier").Value = m_strSupplier
            '.Fields("lngType").Value = m_lngType
            '.Fields("intShelfLife").Value = m_intShelfLife
            '.Fields("strStorageTemp").Value = m_strStorageTemp
            '.Fields("fColdStorage").Value = m_fColdStorage
            '.Fields("memMaterialsRequired").Value = m_strMaterialsRequired
            '.Fields("memNotes").Value = m_strNotes
            '.Fields("fLabTest").Value = m_fIsLab
            '.Fields("fInactive").Value = m_fIsInactive
            'End With

            '/ New Method
            dsNewRow.Item("GuidID") = New Guid(m_strID)
            dsNewRow.Item("strName") = m_strName
            dsNewRow.Item("strSupplier") = m_strSupplier
            dsNewRow.Item("lngType") = m_lngType
            dsNewRow.Item("intShelfLife") = m_intShelfLife
            dsNewRow.Item("strStorageTemp") = m_strStorageTemp
            dsNewRow.Item("fColdStorage") = m_fColdStorage
            dsNewRow.Item("memMaterialsRequired") = m_strMaterialsRequired
            dsNewRow.Item("memNotes") = m_strNotes
            dsNewRow.Item("fLabTest") = m_fIsLab
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
                        .Tables(m_strTable).Rows(i).Item("GuidID") = New Guid(m_strID) '.Fields("GuidID").Value = m_strID
                        .Tables(m_strTable).Rows(i).Item("strName") = m_strName '.Fields("strName").Value = m_strName
                        .Tables(m_strTable).Rows(i).Item("strSupplier") = m_strSupplier '.Fields("strSupplier").Value = m_strSupplier
                        .Tables(m_strTable).Rows(i).Item("lngType") = m_lngType '.Fields("lngType").Value = m_lngType
                        .Tables(m_strTable).Rows(i).Item("intShelfLife") = m_intShelfLife '.Fields("intShelfLife").Value = m_intShelfLife
                        .Tables(m_strTable).Rows(i).Item("strStorageTemp") = m_strStorageTemp '.Fields("strStorageTemp").Value = m_strStorageTemp
                        .Tables(m_strTable).Rows(i).Item("fColdStorage") = m_fColdStorage '.Fields("fColdStorage").Value = m_fColdStorage
                        .Tables(m_strTable).Rows(i).Item("memMaterialsRequired") = m_strMaterialsRequired '.Fields("memMaterialsRequired").Value = m_strMaterialsRequired
                        .Tables(m_strTable).Rows(i).Item("memNotes") = m_strNotes '.Fields("memNotes").Value = m_strNotes
                        .Tables(m_strTable).Rows(i).Item("fLabTest") = m_fIsLab '.Fields("fLabTest").Value = m_fIsLab
                        .Tables(m_strTable).Rows(i).Item("fInactive") = m_fIsInactive '.Fields("fInactive").Value = m_fIsInactive
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
	
	Public Function GetName() As String
		GetName = m_strName
	End Function
	
	Public Function GetSupplier() As String
		GetSupplier = m_strSupplier
	End Function
	
	'UPGRADE_NOTE: GetType was upgraded to GetType_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function GetType_Renamed() As Integer
		GetType_Renamed = m_lngType
	End Function
	
	Public Function GetShelfLife() As Short
		GetShelfLife = m_intShelfLife
	End Function
	
	Public Function GetStorageTemp() As String
		GetStorageTemp = m_strStorageTemp
	End Function
	
	Public Function getColdStorage() As Boolean
		getColdStorage = m_fColdStorage
	End Function
	
	Public Function GetMaterialsRequired() As String
		GetMaterialsRequired = m_strMaterialsRequired
	End Function
	
	Public Function GetNotes() As String
		GetNotes = m_strNotes
	End Function
	
	Public Function IsLab() As Boolean
		IsLab = m_fIsLab
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
	
	Public Sub SetName(ByRef strName As String)
		m_strName = strName
	End Sub
	
	Public Sub SetSupplier(ByRef strSupplier As String)
		m_strSupplier = strSupplier
	End Sub
	
	Public Sub SetType(ByRef lngType As Integer)
		m_lngType = lngType
	End Sub
	
	Public Sub SetShelfLife(ByRef intShelfLife As Short)
		m_intShelfLife = intShelfLife
	End Sub
	
	Public Sub SetStorageTemp(ByRef strStorageTemp As String)
		m_strStorageTemp = strStorageTemp
	End Sub
	
	Public Sub setColdStorage(ByRef fColdStorage As Boolean)
		m_fColdStorage = fColdStorage
	End Sub
	
	Public Sub SetMaterialsRequired(ByRef strMaterialsRequired As String)
		m_strMaterialsRequired = strMaterialsRequired
	End Sub
	
	Public Sub SetNotes(ByRef strNotes As String)
		m_strNotes = strNotes
	End Sub
	
	Public Sub SetLab(ByRef fIsLab As Boolean)
		m_fIsLab = fIsLab
	End Sub
	
	Public Sub setInactive(ByRef fInactive As Boolean)
		m_fIsInactive = fInactive
	End Sub
End Class