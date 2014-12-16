Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("AggregationDB_NET.AggregationDB")> Public Class AggregationDB
	'the purpose of this class is to manage the connections to the db
	'
	'lbailey
	'14 may 2002
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                            constants
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	'result codes
	Private Const S_OK As Short = 0
	
	'error strings
	Private Const ERROR_CREATE_1 As String = "Error: You need to provide input values."
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         private members
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	'the dsn used to connect to this db/table
    Private m_strDSN As String = DB_DSN
	'the name of the table or view used by this class
    Private Const m_strTable As String = DB_TABLE_AGGREGATION

    'whether or not the object has been constructed
    Private m_objConn As OleDbConnection '.Connection
    Private m_objAdapter As OleDbDataAdapter '.Connection
    Private Const m_strSQL As String = "Select * From " & DB_TABLE_AGGREGATION
    Private m_objRS As New DataSet 'ADODB.Recordset
	
	'the guid of the object
    Private m_strID As String
    'the friendly name
	Private m_strName As String
	'like "Ghana, Malawi"  the subject of the analysis
	Private m_strCountry As String
	'the ministry or public/private program for whom we're doing the analysis
	Private m_strProgram As String
	'free text
	Private m_strNotes As String
	'the user who created this aggregation
	Private m_strPreparedBy As String
	'the date on which the user created this aggregation
	Private m_dtmCreated As Date
	'the user who last modified the record
	Private m_strModifiedBy As String
	' The Date/Time the record was last modified
	Private m_dtmModified As Date
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         public members
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	'hide everything from the user
	
	
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
	'
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		
		'get a connection to the db
		'Set m_objConn = GetConnection(m_strDSN)
		
		'load the data into memory
		'm_objRS.Open m_strTable, m_objConn, adOpenDynamic, adLockPessimistic, adCmdTable
		
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	
	'Class_Terminate() (destructor)
	'
	'cleans up whatever needs to be cleaned up when this object is
	'released.
	'
	'lbailey
	'16 may 2002
	'
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'release the connections to the db
		'If m_objRS.State = adStateOpen Then
		' m_objRS.Close
		'End If
		'Set m_objConn = Nothing
		
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	
	
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
    Public Function Load(ByVal strID As String) As String

        Dim i As Integer
        Dim strGuid As String
        'get a connection to the db
        m_objConn = GetConnection(m_strDSN)

        'load the data into memory
        'rst.Open(m_strTable, m_objConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockPessimistic, ADODB.CommandTypeEnum.adCmdTable)
        m_objAdapter = New OleDb.OleDbDataAdapter(m_strSQL, m_objConn)
        m_objRS = New DataSet
        m_objAdapter.Fill(m_objRS, m_strTable)

        'select the desired record
        '.Filter = "guidID = '" & strID & "'"
        For i = 0 To m_objRS.Tables(m_strTable).Rows.Count - 1
            strGuid = CType((m_objRS.Tables(m_strTable).Rows(i).Item("GuidID")), Guid).ToString()
            If strGuid = strID Then
                With m_objRS.Tables(m_strTable).Rows(i)
                    'copy the values into the class members
                    m_strID = CType((.Item("GuidID")), Guid).ToString()             '.Fields("guidID").Value
                    m_strName = .Item("strName")                                    '.Fields("strName").Value
                    m_strCountry = CType(.Item("guidID_Country"), Guid).ToString()  '.Fields("guidID_Country").Value
                    m_strProgram = CType(.Item("guidID_Program"), Guid).ToString()  '.Fields("guidID_Program").Value
                    m_strNotes = IIf(IsDBNull(.Item("memNotes")), "", .Item("memNotes")) 'IIf(IsDBNull(.Fields("memNotes").Value), "", .Fields("memNotes").Value)
                    m_strPreparedBy = .Item("strPreparedBy")                        '.Fields("strPreparedBy").Value
                    m_dtmCreated = .Item("dtmCreated")                              '.Fields("dtmCreated").Value
                    m_dtmModified = .Item("dtmModified")                            '.Fields("dtmModified").Value
                End With
                Exit For
            End If
        Next i

        CloseDB(m_objConn, m_objRS)

        'return result code
        Return S_OK


    End Function
	
	'+
	'Create()
	'
	'Adds the New Record to the db using the info that we've got
	'
	'lbailey
	'16 may 2002
	'-
	Public Function Create() As String
		
		'get a connection to the db
		m_objConn = GetConnection(m_strDSN)
        Dim cb As OleDb.OleDbCommandBuilder
        Dim dsNewRow As DataRow
        Dim rst As New DataSet

        'connect to the db and get a recordset
        m_objConn = GetConnection(DB_DSN)
        m_objAdapter = New OleDb.OleDbDataAdapter(m_strSQL, m_objConn)
        m_objAdapter.Fill(rst, m_strTable)
        cb = New OleDb.OleDbCommandBuilder(m_objAdapter)

        dsNewRow = rst.Tables(m_strTable).NewRow()

        'if we don't already have an id for this, then create one
        If (Len(m_strID) < 1) Then
            'use the guid generator to create an id
            m_strID = g_objGUIDGenerator.GetGUID()
        End If
        'stuff all of the record's properties

        '/ Old Method
        'With rst
        '.Fields("strName").Value = m_strName
        '.Fields("guidID_Country").Value = m_strCountry
        '.Fields("guidID_Program").Value = m_strProgram
        '.Fields("memNotes").Value = m_strNotes
        '.Fields("strPreparedBy").Value = m_strPreparedBy
        '.Fields("dtmCreated").Value = m_dtmCreated
        '.Fields("dtmModified").Value = Now m_dtmModified
        '.Fields("strModifiedBy").Value = IIf(Len(m_strModifiedBy) = 0, System.DBNull.Value, m_strModifiedBy)
        'End With

        '/ New Method    
        dsNewRow.Item("guidID") = New Guid(m_strID)
        dsNewRow.Item("strName") = m_strName
        dsNewRow.Item("guidID_Country") = New Guid(m_strCountry)
        dsNewRow.Item("guidID_Program") = New Guid(m_strProgram)
        dsNewRow.Item("memNotes") = m_strNotes
        dsNewRow.Item("strPreparedBy") = m_strPreparedBy
        dsNewRow.Item("dtmCreated") = m_dtmCreated
        dsNewRow.Item("dtmModified") = Now
        dsNewRow.Item("strModifiedBy") = IIf(Len(m_strModifiedBy) = 0, System.DBNull.Value, m_strModifiedBy)
        rst.Tables(m_strTable).Rows.Add(dsNewRow)
        m_objAdapter.Update(rst, m_strTable)

        'clean up
        CloseDB(m_objConn, rst)

        'return the id of the created object
        Create = m_strID '.Fields("guidID").Value()

    End Function
    Public Sub Update()
        Dim i As Integer
        Dim cb As OleDb.OleDbCommandBuilder
        Dim rst As New DataSet
        'connect to the db and get a recordset
        m_objConn = GetConnection(DB_DSN)
        m_objAdapter = New OleDb.OleDbDataAdapter(m_strSQL, m_objConn)
        m_objAdapter.Fill(rst, m_strTable)
        cb = New OleDb.OleDbCommandBuilder(m_objAdapter)

        With rst
            For i = 0 To .Tables(m_strTable).Rows.Count - 1
                Dim strGuid As String
                strGuid = CType(.Tables(m_strTable).Rows(i).Item(DB_TABLE_PK), Guid).ToString
                If strGuid = m_strID Then
                    'stuff all of the record's properties
                    With .Tables(m_strTable).Rows(i)
                        .Item("strName") = m_strName
                        .Item("guidID_Country") = New Guid(m_strCountry)
                        .Item("guidID_Program") = New Guid(m_strProgram)
                        .Item("memNotes") = m_strNotes
                        .Item("strPreparedBy") = m_strPreparedBy
                        .Item("strPreparedBy") = m_strPreparedBy
                        .Item("dtmModified") = Now()
                        .Item("strModifiedBy") = IIf(Len(m_strModifiedBy) = 0, System.DBNull.Value, m_strModifiedBy)
                    End With
                    'write the record
                    m_objAdapter.Update(rst, m_strTable) '.Update()
                    Exit For
                End If
            Next
        End With

        'clean up
        CloseDB(m_objConn, rst)

    End Sub

    '+
    'Delete()
    '
    'Removes the specified object from the repository
    '
    'lbailey 16 may 2002
    '-
    Public Sub Delete(ByVal strID As String)

        Dim rst As New DataSet

        DeleteRecord(m_objConn, rst, m_strTable, m_strSQL, strID)
        CloseDB(m_objConn, rst)
    End Sub
    Public Function ReturnRS(ByRef strStoredProc As String, ByRef aParams(,) As Object) As dataset 'ADODB.Recordset

        Dim objDbExec As New DBExec
        m_objConn = GetConnection(m_strDSN)
        ReturnRS = objDbExec.ReturnRS(m_objConn, strStoredProc, aParams)
        objDbExec = Nothing
        m_objConn = Nothing

    End Function
    
	
	'+
	'Exec()
	'
	'Executes the specified query, using the supplied list of parameters, and
	'returns a query string
	'
	'-
    Public Function Exec(ByVal strStoredProc As String, ByRef aParams(,) As Object) As dataset

        'create a command object
        Dim objCommand As New ADODB.Command
        Dim aParam As Object
        Dim prmParam As ADODB.Parameter
        'associate the command with the conn
        objCommand.let_ActiveConnection(m_objConn)
        'declare that it's a stored proc
        objCommand.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        'give it a timeout (30 sec is the default...)
        objCommand.CommandTimeout = DB_COMMAND_TIMEOUT
        'name the sp
        objCommand.CommandText = strStoredProc
        'create the sp's parameter list
        For Each aParam In aParams
            prmParam = New ADODB.Parameter
            prmParam.Type = ADODB.DataTypeEnum.adVarChar
            prmParam.Direction = ADODB.ParameterDirectionEnum.adParamInput
            objCommand.Parameters.Append(prmParam)
        Next aParam

        m_objRS = objCommand.Execute()

        'get the id
        'm_strID = m_objRS.Fields("guidID").Value
        m_strID = CType((m_objRS.Tables(m_strTable).Rows(0).Item("GuidID")), Guid).ToString()
        'set filter based on id
        Load(m_strID)

        Exec = m_objRS

    End Function
	
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
	
	Public Function GetCountry() As String
		GetCountry = m_strCountry
	End Function
	
	Public Function GetProgram() As String
		GetProgram = m_strProgram
	End Function
	
	Public Function GetNotes() As String
		GetNotes = m_strNotes
	End Function
	
	Public Function GetCreator() As String
		GetCreator = m_strPreparedBy
	End Function
	
	Public Function GetCreationDate() As String
		GetCreationDate = CStr(m_dtmCreated)
	End Function
	Public Function GetModifiedBy() As String
		GetModifiedBy = m_strModifiedBy
	End Function
	Public Function GetModifiedDate() As String
		GetModifiedDate = CStr(m_dtmModified)
	End Function
	
	
	'+
	'standard manipulator functions
	'
	'lbailey
	'16 may 2002
	'-
	Public Function SetID(ByRef strID As String) As Short
		'set the member
		m_strID = strID
		'set the return code
		SetID = S_OK
	End Function
	
	Public Function SetName(ByRef strName As String) As Short
		'set the member
		m_strName = strName
		'set the return code
		SetName = S_OK
	End Function
	
	Public Function SetCountry(ByRef strCountry As String) As Short
		'set the member
		m_strCountry = strCountry
		'set the return code
		SetCountry = S_OK
	End Function
	
	Public Function SetProgram(ByRef strProgram As String) As Short
		'set the member
		m_strProgram = strProgram
		'set the return code
		SetProgram = S_OK
	End Function
	
	Public Function SetNotes(ByRef strNotes As String) As Short
		'set the member
		m_strNotes = strNotes
		'set the return code
		SetNotes = S_OK
	End Function
	
	Public Function SetCreator(ByRef strCreator As String) As Short
		'set the member
		m_strPreparedBy = strCreator
		'set the return code
		SetCreator = S_OK
	End Function
	
	Public Function SetCreationDate(ByRef dCreationDate As Date) As Short
		'set the member
		m_dtmCreated = dCreationDate
		'set the return code
		SetCreationDate = S_OK
	End Function
	
	Public Function SetModifiedBy(ByRef strModifiedBy As String) As Short
		'set the member
		m_strModifiedBy = strModifiedBy
		'set the return code
		SetModifiedBy = S_OK
	End Function
End Class