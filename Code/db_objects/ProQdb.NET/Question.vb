Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("QuestionDB_NET.QuestionDB")> Public Class QuestionDB
	'Question.cls
	'Manages the questions
	'31-May-2002 lblanken
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         private members
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	'the dsn used to connect to this db/table
    Private m_strDSN As String = G_STRDSN
	'the name of the table or view used by this class
	Private Const m_strTable As String = DB_TABLE_QUESTION
    Private Const m_strSQL As String = "Select * FROM " & DB_TABLE_QUESTION

	'whether or not the object has been constructed
    Private m_objConn As New OleDbConnection
    Private m_objRS As DataSet
	
	'the ID of the object
	Private m_lngID As Integer
	'the ID of the edit rule associated with
	Private m_lngEditRuleID As Integer
	'the Title of the question
	Private m_strName As String
	'the Text of the question
	Private m_strText As String
	'whether or not the question is an override question
	Private m_fOverride As Boolean
	'type of override
	Private m_lngORType As Integer
    'sp & recordset functions
    Private m_objDBExec As New DBExec

	
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
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		'Intialize the object here
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'Terminate the object here
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
	
	
	
	Public Function Load(ByVal lngID As Integer) As Object
		' Comments   : gets the specified record and populates this object with its values
		' Parameters : lngID - ID to load
		' Returns    :
		' Modified   : 31-May-2002 LKB
		' --------------------------------------------------------
		
        Dim i As Integer
        Dim ID As Integer

        'UPGRADE_WARNING: Couldn't resolve default property of object m_strTable. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        m_objRS = OpenDB(m_objConn, m_objRS, m_strTable, m_strSQL & " WHERE ID = " & lngID)


        'select the desired record
        'm_objRS.Filter = "ID = " & lngID

        For i = 0 To m_objRS.Tables(m_strTable).Rows.Count - 1
            ID = m_objRS.Tables(m_strTable).Rows(i).Item("ID")
            If ID = lngID Then
                With m_objRS.Tables(m_strTable).Rows(i)
                    'copy the values into the class members
                    m_lngID = .Item("id") '.Fields("id").Value
                    'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                    m_lngEditRuleID = IIf(IsDBNull(.Item("ID_EditRule")), 0, .Item("ID_EditRule")) 'IIf(IsDBNull(.Fields("ID_EditRule").Value), 0, .Fields("ID_EditRule").Value)
                    m_strName = .Item("strName") '.Fields("strName").Value
                    m_strText = .Item("strText") '.Fields("strText").Value
                    m_fOverride = .Item("fOverride") '.Fields("fOverride").Value
                    'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                    m_lngORType = IIf(IsDBNull(.Item("lngORType")), 0, .Item("lngORType")) 'IIf(IsDBNull(.Fields("lngORType").Value), 0, .Fields("lngORType").Value)
                End With
                'CloseDB(m_objConn, m_objRS)
                Exit For
            End If
        Next i

        CloseDB(m_objConn, m_objRS)

    End Function

    Public Function Update(ByRef lngID As Integer) As Object
        ' Comments   : Updates the Record to the db using the info that we've got
        ' Parameters : lngID - ID to update
        ' Returns    : -
        ' Modified   : 31-May-2002 LKB
        ' --------------------------------------------------------

        Dim i As Integer
        Dim cb As OleDb.OleDbCommandBuilder
        Dim objAdapter As OleDbDataAdapter
        'connect to the db and get a recordset
        m_objConn = GetConnection(DB_DSN)
        objAdapter = New OleDb.OleDbDataAdapter(m_strSQL, m_objConn)
        objAdapter.Fill(m_objRS, m_strTable)
        cb = New OleDb.OleDbCommandBuilder(objAdapter)

        With m_objRS
            For i = 0 To .Tables(m_strTable).Rows.Count - 1
                If .Tables(m_strTable).Rows(i).Item("ID") = lngID Then
                    'stuff all of the record's properties
                    .Tables(m_strTable).Rows(i).Item("id") = m_lngID
                    .Tables(m_strTable).Rows(i).Item("ID_EditRule") = m_lngEditRuleID
                    .Tables(m_strTable).Rows(i).Item("strName") = m_strName
                    .Tables(m_strTable).Rows(i).Item("strText") = m_strText
                    .Tables(m_strTable).Rows(i).Item("fOverride") = m_fOverride
                    .Tables(m_strTable).Rows(i).Item("lngORType") = m_lngORType
                    'write the record
                    objAdapter.Update(m_objRS, m_strTable) '.Update()

                    Exit For
                End If
            Next
        End With

        'clean up
        CloseDB(m_objConn, m_objRS)

    End Function

    Public Function Delete(ByVal lngID As Integer) As Object
        ' Comments   : Removes the specified object from the repository
        ' Parameters : lngID - ID to delete
        ' Returns    :  -
        ' Modified   : 31-May-2002 LKB
        ' --------------------------------------------------------

        'UPGRADE_WARNING: Couldn't resolve default property of object m_strTable. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Dim i As Integer
        Dim ID As Integer
        Dim da As OleDbDataAdapter
        'get a connection to the db
        m_objConn = GetConnection(DB_DSN)
        da = New OleDbDataAdapter(m_strSQL, m_objConn)
        m_objRS = New DataSet
        da.Fill(m_objRS, m_strTable)
        Dim cb As New OleDb.OleDbCommandBuilder(da)

        'load the data into memory
        For i = 0 To m_objRS.Tables(m_strTable).Rows.Count - 1
            ID = (m_objRS.Tables(m_strTable).Rows(i).Item("ID"))
            If ID = lngID Then
                m_objRS.Tables(m_strTable).Rows(i).Delete()
            End If
        Next
        da.Update(m_objRS, m_strTable)

        CloseDB(m_objConn, m_objRS)

    End Function
	
    Public Function Exec(ByVal strStoredProc As String, ByRef aParams(,) As Object) As DataSet 'ADODB.Recordset
        ' Comments   : Executes the specified query, using the supplied list of parameters, and
        '              returns a query string
        ' Parameters : strStoredProc - Query to open
        '              aParams - Parameters for query
        ' Returns    : recordset based on query and parameters
        ' Modified   : 31-May-2002 LKB
        ' --------------------------------------------------------

        'create a command object
        'Dim objCommand As New ADODB.Command
        'Dim aParam As Object
        'Dim prmParam As ADODB.Parameter

        'associate the command with the conn
        'objCommand.let_ActiveConnection(m_objConn)
        'declare that it's a stored proc
        'objCommand.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        'give it a timeout (30 sec is the default...)
        'objCommand.CommandTimeout = DB_COMMAND_TIMEOUT
        'name the sp
        'objCommand.CommandText = strStoredProc

        'create the sp's parameter list
        'For Each aParam In aParams
        'prmParam = New ADODB.Parameter
        'prmParam.Type = ADODB.DataTypeEnum.adVarChar
        'prmParam.Direction = ADODB.ParameterDirectionEnum.adParamInput
        'objCommand.Parameters.Append(prmParam)
        'Next aParam

        'm_objRS = objCommand.Execute()
        If aParams Is Nothing Then
            m_objRS = m_objDBExec.ReturnRSNoParams(strStoredProc, aParams)
        Else
            m_objRS = m_objDBExec.ReturnRS(m_objConn, strStoredProc, aParams)
        End If
        'get the id
        m_lngID = m_objRS.Tables(m_strTable).Rows(0).Item("ID")

        'set filter based on id
        Load(m_lngID)

        Exec = m_objRS

    End Function
	
	
	'-----------------------------------------------------------------
	'Standard Accessor Functions
	' Modified   : 31-May-2002 LKB
	'-----------------------------------------------------------------
	
	Public Function GetID() As Integer
		GetID = m_lngID
	End Function
	
	Public Function GetEditRuleID() As Integer
		GetEditRuleID = m_lngEditRuleID
	End Function
	
	Public Function GetName() As String
		GetName = m_strName
	End Function
	
	Public Function GetText() As String
		GetText = m_strText
	End Function
	
	Public Function GetOverride() As Boolean
		GetOverride = m_fOverride
	End Function
	Public Function GetORType() As Integer
		GetORType = m_lngORType
	End Function
	
	'-----------------------------------------------------------------
	'Standard Manipulator Functions
	' Modified   : 31-May-2002 LKB
	'-----------------------------------------------------------------
	
	Public Function SetID(ByRef lngID As Integer) As Short
		'set the member
		m_lngID = lngID
		'set the return code
		SetID = S_OK
	End Function
	
	Public Function SetEditRuleID(ByRef lngEditRuleID As Integer) As Short
		'set the member
		m_lngEditRuleID = lngEditRuleID
		'set the return code
		SetEditRuleID = S_OK
	End Function
	
	Public Function SetName(ByRef strName As String) As Short
		'set the member
		m_strName = strName
		'set the return code
		SetName = S_OK
	End Function
	
	Public Function SetText(ByRef strText As String) As Short
		'set the member
		m_strText = strText
		'set the return code
		SetText = S_OK
	End Function
	
	Public Function SetOverride(ByRef fOverride As Boolean) As Short
		'set the member
		m_fOverride = fOverride
		'set the return code
		SetOverride = S_OK
	End Function
	
	Public Function SetORType(ByRef lngORType As Integer) As Short
		'set the member
		m_lngORType = lngORType
		'set the return code
		SetORType = S_OK
	End Function
End Class