Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("ScriptRelationshipDB_NET.ScriptRelationshipDB")> Public Class ScriptRelationshipDB
	'Question.cls
	'Manages the script relationships
	'31-May-2002 lblanken
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         private members
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	'the dsn used to connect to this db/table
    Private m_strDSN As String = G_STRDSN
	'the name of the table or view used by this class
	Private Const m_strTable As String = DB_TABLE_SCRIPT_RELATIONSHIP
    Private Const m_strSQL As String = "Select * FROM " & DB_TABLE_SCRIPT_RELATIONSHIP

	'whether or not the object has been constructed
    Private m_objConn As New OleDbConnection
    Private m_objRS As DataSet
    Private m_objAdapter As OleDbDataAdapter
    'the ID of the object
    Private m_lngID As Integer
    'the ID of the Parent Index
    Private m_lngParentID As Integer
    'the ID of the question
    Private m_lngQuestionID As Integer
    'the description/title for the entry in the script
    Private m_strDescription As String
    'type of the question (i.e. Brand)
    Private m_lngType As Integer
    'type of entry for the script (question or group)
    Private m_lngGroupType As Integer
    'question group to go to if answer to question is true
    Private m_lngTrueGroup As Integer
    'question group to go to if answer to question is False
    Private m_lngFalseGroup As Integer
    'type of entry for the treeview
    Private m_strTreeType As String
    'default sort order for the treeview
    Private m_strDefaultAction As String
    'guid of the methodology assigned for this relationship (null = all)
    Private m_strMethodologyID As String
    'guid of the use assigned for this relationship (null = all)
    Private m_strUseID As String
    'form to open for treeview
    Private m_strForm As String
    'parent group
    Private m_lngGroupParent As Integer
    'master parent group
    Private m_lngMasterParent As Integer
    'master grandparent group
    Private m_lngMasterGrandParent As Integer
    'Default enabled
    Private m_fEnabled As Boolean



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
        'UPGRADE_WARNING: Couldn't resolve default property of object m_strTable. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        m_objRS = OpenDB(m_objConn, m_objRS, m_strTable, m_strSQL & " WHERE ID = " & lngID)

        'select the desired record
        'm_objRS.Filter = "ID = " & lngID
        For i = 0 To m_objRS.Tables(m_strTable).Rows.Count - 1
            If m_objRS.Tables(m_strTable).Rows(i).Item("ID") = lngID Then
                With m_objRS.Tables(m_strTable).Rows(i)
                    On Error GoTo ExitFunction
                    'copy the values into the class members
                    m_lngID = .Item("id") '.Fields("id").Value
                    m_lngParentID = IIf(IsDBNull(.Item("lngParentID")), 0, .Item("lngParentID"))
                    m_lngQuestionID = IIf(IsDBNull(.Item("ID_Question")), 0, .Item("ID_Question"))
                    m_strDescription = IIf(IsDBNull(.Item("txtDescription")), "", .Item("txtDescription"))
                    m_lngType = IIf(IsDBNull(.Item("lngType")), 0, .Item("lngType"))
                    m_lngGroupType = IIf(IsDBNull(.Item("lngGroupType")), 0, .Item("lngGroupType"))
                    m_lngTrueGroup = IIf(IsDBNull(.Item("lngTrueGroup")), 0, .Item("lngTrueGroup"))
                    m_lngFalseGroup = IIf(IsDBNull(.Item("lngFalseGroup")), 0, .Item("lngFalseGroup"))
                    m_strTreeType = IIf(IsDBNull(.Item("txtTreeType")), "", .Item("txtTreeType"))
                    m_strDefaultAction = IIf(IsDBNull(.Item("txtDefaultAction")), "", .Item("txtDefaultAction"))
                    'm_strUseID = IIf(IsDBNull(CType((.Item("guidID_Use")), Guid).ToString()), "", CType((.Item("guidID_Use")), Guid).ToString())
                    'm_strMethodologyID = IIf(IsDBNull(CType((.Item("guidID_Methodology")), Guid).ToString()), "", CType((.Item("guidID_Methodology")), Guid).ToString())

                    m_strUseID = .Item("GuidID_Use").ToString
                    m_strMethodologyID = .Item("guidID_Methodology").ToString

                    m_strForm = IIf(IsDBNull(.Item("txtForm")), "", .Item("txtForm"))
                    m_lngGroupParent = IIf(IsDBNull(.Item("groupParent")), 0, .Item("groupParent"))
                    m_lngMasterParent = IIf(IsDBNull(.Item("masterParent")), 0, .Item("masterParent"))
                    m_lngMasterGrandParent = IIf(IsDBNull(.Item("mastergrandParent")), 0, .Item("mastergrandParent"))
                    m_fEnabled = .Item("dftenabled") '.Fields("dftenabled").Value
                End With
                Exit For
            End If
        Next i

ExitFunction:
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
        Dim rst As New DataSet
        'connect to the db and get a recordset
        m_objConn = GetConnection(DB_DSN)
        m_objAdapter = New OleDb.OleDbDataAdapter(m_strSQL, m_objConn)
        m_objAdapter.Fill(rst, m_strTable)
        cb = New OleDb.OleDbCommandBuilder(m_objAdapter)

        With rst
            For i = 0 To .Tables(m_strTable).Rows.Count - 1
                If .Tables(m_strTable).Rows(i).Item("ID") = lngID Then
                    'stuff all of the record's properties
                    .Tables(m_strTable).Rows(i).Item("id") = m_lngID
                    .Tables(m_strTable).Rows(i).Item("lngParentID") = m_lngParentID
                    .Tables(m_strTable).Rows(i).Item("ID_Question") = m_lngQuestionID
                    .Tables(m_strTable).Rows(i).Item("txtDescription") = m_strDescription
                    .Tables(m_strTable).Rows(i).Item("lngType") = m_lngType
                    .Tables(m_strTable).Rows(i).Item("lngGroupType") = m_lngGroupType
                    .Tables(m_strTable).Rows(i).Item("lngTrueGroup") = m_lngTrueGroup
                    .Tables(m_strTable).Rows(i).Item("lngFalseGroup") = m_lngFalseGroup
                    .Tables(m_strTable).Rows(i).Item("txtTreeType") = m_strTreeType
                    .Tables(m_strTable).Rows(i).Item("txtDefaultAction") = m_strDefaultAction
                    .Tables(m_strTable).Rows(i).Item("guidID_Methodology") = New Guid(m_strMethodologyID)
                    .Tables(m_strTable).Rows(i).Item("guidID_Use") = New Guid(m_strUseID)
                    .Tables(m_strTable).Rows(i).Item("txtForm") = m_strForm
                    .Tables(m_strTable).Rows(i).Item("groupParent") = m_lngGroupParent
                    .Tables(m_strTable).Rows(i).Item("masterParent") = m_lngMasterParent
                    .Tables(m_strTable).Rows(i).Item("mastergrandParent") = m_lngMasterGrandParent
                    .Tables(m_strTable).Rows(i).Item("dftenabled") = m_fEnabled                    'write the record
                    m_objAdapter.Update(rst, m_strTable) '.Update()

                    Exit For
                End If
            Next
        End With

        CloseDB(m_objConn, m_objRS)

    End Function

    Public Function Delete(ByVal lngID As Integer) As Object
        ' Comments   : Removes the specified object from the repository
        ' Parameters : lngID - ID to delete
        ' Returns    :  -
        ' Modified   : 31-May-2002 LKB
        ' --------------------------------------------------------

        'UPGRADE_WARNING: Couldn't resolve default property of object m_strTable. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'm_objRS = OpenDB(m_objConn, m_objRS, m_strTable, m_strSQL)
        'point to the desired record
        'm_objRS.Filter = "ID =" & lngID
        'delete it
        'm_objRS.Delete((ADODB.AffectEnum.adAffectCurrent))
        'commit the changes
        'm_objRS.UpdateBatch()
        'CloseDB(m_objConn, m_objRS)

        Dim i As Integer
        'get a connection to the db
        m_objConn = GetConnection(DB_DSN)
        m_objAdapter = New OleDbDataAdapter(m_strSQL, m_objConn)
        m_objRS = New DataSet
        m_objAdapter.Fill(m_objRS, m_strTable)
        Dim cb As New OleDb.OleDbCommandBuilder(m_objAdapter)

        'load the data into memory
        For i = 0 To m_objRS.Tables(m_strTable).Rows.Count - 1
            If m_objRS.Tables(m_strTable).Rows(i).Item("ID") = lngID Then
                m_objRS.Tables(m_strTable).Rows(i).Delete()
            End If
        Next
        m_objAdapter.Update(m_objRS, m_strTable)
        CloseDB(m_objConn, m_objRS)
    End Function
	
    Public Function Exec(ByVal strStoredProc As String, ByRef aParams(,) As Object) As dataset
        ' Comments   : Executes the specified query, using the supplied list of parameters, and
        '              returns a query string
        ' Parameters : strStoredProc - Query to open
        '              aParams - Parameters for query
        ' Returns    : recordset based on query and parameters
        ' Modified   : 31-May-2002 LKB
        ' --------------------------------------------------------

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
	Public Function GetParentID() As Integer
		GetParentID = m_lngParentID
	End Function
	Public Function GetQuestionID() As Integer
		GetQuestionID = m_lngQuestionID
	End Function
	Public Function GetDescription() As String
		GetDescription = m_strDescription
	End Function
	'UPGRADE_NOTE: GetType was upgraded to GetType_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function GetType_Renamed() As Integer
		GetType_Renamed = m_lngType
	End Function
	Public Function GetGroupType() As Integer
		GetGroupType = m_lngGroupType
	End Function
	Public Function GetTrueGroup() As Integer
		GetTrueGroup = m_lngTrueGroup
	End Function
	Public Function GetFalseGroup() As Integer
		GetFalseGroup = m_lngFalseGroup
	End Function
	Public Function getTreeType() As String
		getTreeType = m_strTreeType
	End Function
	Public Function GetDefaultAction() As String
		GetDefaultAction = m_strDefaultAction
	End Function
	Public Function getMethodologyID() As String
		getMethodologyID = m_strMethodologyID
	End Function
	Public Function GetUseID() As String
		GetUseID = m_strUseID
	End Function
	Public Function GetForm() As String
		GetForm = m_strForm
	End Function
	Public Function GetGroupParent() As String
		GetGroupParent = CStr(m_lngGroupParent)
	End Function
	Public Function GetMasterParent() As String
		GetMasterParent = CStr(m_lngMasterParent)
	End Function
	Public Function GetMasterGrandParent() As String
		GetMasterGrandParent = CStr(m_lngMasterGrandParent)
	End Function
	Public Function GetDefaultEnabled() As Boolean
		GetDefaultEnabled = m_fEnabled
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
	Public Function SetParentID(ByRef lngParentID As Integer) As Short
		'set the member
		m_lngParentID = lngParentID
		'set the return code
		SetParentID = S_OK
	End Function
	Public Function SetQuestionID(ByRef lngQuestionID As Integer) As Short
		'set the member
		m_lngQuestionID = lngQuestionID
		'set the return code
		SetQuestionID = S_OK
	End Function
	Public Function SetDescription(ByRef strDescription As String) As Short
		'set the member
		m_strDescription = strDescription
		'set the return code
		SetDescription = S_OK
	End Function
	Public Function SetType(ByRef lngType As Integer) As Short
		'set the member
		m_lngType = lngType
		'set the return code
		SetType = S_OK
	End Function
	Public Function SetGroupType(ByRef lngGroupType As Integer) As Short
		'set the member
		m_lngGroupType = lngGroupType
		'set the return code
		SetGroupType = S_OK
	End Function
	Public Function SetTrueGroup(ByRef lngTrueGroup As Integer) As Short
		'set the member
		m_lngTrueGroup = lngTrueGroup
		'set the return code
		SetTrueGroup = S_OK
	End Function
	Public Function SetFalseGroup(ByRef lngFalseGroup As Integer) As Short
		'set the member
		m_lngFalseGroup = lngFalseGroup
		'set the return code
		SetFalseGroup = S_OK
	End Function
	Public Function SetTreeType(ByRef strTreeType As String) As Short
		'set the member
		m_strTreeType = strTreeType
		'set the return code
		SetTreeType = S_OK
	End Function
	Public Function SetDefaultAction(ByRef strDefaultAction As String) As Short
		'set the member
		m_strDefaultAction = strDefaultAction
		'set the return code
		SetDefaultAction = S_OK
	End Function
	Public Function setMethodologyID(ByRef strMethodologyID As String) As Short
		'set the member
		m_strMethodologyID = strMethodologyID
		'set the return code
		setMethodologyID = S_OK
	End Function
	Public Function SetUseID(ByRef strUseID As String) As Short
		'set the member
		m_strUseID = strUseID
		'set the return code
		SetUseID = S_OK
	End Function
	Public Function SetForm(ByRef strForm As String) As Short
		'set the member
		m_strForm = strForm
		'set the return code
		SetForm = S_OK
	End Function
	Public Function SetGroupParent(ByRef lngGroupParent As Integer) As Integer
		'set the member
		m_lngGroupParent = lngGroupParent
		'set the return code
		SetGroupParent = S_OK
	End Function
	Public Function SetMasterParent(ByRef lngMasterParent As Integer) As Integer
		'set the member
		m_lngMasterParent = lngMasterParent
		'set the return code
		SetMasterParent = S_OK
	End Function
	Public Function SetMasterGrandParent(ByRef lngMasterGrandParent As Integer) As Integer
		'set the member
		m_lngMasterGrandParent = lngMasterGrandParent
		'set the return code
		SetMasterGrandParent = S_OK
	End Function
	Public Function SetDefaultEnabled(ByRef fEnabled As Boolean) As Integer
		'set the member
		m_fEnabled = fEnabled
		'set the return code
		SetDefaultEnabled = S_OK
	End Function
End Class