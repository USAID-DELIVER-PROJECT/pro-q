Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("ResponseDB_NET.ResponseDB")> Public Class ResponseDB
	'Response.cls
	'Response.cls
	'Manages the responses to the questions
	'31-May-2002 lblanken
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         private members
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	'the dsn used to connect to this db/table
    Private m_strDSN As String = DB_DSN
	'the name of the table or view used by this class
	Private Const m_strTable As String = DB_TABLE_RESPONSE
    Private Const m_strSQL As String = "Select * FROM " & DB_TABLE_RESPONSE

	'whether or not the object has been constructed
    Private m_objConn As OleDbConnection
    Private m_objRS As DataSet
	Private m_fIsNew As Boolean
	
	
	'the guid of the object
	Private m_strID As String
	'the guid of the quantification associated with
	Private m_strQuantificationID As String
	'the guid of the aggrgation associated with
	Private m_strAggregationID As String
	'the id of the ScriptRelationship assocatied with
	Private m_lngScriptRelationshipID As Integer
	'the answer inputed by the user
	Private m_varAnswer As Object
	'the status of the response (0=Not Complete, 1=In Progress, 2=Complete)
	Private m_lngStatus As Integer
	'the quid of the referenced type (i.e. BrandID, etc.)
	Private m_strTypeID As String
	'whether or not the response in enabled on the treeview/script
	Private m_fEnabled As Boolean
	'the suborder for the treeview
	Private m_lngSubActionID As Integer
	'the id for treeview
	Private m_lngTreeviewID As Integer
	'free text
	Private m_strNotes As String
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

        m_objRS = OpenDB(m_objConn, m_objRS, m_strTable, m_strSQL & " WHERE lngTreeviewID = " & lngID)

        'select the desired record
        'm_objRS.Filter = "lngTreeviewID = " & lngID
        For i = 0 To m_objRS.Tables(m_strTable).Rows.Count - 1
            ID = m_objRS.Tables(m_strTable).Rows(i).Item("lngTreeviewID")
            If ID = lngID Then
                With m_objRS.Tables(m_strTable).Rows(i)
                    'copy the values into the class members
                    m_strID = .Item("GuidID").ToString() '.Fields("guidID").Value
                    m_lngTreeviewID = .Item("lngtreeviewid") '.Fields("lngtreeviewid").Value
                    m_strQuantificationID = .Item("guidID_Quantification").ToString() '.Fields("guidID_Quantification").Value
                    m_strAggregationID = .Item("guidID_Aggregation").ToString() '.Fields("guidID_Aggregation").Value
                    m_lngScriptRelationshipID = .Item("ID_ScriptRelationship") '.Fields("ID_ScriptRelationship").Value
                    m_varAnswer = IIf(IsDBNull(.Item("dblAnswer")), System.DBNull.Value, .Item("dblAnswer"))
                    m_lngStatus = .Item("lngStatus") '.Fields("lngStatus").Value
                    m_strTypeID = IIf(IsDBNull((.Item("guidID_Type")).ToString()), "", (.Item("guidID_Type")).ToString())
                    m_fEnabled = .Item("fEnabled") '.Fields("fEnabled").Value
                    m_lngSubActionID = .Item("lngSubActionID") '.Fields("lngSubActionID").Value
                    m_strNotes = IIf(IsDBNull(.Item("memNotes")), "", .Item("memNotes"))
                End With

                'remember that this record already exists
                m_fIsNew = False
                Exit For
            End If
        Next i

        'Jleiner, moved Out of Loop
        CloseDB(m_objConn, m_objRS)

    End Function

    Public Function Create() As String
        ' Comments   : Adds the New Record to the db using the info that we've got
        ' Parameters :  -
        ' Returns    :  -
        ' Modified   : 31-May-2002 LKB
        ' --------------------------------------------------------

        'create the id and set the defaults
        'UPGRADE_WARNING: Couldn't resolve default property of object g_objGUIDGenerator.GetGUID(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        m_strID = g_objGUIDGenerator.GetGUID()

        'remember that this is a new record
        m_fIsNew = True

        Create = m_strID


    End Function

    Public Function Update() As Object
        ' Comments   : Updates the Record to the db using the info that we've got
        ' Parameters : -
        ' Returns    : -
        ' Modified   : 31-May-2002 LKB
        ' --------------------------------------------------------
        Dim i As Integer
            Dim cb As OleDb.OleDbCommandBuilder
            Dim dsNewRow As DataRow
            Dim objAdapter As OleDbDataAdapter
            'connect to the db and get a recordset
            'OpenDB(m_objConn, m_objRS, m_strTable, m_strSQL)
        Try
            m_objConn = GetConnection(DB_DSN)
            m_objRS = New DataSet
            objAdapter = New OleDb.OleDbDataAdapter(m_strSQL, m_objConn)
            objAdapter.Fill(m_objRS, m_strTable)
            cb = New OleDb.OleDbCommandBuilder(objAdapter)

            If m_fIsNew Then

                dsNewRow = m_objRS.Tables(m_strTable).NewRow()

                'stuff all of the record's properties
                dsNewRow.Item("guidID") = New Guid(m_strID)
                dsNewRow.Item("guidID_Aggregation") = New Guid(m_strAggregationID)
                'cant set guid to null so need to do check and set only if needed.
                If Len(m_strQuantificationID) > 0 Then
                    dsNewRow.Item("guidID_Quantification") = New Guid(m_strQuantificationID) 'IIf(Len(m_strQuantificationID) > 0, New Guid(m_strQuantificationID), System.DBNull.Value)
                End If
                dsNewRow.Item("ID_ScriptRelationship") = m_lngScriptRelationshipID
                'cant set dbl to empty string
                dsNewRow.Item("dblAnswer") = IIf(Len(m_varAnswer) > 0, m_varAnswer, DBNull.Value)
                dsNewRow.Item("lngStatus") = m_lngStatus
                dsNewRow.Item("fEnabled") = m_fEnabled
                dsNewRow.Item("lngSubActionID") = m_lngSubActionID
                dsNewRow.Item("memNotes") = m_strNotes

                If m_strTypeID <> "" Then
                    dsNewRow.Item("guidID_Type") = New Guid(m_strTypeID)
                End If
                m_objRS.Tables(m_strTable).Rows.Add(dsNewRow)

                objAdapter.Update(m_objRS, m_strTable)
                m_fIsNew = False
            Else
                With m_objRS
                    For i = 0 To .Tables(m_strTable).Rows.Count - 1
                        Dim strGuid As String
                        strGuid = CType(.Tables(m_strTable).Rows(i).Item(DB_TABLE_PK), Guid).ToString
                        If strGuid = m_strID Then
                            'stuff all of the record's properties
                            .Tables(m_strTable).Rows(i).Item("guidID") = New Guid(m_strID)
                            .Tables(m_strTable).Rows(i).Item("guidID_Aggregation") = New Guid(m_strAggregationID)
                            .Tables(m_strTable).Rows(i).Item("guidID_Quantification") = IIf(Len(m_strQuantificationID) > 0, New Guid(m_strQuantificationID), System.DBNull.Value)
                            .Tables(m_strTable).Rows(i).Item("ID_ScriptRelationship") = m_lngScriptRelationshipID
                            If IsDBNull(m_varAnswer) Then
                                .Tables(m_strTable).Rows(i).Item("dblAnswer") = System.DBNull.Value
                            Else
                                .Tables(m_strTable).Rows(i).Item("dblAnswer") = CDbl(m_varAnswer)
                            End If

                            .Tables(m_strTable).Rows(i).Item("lngStatus") = CByte(m_lngStatus)
                            .Tables(m_strTable).Rows(i).Item("fEnabled") = m_fEnabled
                            .Tables(m_strTable).Rows(i).Item("lngSubActionID") = m_lngSubActionID
                            .Tables(m_strTable).Rows(i).Item("memNotes") = m_strNotes
                            .Tables(m_strTable).Rows(i).Item("lngtreeviewid") = m_lngTreeviewID '.Fields("lngtreeviewid").Value

                            If m_strTypeID <> "" Then
                                .Tables(m_strTable).Rows(i).Item("guidID_Type") = New Guid(m_strTypeID)
                            End If
                            'write the record
                            objAdapter.Update(m_objRS, m_strTable) '.Update()
                            Exit For
                        End If
                    Next
                End With
            End If
        Catch ex As OleDbException
            MsgBox(ex.Message, vbCritical, "Ole db Error: " & ex.ErrorCode)
        Catch ex As Exception

            MsgBox(ex.Message)

        Finally
            'clean up
            CloseDB(m_objConn, m_objRS)
        End Try
    End Function

    Public Function Delete(ByVal strID As String) As Object
        ' Comments   : Removes the specified object from the repository
        ' Parameters : lngID - ID to delete
        ' Returns    :  -
        ' Modified   : 31-May-2002 LKB
        ' --------------------------------------------------------

        'OpenDB(m_objConn, m_objRS, m_strTable, m_strSQL)
        DeleteRecord(m_objConn, m_objRS, m_strTable, m_strSQL, strID)
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
        m_objRS = m_objDBExec.ReturnRS(m_objConn, strStoredProc, aParams)

        'set filter based on id
        Load(CInt(m_strID))

        Exec = m_objRS

    End Function
	
	
	'-----------------------------------------------------------------
	'Standard Accessor Functions
	' Modified   : 31-May-2002 LKB
	'-----------------------------------------------------------------
	
	Public Function GetID() As String
		GetID = m_strID
	End Function
	Public Function GetTreeviewID() As Integer
		GetTreeviewID = m_lngTreeviewID
	End Function
	Public Function getQuantificationID() As String
		getQuantificationID = m_strQuantificationID
	End Function
	
	Public Function GetAggregationID() As String
		GetAggregationID = m_strAggregationID
	End Function
	
	Public Function GetScriptRelationshipID() As Integer
		GetScriptRelationshipID = m_lngScriptRelationshipID
	End Function
	
	Public Function GetAnswer() As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varAnswer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object GetAnswer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetAnswer = m_varAnswer
	End Function
	
	Public Function GetStatus() As Integer
		GetStatus = m_lngStatus
	End Function
	
	Public Function GetTypeID() As String
		GetTypeID = m_strTypeID
	End Function
	
	Public Function GetEnabled() As Boolean
		GetEnabled = m_fEnabled
	End Function
	Public Function GetSubActionID() As Integer
		GetSubActionID = m_lngSubActionID
	End Function
	Public Function GetNotes() As String
		GetNotes = m_strNotes
	End Function
	
	'-----------------------------------------------------------------
	'Standard Manipulator Functions
	' Modified   : 31-May-2002 LKB
	'-----------------------------------------------------------------
	Public Function SetID(ByRef strID As String) As Short
		'set the member
		m_strID = strID
		'set the return code
		SetID = S_OK
	End Function
	
	Public Function setQuantificationID(ByRef strQuantificationID As String) As Short
		'set the member
		m_strQuantificationID = strQuantificationID
		'set the return code
		setQuantificationID = S_OK
	End Function
	
	Public Function SetAggregationID(ByRef strAggregationID As String) As Short
		'set the member
		m_strAggregationID = strAggregationID
		'set the return code
		SetAggregationID = S_OK
	End Function
	
	Public Function SetScriptRelationshipID(ByRef lngScriptRelationshipID As Integer) As Short
		'set the member
		m_lngScriptRelationshipID = lngScriptRelationshipID
		'set the return code
		SetScriptRelationshipID = S_OK
	End Function
	
	Public Function SetAnswer(ByRef varAnswer As Object) As Short
		'set the member
		'UPGRADE_WARNING: Couldn't resolve default property of object varAnswer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object m_varAnswer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If IsDBNull(varAnswer) Then
            m_varAnswer = System.DBNull.Value
        Else
            m_varAnswer = System.Math.Round(varAnswer, 3)
        End If
        'set the return code
		SetAnswer = S_OK
	End Function
	
	Public Function SetStatus(ByRef lngStatus As Integer) As Short
		'set the member
		m_lngStatus = lngStatus
		'set the return code
		SetStatus = S_OK
	End Function
	
	Public Function SetTypeID(ByRef strTypeID As String) As Short
		'set the member
		m_strTypeID = strTypeID
		'set the return code
		SetTypeID = S_OK
	End Function
	
	Public Function SetEnabled(ByRef fEnabled As Boolean) As Short
		'set the member
		m_fEnabled = fEnabled
		'set the return code
		SetEnabled = S_OK
	End Function
	Public Function SetSubActionID(ByRef lngSubActionID As Integer) As Short
		'set the member
		m_lngSubActionID = lngSubActionID
		'set the return code
		SetSubActionID = S_OK
	End Function
	Public Function SetNotes(ByRef strNotes As String) As Short
		'set the member
		m_strNotes = strNotes
		'set the return code
		SetNotes = S_OK
	End Function
End Class