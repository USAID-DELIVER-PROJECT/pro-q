Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("UseLookup_NET.UseLookup")> Public Class UseLookup
	
	'Lookup functions for Displaying Uses.
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                            constants
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	'result codes
	Private Const S_OK As Short = 0
	Private Const strTable As String = "tlkUse"
	
	'error strings
	Private Const ERROR_CREATE_1 As String = "Error: You need to provide input values."
	
	'whether or not the object has been constructed
    Private m_objConn As New OleDbConnection
    Private m_objRS As DataSet
    'sp & recordset functions
    Private m_objDBExec As New DBExec

	Public Function getUseName(ByRef strID As String) As String
		
        Dim aParams(1, 1) As Object

        'Dim objCommand As New ADODB.Command
        'Dim prmParam As ADODB.Parameter
		'associate the command with the conn
		
        'm_objConn.ConnectionString = G_STRDSN
        'm_objConn.Open()
		
        'objCommand.let_ActiveConnection(m_objConn)
		'declare that it's a stored proc
        'objCommand.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
		'give it a timeout (30 sec is the default...)
        'objCommand.CommandTimeout = DB_COMMAND_TIMEOUT
		'name the sp
        'objCommand.CommandText = "qlkpUseName"
		'create the sp's parameter list
		
        'prmParam = New ADODB.Parameter
		'prmParam.Name = "strName"
        '     prmParam.Type = DbType.Guid
        'prmParam.Direction = ADODB.ParameterDirectionEnum.adParamInput
        'prmParam.Value = strID
        'objCommand.Parameters.Append(prmParam)
        aParams(0, 0) = New Guid(strID)
        aParams(0, 1) = DbType.Guid

        'm_objRS = objCommand.Execute()
        m_objRS = m_objDBExec.ReturnRS(m_objConn, "qlkpUseName", aParams)

        'get the id
		'Debug.Print m_objRS!guidID
		
        getUseName = m_objRS.Tables("qlkpUseName").Rows(0).Item("strName")
		
        'm_objRS.Close()
		'm_objConn.Close
		
	End Function
End Class