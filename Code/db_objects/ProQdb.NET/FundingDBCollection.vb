Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("FundingDBCollection_NET.FundingDBCollection")> Public Class FundingDBCollection
	'FundingDBCollection.cls
	'
	'manages groups of records
	'
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                            constants
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         private members
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	
	'the dsn used to connect to this db/table
    Private m_strDSN As String = DB_DSN
	'the name of the table or view used by this class
	Private Const m_strTable As String = DB_TABLE_FUNDING
	
	'whether or not the object has been constructed
    Private m_objConn As OleDbConnection
    'Private m_objRS As DataSet 
	Private m_objDBExec As New DBExec
	
	'the collection of Fundings
	Private m_cFundings As New Collection
	
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
	'1 june 2002
	'
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		
		'get a connection to the db
		'Set m_objConn = GetConnection(m_strDSN)
		
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
	'1 june 2002
	'
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		
		'release the connections to the db
		'Set m_objConn=Nothing
		
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
	'GetCollection()
	'
	'executes the specified stored procedure and returns the results in a
	'collection of objects
	'
	'lbailey
	'1 june 2002
	'-
    Public Function GetCollection(ByRef strStoredProc As String, ByRef aParams(,) As Object) As Collection

        On Error Resume Next

        'go get the recordset
        Dim rst As DataSet
        Dim i As Integer
        'UPGRADE_WARNING: Couldn't resolve default property of object m_strDSN. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        m_objConn = GetConnection(m_strDSN)
        If aParams Is Nothing Then
            rst = m_objDBExec.ReturnRSNoParams(strStoredProc, aParams)
        Else
            rst = m_objDBExec.ReturnRS(m_objConn, strStoredProc, aParams)
        End If


        'for each record
        With rst
            For i = 0 To .Tables(strStoredProc).Rows.Count - 1 'Do Until .EOF
                Dim objFundingDB As New FundingDB
                'put the object in the collection


                'stuff all of the record's properties into the object
                objFundingDB.Load(CType((.Tables(strStoredProc).Rows(i).Item("GuidID")), Guid).ToString()) 'objFundingDB.Load(.Fields("guidID").Value)

                'add the object to the collection
                m_cFundings.Add(objFundingDB)
                'UPGRADE_NOTE: Object objFundingDB may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                objFundingDB = Nothing

                'get the next record
                '.MoveNext()
            Next 'Loop
            '.Close()
        End With

        GetCollection = m_cFundings
        'UPGRADE_NOTE: Object rst may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        rst = Nothing
        'UPGRADE_NOTE: Object m_objConn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        m_objConn = Nothing
    End Function
End Class