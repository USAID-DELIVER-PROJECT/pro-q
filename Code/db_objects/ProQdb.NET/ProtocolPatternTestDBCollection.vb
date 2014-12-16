Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("ProtocolPatternTestDBCollection_NET.ProtocolPatternTestDBCollection")> Public Class ProtocolPatternTestDBCollection
	'ProtocolTestCollectionDB.cls
	'
	'manages groups of records
	'
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         private members
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	'the dsn used to connect to this db/table
    Private m_strDSN As String = DB_DSN
	
	'whether or not the object has been constructed
    Private m_objConn As OleDbConnection
	Private m_objDBExec As New DBExec
	
	'the collection of FundingSources
	Private m_colDB As New Collection
	
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
		'UPGRADE_NOTE: Object m_objDBExec may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_objDBExec = Nothing
		
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
        Dim objRS As DataSet
        Dim i As Integer
        m_objConn = GetConnection(m_strDSN)
        If aParams Is Nothing Then
            objRS = m_objDBExec.ReturnRSNoParams(strStoredProc, aParams)
        Else
            objRS = m_objDBExec.ReturnRS(m_objConn, strStoredProc, aParams)
        End If

        'for each record
        With objRS
            For i = 0 To .Tables(strStoredProc).Rows.Count - 1 'Do Until .EOF
                Dim objDB As New ProtocolPatternTestDB
                'put the object in the collection

                'stuff all of the record's properties into the object
                objDB.Load(CType((.Tables(strStoredProc).Rows(i).Item("GuidID")), Guid).ToString()) '.Fields("guidID").Value)

                'add the object to the collection
                m_colDB.Add(objDB)
                'UPGRADE_NOTE: Object objDB may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                objDB = Nothing

                'get the next record
            Next
        End With

        GetCollection = m_colDB

        'UPGRADE_NOTE: Object objRS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        objRS = Nothing
        'UPGRADE_NOTE: Object m_objConn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        m_objConn = Nothing

    End Function
End Class