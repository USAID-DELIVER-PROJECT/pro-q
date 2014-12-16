Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("LogisticsConsDBCollection_NET.LogisticsConsDBCollection")> Public Class LogisticsConsDBCollection
	'LogisticsConsiderationsDBCollection.cls
	'
	'provides collections of LogisticsConsiderationsDB Objects
	'
	'lbailey
	'8 june 2002
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                            constants
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	'all in constants.bas
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         private members
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	'the dsn used to connect to this db/table
    Private m_strDSN As String = DB_DSN
	
	'connection to the db
    Private m_objConn As OleDbConnection
	'sp & recordset functions
	Private m_objDBExec As New DBExec
	
	'the collection of quantities
	Private m_colLogDB As New Collection
	
	
	
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
		'Set m_objConn = Nothing
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
	Public Function GetCollection(ByRef strStoredProc As String) As Collection
		
		'go get the recordset
        Dim rst As DataSet
        Dim i As Integer

        m_objConn = GetConnection(m_strDSN)
        rst = m_objDBExec.ReturnRSfromSQL(strStoredProc, "LogisticsConsideration")
		
		'for each record
        With rst
            For i = 0 To .Tables("LogisticsConsideration").Rows.Count - 1 'Do Until .EOF
                Dim objDB As New LogisticsConsDB
                'put the object in the collection

                'stuff all of the record's properties into the object
                objDB.setQuantificationID(CType((.Tables("LogisticsConsideration").Rows(i).Item("guidID_Quantification")), Guid).ToString()) '.Fields("guidID_Quantification").Value)
                'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                objDB.setBrandID(IIf(IsDBNull(CType((.Tables("LogisticsConsideration").Rows(i).Item("guidID_Type")), Guid).ToString()), "", CType((.Tables("LogisticsConsideration").Rows(i).Item("guidID_Type")), Guid).ToString()))
                objDB.setLeadTime(.Tables("LogisticsConsideration").Rows(i).Item("Leadtime")) '.Fields("Leadtime"))
                objDB.setBufferStock(.Tables("LogisticsConsideration").Rows(i).Item("BufferStock")) '.Fields("BufferStock"))
                objDB.setShipments(.Tables("LogisticsConsideration").Rows(i).Item("ShipmentS")) '.Fields("ShipmentS"))
                objDB.setStorageCapacity(.Tables("LogisticsConsideration").Rows(i).Item("StorageCapacity")) '.Fields("StorageCapacity"))
                objDB.setColdStorage(.Tables("LogisticsConsideration").Rows(i).Item("ColdStorage")) '.Fields("ColdStorage"))
                objDB.setQtyOnHand(.Tables("LogisticsConsideration").Rows(i).Item("qtyonhand")) '.Fields("qtyonhand"))
                objDB.setQtyOnOrder(.Tables("LogisticsConsideration").Rows(i).Item("qtyonOrder")) '.Fields("qtyonOrder"))

                'add the object to the collection
                m_colLogDB.Add(objDB)
                'UPGRADE_NOTE: Object objDB may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                objDB = Nothing
                'get the next record
            Next
        End With

		GetCollection = m_colLogDB
		'UPGRADE_NOTE: Object rst may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rst = Nothing
		'UPGRADE_NOTE: Object m_objConn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_objConn = Nothing
	End Function
End Class