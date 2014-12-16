Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("ProtocolMgr_NET.ProtocolMgr")> Public Class ProtocolMgr
	'ProtocolMgr
	'
	'manages the set of protocols in the db.  this class will need much
	'more work later on (adding new protocols, modifying aspects of each)
	'but for now, we just need it to get the list of protocols so that we
	'can gain access to the one we want for this quantification.
	'
	'lbailey
	'31 may 2002
	'
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'constants
	
	Private Const SP_PROTOCOL_LOAD_ALL As String = "sp_Protocol_Load_All"
	Private Const SP_PROTOCOL_LOAD_QUANTIFICATION As String = "sp_Protocol_Load_Quantification"
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'private properties
	
	Private m_objProtocol As New Protocol
	Private m_cProtocols As Collection
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'public properties
	
	'none
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' private methods
	
	'constructor
	
	'destructor
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' public methods
	
	'Load()
	'
	'Gets a list of all protocols into this object.
	'
	'lbailey
	'31 may 2002
	'
	Public Sub LoadAll()
		
		'set stored procedure
		Dim strSP As String
		strSP = SP_PROTOCOL_LOAD_ALL
		
		
		'set up param list
		Dim aParams(1, 1) As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object aParams(0, 0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		aParams(0, 0) = "*"
		'UPGRADE_WARNING: Couldn't resolve default property of object aParams(0, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(0, 1) = DbType.Guid
		
		'connect to the db
		Dim objProtocolDBCollection As ProQdb.ProtocolCollectionDB
		'create a collection to hold the protocoldb objects
		Dim cProtDBs As Collection
		Dim objProtocolDB As ProQdb.ProtocolDB
		
		'get all of the protocoldb objects
		cProtDBs = objProtocolDBCollection.GetCollection(strSP, aParams)
		'now go create all of the protocol objects
		Dim objProtocol As New Protocol
		For	Each objProtocolDB In cProtDBs
			'UPGRADE_WARNING: Couldn't resolve default property of object objProtocolDB. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            objProtocol.Load(CStr(objProtocolDB.GetID()))
		Next objProtocolDB
		
	End Sub
End Class