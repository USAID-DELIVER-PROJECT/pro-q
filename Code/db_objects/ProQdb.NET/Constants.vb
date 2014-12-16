Option Strict Off
Option Explicit On
'UPGRADE_NOTE: Constants was upgraded to Constants_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
Module Constants_Renamed
	'+
	'Constants.bas
	'
	'keeper of all of the constants used in the backend
	'
	'lbailey
	'2 june 2002
	'-
	
	
	'result codes
	Public Const S_OK As Short = 0
	Public Const S_FAILURE As Short = -1
	
    'db connection  MOVED TO GLOBALS
    'Public DB_DSN As String = ProQdb.My.MySettings.Default.JSI_ProQ_ConnectionString ''"PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source = F:\SDG\dev\Deliver\proq\ProQ.NET\code\db\ProQ_db.mdb;" '"DSN=JSI_ProQNew"
    'Public G_STRDSN As String = ProQdb.My.MySettings.Default.JSI_ProQ_ConnectionString '"PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source = F:\SDG\dev\Deliver\proq\ProQ.NET\code\db\ProQ_db.mdb;" '"DSN=JSI_ProQNew"



    'standard column names
	Public Const DB_COL_PK As String = "guidID"
	Public Const DB_TABLE_PK As String = "guidID"
	
	'COmmand Object Property Constansts
	Public Const DB_COMMAND_TIMEOUT As Short = 0
	
	
	'lookup tables
	Public Const DB_TABLE_BRAND As String = "tlkBrand"
	Public Const DB_TABLE_CARTON As String = "tlkKitCarton"
	Public Const DB_TABLE_COUNTRY As String = "tlkCountry"
	Public Const DB_TABLE_KIT As String = "tlkKit"
    Public Const DB_TABLE_METHODOLOGY As String = "tlkMethodology"
	Public Const DB_TABLE_PROTOCOL_PATTERN As String = "tlkPattern"
	Public Const DB_TABLE_PROTOCOL_PATTERN_LEVEL As String = "tlkPatternLevel"
	Public Const DB_TABLE_PROTOCOL_PATTERN_TEST As String = "tlkPatternTest"
	Public Const DB_TABLE_QUESTION As String = "tlkQuestion"
	Public Const DB_TABLE_USE As String = "tlkUse"
    Public Const DB_TABLE_SCRIPT_RELATIONSHIP As String = "tlkScriptRelationship"
    Public Const DB_TABLE_PROGRAM As String = "tlkProgram"
    Public Const DB_TABLE_QUANTIFICATIONMETHOD As String = "tblQuantMethod"
	'data
	Public Const DB_TABLE_AGGREGATION As String = "tblAggregation"
	Public Const DB_TABLE_PROTOCOL As String = "tblProtocol"
	Public Const DB_TABLE_PROTOCOL_BRAND As String = "tblProtocolBrand"
	Public Const DB_TABLE_PROTOCOL_LEVEL As String = "tblProtocolLevel"
	Public Const DB_TABLE_PROTOCOL_STEP As String = "tblProtocolStep"
	Public Const DB_TABLE_PROTOCOL_TEST As String = "tblProtocolTest"
	Public Const DB_TABLE_QUANTIFICATION As String = "tblQuantification"
	Public Const DB_TABLE_QUANTITY As String = "tblQuantity"
	Public Const DB_TABLE_RESPONSE As String = "tblResponse"
	Public Const DB_TABLE_SELECTED_BRAND As String = "tblSelectedBrand"
	
	Public Const DB_TABLE_COST As String = "tblCost"
	Public Const DB_TABLE_FUNDINGSOURCE As String = "tblFundingSource"
	Public Const DB_TABLE_FUNDING As String = "tblFunding"
	Public Const DB_TABLE_USEFUNDING As String = "tblUseFunding"
	
	'error messages
	Public Const ERROR_INVALID_INPUT As String = "Invalid input"
	Public Const ERROR_CREATE_1 As String = "Input values missing"
End Module