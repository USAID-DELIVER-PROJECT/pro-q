Option Strict Off
Option Explicit On
'UPGRADE_NOTE: Constants was upgraded to Constants_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
Module Constants_Renamed
	'+
	'Constants.bas
	'
	'place for putting all of the constants used by the middle tier
	'classes.
	'
	'lbailey
	'7 june 2002
	'-
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                            result codes
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	Public Const S_OK As Short = 0
	Public Const S_FAILED As Short = -1
	Public Const S_FAILURE As Short = -1 'referenced in Code so I added here
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                            queries
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Public Const SP_GET_KITS_BY_BRAND As String = "qlkpKitByBrand"
	Public Const SP_GET_PROTOCOLPATTERNS As String = "qlkpProtocolPatterns"
	Public Const SP_GET_PROTPATLEV_BY_PROTPAT_ID As String = "qselProtPattLevByProtPattID"
	Public Const SP_GET_PROTOCOLLEVELS_BY_PROTOCOLID As String = "qselProtLevelsByProtID"
	Public Const SP_GET_PROTOCOL_BY_QUANTID As String = "qselProtByQuantID"
	Public Const SP_GET_PROTOCOLTESTS_BY_LEVEL As String = "qselProtocolTestsByLevelID"
	Public Const SP_GET_PATTERNTEST_BY_LEVEL As String = "qselPatternTestsByLevelID"
	Public Const SP_GET_CHILD_PROTOCOLTESTS As String = "qselProtocolTestsByProtocolID"
	Public Const SP_GET_PROTOCOLBRANDS_BY_TEST As String = "qselProtocolBrandsByProtocolTestID"
	Public Const SP_GET_QUANTITY_BY_SELBRANDID_CATEGORY As String = "qselQuantityBySelbrandAndCategory"
	Public Const SP_GET_QUANTITY_BY_SELBRAND As String = "qselQuantityBySelbrand"
	Public Const SP_GET_SELECTEDBRAND_BY_ID As String = "qrySelectedBrandByID"
	Public Const SP_GET_SELECTEDBRAND_BY_QUANT As String = "qselSelectedBrandsByQuantification"
	Public Const SP_SELECTED_BRANDS_BY_QUANTIFICATION As String = SP_GET_SELECTEDBRAND_BY_QUANT
	
	Public Const SP_GET_QUANTIFICATIONS_BY_AGGREGATION As String = "qselQuantificationsByAggregation"
	Public Const SP_FUNDINGSOURCE_BY_AGGREGATION As String = "qselFundingSourceByAggregation"
	
	Public Const SP_QSELBRANDINUSE As String = "qselBrandInUse"
	Public Const SP_QSELKITINUSE As String = "qselKitInUse"
	
	'Public Const SP_GET_CHILD_PROTOCOLBRANDS = ""
	
	Public Const SP_COSTS_BY_QUANTIFICATION As String = "qselCostsbyQuantification"
	Public Const SP_QUANTIFICATION_METHODOLOGIES As String = "qselQuantificationMethodology"
	Public Const SP_DELETE_SCRIPT_BY_QUANTMETHOD As String = "qdelScriptMethodology"
	Public Const SP_DELETE_SCRIPT_BY_TYPE As String = "qdelScriptType"
	Public Const SP_DELETE_SCRIPT_BY_TYPEID As String = "qdelScriptTypeByID"
	Public Const SP_UPDATE_SCRIPT_BY_TYPEID As String = "qupdScriptTypeByID"
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                            Code Constants
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Public Const TYPE_KITS As Short = 1
	Public Const TYPE_CUSTOMS As Short = 2
	Public Const TYPE_STORAGE As Short = 3
	
	Public Const NO_INPUT As Short = -1
	
	Public Const PCT_AIDS_PREVALENCE As Double = 17#
	Public Const PCT_DISCORDANCE As Double = 1.5
	
	Public Const SC_SAMPLES As Short = 0
	Public Const SC_TESTS As Short = -1
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                            DataTypes?
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                            error messages
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	
	Public Const ERROR_INVALID_BRAND As String = "Invalid brand."
	Public Const ERROR_INVALID_PERCENTAGE As String = "The percentages total greater than 100%"
	
	Public Const ERROR_SB_LOAD As String = "Could not load selected brands. Bad data."
	Public Const ERROR_SB_CANT_SUBTRACT As String = "Cannot have quantities less than zero."
End Module