Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("DemandCalculator_NET.DemandCalculator")> Public Class DemandCalculator
	'DemandCalculator.cls
	
	Dim m_colNeeds As Collection
	Dim m_colFormulas As Collection
	
	Private Const SP_TABLE_NEED As String = "qlkpNeeds"
	Private Const SP_TABLE_FORMULA As String = "qlkpFormulas"
	
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Public Functions
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Function GetNeedSQLString(ByRef strUSE As String, ByRef strMethodology As String) As String
		
		Dim objNeedDB As ProQdb.NeedDB
        Dim strSQL As String = ""
		
		
        For Each objNeedDB In m_colNeeds
            If objNeedDB.getMethodologyID = strMethodology And objNeedDB.GetUseID = strUSE Then
                strSQL = objNeedDB.GetFormula
                Exit For
            End If
        Next objNeedDB
		
		GetNeedSQLString = strSQL
		
	End Function
	Function GetFormulaSQLString(ByRef strUSE As String, ByRef lngSection As Integer, ByRef fElisa As Boolean) As String
		
		Dim objFormulaDB As ProQdb.FormulaDB
		Dim strSQL As String
		
		For	Each objFormulaDB In m_colFormulas
			If objFormulaDB.GetUseID = strUSE And objFormulaDB.getSection = lngSection And objFormulaDB.getIsElisa = fElisa Then
				strSQL = objFormulaDB.GetFormula
				Exit For
			End If
		Next objFormulaDB
		
		GetFormulaSQLString = strSQL
		
	End Function
	
	
	Function ReplaceSQLString(ByRef strSQL As Object, ByRef aValues() As Object) As String
		' Comments  : Take a passed in SQL string and an array for each value
		'           : in the placeholder in the string with the value for that
		'           : Array /%1/, /%2/. etc.
		' Parameters: strSQL - the SQL statement
		'           : aValues - the values to replace
		' Returns   : String - The Modified String
		' Created   : 18-June-2002 jleiner
		'------------------------------------------------------------------------
		Dim i As Short
		Dim x As Short
		Dim strTemp As String
		
		'UPGRADE_WARNING: Couldn't resolve default property of object strSQL. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		strTemp = strSQL
        strTemp = Replace(strTemp, "'/", "{/")
        strTemp = Replace(strTemp, "/'", "/}")
        For i = 0 To UBound(aValues) - 1
            x = i + 1
            'UPGRADE_WARNING: Couldn't resolve default property of object aValues(i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            strTemp = Replace(strTemp, "/%" & CStr(x) & "/", aValues(i))
        Next
		
		ReplaceSQLString = strTemp
		
	End Function
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Private Functions
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Private Sub GetNeeds()
		
		Dim objNeedColDB As New ProQdb.NeedDBCollection
		Dim strProc As String
        Dim aParams(,) As Object
		
		strProc = SP_TABLE_NEED
		
		m_colNeeds = objNeedColDB.GetCollection(strProc, aParams)
		'UPGRADE_NOTE: Object objNeedColDB may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objNeedColDB = Nothing
		
	End Sub
	Private Sub GetFormulas()
		
		Dim objFormulaColDB As New ProQdb.FormulaDBCollection
		Dim strProc As String
        Dim aParams(,) As Object
		
		strProc = SP_TABLE_FORMULA
		
		m_colFormulas = objFormulaColDB.GetCollection(strProc, aParams)
		'UPGRADE_NOTE: Object objFormulaColDB may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objFormulaColDB = Nothing
		
	End Sub
	
	
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Class Functions
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		GetNeeds()
		GetFormulas()
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object m_colNeeds may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_colNeeds = Nothing
		'UPGRADE_NOTE: Object m_colFormulas may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_colFormulas = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
End Class