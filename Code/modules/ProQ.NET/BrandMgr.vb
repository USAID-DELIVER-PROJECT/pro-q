Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("BrandMgr_NET.BrandMgr")> Public Class BrandMgr
	
	' Brand Manager
	' Controls the functions on the Backend for returning lists of Brands etc.
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'' Public Functions
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	Public Function GetCollection() As Collection
		
		Dim objBrandCol As New ProQdb.BrandCollectionDB
		Dim strProc As String
        Dim aParams(,) As Object
		
		strProc = "qlkpBrands"
		
        GetCollection = objBrandCol.GetCollection(strProc, aParams)
		
	End Function
	
	
	
	' Delete (From the DB)
	Public Function Delete(ByRef strID As String) As Short
		
        'Dim strProc As String
		Dim objBrand As New ProQdb.BrandDB
		
		objBrand.Load(strID)
		objBrand.Delete()
		
		'UPGRADE_NOTE: Object objBrand may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objBrand = Nothing
		
	End Function
End Class