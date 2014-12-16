Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("KitMgr_NET.KitMgr")> Public Class KitMgr
	'+
	'KitMgr.cls
	'
	'this class manages lists of kits associated with brands
	'
	'lbailey
	'8 june 2002
	'-
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'private properties
	
	
	'array of brands information, id, count, and pct
	Private m_cKits As New Collection
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' private methods
	
	
	'+
	'FindKitBySize()
	'
	'Returns the kit
	'
	'lbailey
	'1 July 2002
	'-
	Public Function FindKitBySize(ByRef nSize As Short) As Kit
		
		Dim objFoundKit As New Kit
		
		Dim objKit As Kit
		If (m_cKits.Count() > 0) Then
			
			'iterate through the collection
			For	Each objKit In m_cKits
				'check id
				If (nSize = objKit.GetTestsPerKit()) Then
					'found it, so get the index and return
					objFoundKit = objKit
				End If
			Next objKit
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object FindKitBySize. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FindKitBySize = objFoundKit
		
	End Function
	
	
	
	'+
	'FindKitByID()
	'
	'Returns the kit
	'
	'lbailey
	'1 July 2002
	'-
	Public Function FindKitByID(ByRef strID As String) As Kit
		
		Dim objFoundKit As New Kit
		
		Dim objKit As Kit
		If (m_cKits.Count() > 0) Then
			
			'iterate through the collection
			For	Each objKit In m_cKits
				'check id
				If (strID = objKit.GetID()) Then
					'found it, so get the index and return
					objFoundKit = objKit
				End If
			Next objKit
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object FindKitByID. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FindKitByID = objFoundKit
		
	End Function
	
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' public methods
	
	'+
	'Load()
	'
	'find all of the selected brands who are children of the specified
	'quantification, then find all of the quantities who are children of
	'those selected brands, and whose type matches the type specified
	'
	'having these lists, then, we can create a list of brands and of the
	'quantities appropriate to the given context
	'
	'lbailey
	'6 june 2002
	'-
	Public Sub Load(ByRef strBrandID As String)
		
		'set up input parameters for stored procudure
        Dim aParams(1, 1) As Object
		
		'UPGRADE_WARNING: Couldn't resolve default property of object aParams(0, 0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(0, 0) = (strBrandID)
		'UPGRADE_WARNING: Couldn't resolve default property of object aParams(0, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(0, 1) = DbType.String
		'create a collection to hold the db objects that will be returned
		Dim cKitDB As Collection
		
		'go get this brand's kits
		Dim objKitColl As New ProQdb.KitDBCollection
		Dim objKitDB As New ProQdb.KitDB
        cKitDB = objKitColl.GetCollection(SP_GET_KITS_BY_BRAND, aParams)
		
		'for each one of the kits, create a mid-tier obj and
		'put it in our collection
        For Each objKitDB In cKitDB
            Dim objKit As New Kit
            objKit.Load(objKitDB)
            m_cKits.Add(objKit)
            'UPGRADE_NOTE: Object objKit may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            objKit = Nothing
        Next objKitDB
		
	End Sub
	
	
	
	'+
	'RemoveAllKits()
	'
	'empties the list of kits
	'
	'lbailey
	'5 june 2002
	'-
	Public Sub RemoveAll()
		Dim i As Short
		For i = 1 To m_cKits.Count()
			m_cKits.Remove(1)
		Next i
	End Sub
	
	
	
	'+
	'standard accessors
	'
	'lbailey
	'14 june 2002
	'-
	Public Function GetKits() As Collection
		GetKits = m_cKits
	End Function
End Class