Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("Quantity_NET.Quantity")> Public Class Quantity
	'+
	'Quantity.cls
	'
	'a quantity is a number that represents the number of test of a
	'given brand that will be needed at the given point in the
	'quantification.  the list of points is in the QuantityType enum.
	'
	'this class manages these selected brands and their quantities
	'for the quantification and for the protocol brands
	'
	'lbailey
	'6 june 2002
	'-
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'constants
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'private properties
	
	'type (demand, adj demand, requirements, costs, funds, or total
    Private m_eType As Globals_Renamed.QuantityCategory
	'array of brands information, id, count, and pct
	Private m_cSelectedBrands As Collection
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'public properties
	
	'none
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' private methods
	
	
	'+
	'FindBrand()
	'
	'Returns the index of the brand in the brand collection, or
	'a S_FAILURE if the brand is not present in the collection.
	'
	'lbailey
	'1 june 2002
	'-
	Private Function FindBrand(ByRef objBrand As ProtocolBrand) As Short
		
		'iterate through the collection
		Dim i As Short
		For i = 1 To m_cSelectedBrands.Count()
			'check id
			'UPGRADE_WARNING: Couldn't resolve default property of object m_cSelectedBrands().GetID. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (objBrand.GetID() = m_cSelectedBrands.Item(i).GetID()) Then
				'found it, so get the index and return
				FindBrand = i
				Exit Function
			End If
		Next 
		
		'didn't find it
		FindBrand = S_FAILED
		
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
	Public Sub Load(ByRef strQuantificationID As String, ByRef nCategory As Short)
		
		
		
	End Sub
	
	'+
	'AddProtocolBrand()
	'
	'checks to see whether brand is already present in the collection.
	'if it is, then raises an error.  if it isn't, then adds it.
	'
	'lbailey
	'1 june 2002
	'-
	Public Function AddBrand(ByRef objProtocolBrand As ProtocolBrand) As Short
		
		'is brand already present?
		Dim intIndex As Short
		intIndex = FindBrand(objProtocolBrand)
		If (intIndex <> S_FAILURE) Then
			'yes, it's already here, so can't add
			
			'TODO: raise error here
			
			Exit Function
		End If
		
		m_cSelectedBrands.Add(objProtocolBrand)
		
		'get the index of the new brand and return it
		AddBrand = FindBrand(objProtocolBrand)
		
	End Function
	
	
	
	'+
	'AccumulateBrandCount()
	'
	'if the brand doesn't already exist, this func will add it.
	'returns the new total count for this brand
	'
	'lbailey
	'2 june 2002
	'-
	Public Function AccumulateBrand(ByRef objBrand As ProtocolBrand) As Integer
		
		'see if brand is already here
		Dim nIndex As Short
		nIndex = FindBrand(objBrand)
		
		'if not, then add it
		If (nIndex = S_FAILURE) Then
			nIndex = AddBrand(objBrand)
		End If
		
		'now we have a good index, so accumulate to it
		Dim lStoredCount As Integer
		'UPGRADE_WARNING: Couldn't resolve default property of object m_cSelectedBrands().GetCount. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		lStoredCount = m_cSelectedBrands.Item(nIndex).GetCount()
		Dim lNewCount As Integer
		lNewCount = objBrand.GetCount()
		'look! it's C! (look-see)
		'UPGRADE_WARNING: Couldn't resolve default property of object m_cSelectedBrands().GetCount. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object m_cSelectedBrands().SetCount. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_cSelectedBrands.Item(nIndex).SetCount(m_cSelectedBrands.Item(nIndex).GetCount() + objBrand.GetCount())
		
		'return the new total
		'UPGRADE_WARNING: Couldn't resolve default property of object m_cSelectedBrands.GetCount. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        AccumulateBrand = m_cSelectedBrands.Count()
		
	End Function
	
	
	
	'+
	'GetBrand()
	'
	'returns the specified protocolbrand object
	'
	'lbailey
	'5 june 2002
	'-
	Public Function GetBrand(ByRef strID As String) As ProtocolBrand
		
		Dim objReturn As New ProtocolBrand
		Dim objProtocolBrand As ProtocolBrand
		
		'loop through the collection, finding the one with this id
		For	Each objProtocolBrand In m_cSelectedBrands
			If (strID = objProtocolBrand.GetID()) Then
				'set the return value
				objReturn = objProtocolBrand
				'leave this loop
				Exit For
			End If
		Next objProtocolBrand
		
		GetBrand = objReturn
		
	End Function
	
	
	'+
	'RemoveAllBrands()
	'
	'empties the list of protocolbrands
	'
	'lbailey
	'5 june 2002
	'-
	Public Sub RemoveAllBrands()
		Dim i As Short
		For i = 1 To m_cSelectedBrands.Count()
			m_cSelectedBrands.Remove(1)
		Next i
	End Sub
	
	
	Public Function GetProtocolBrands() As Collection
		GetProtocolBrands = m_cSelectedBrands
	End Function
End Class