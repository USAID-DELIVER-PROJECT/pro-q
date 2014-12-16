Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("SelectedBrands_NET.SelectedBrands")> Public Class SelectedBrands
	'+
	'SelectedBrands.cls
	'
	'this class manages these selected brands
	'for the quantification and for the protocol brands
	'
	'lbailey
	'8 june 2002
	'-
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'private properties
	
	
	'array of brands information, id, count, and pct
	Private m_strQuantificationID As String
	Private m_cSelectedBrands As New Collection
	
	
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
	Public Function FindBrand(ByRef objBrand As Brand, ByRef objKit As Kit, Optional ByRef fGeneric As Boolean = False) As Short
		
		Dim nIndex As Short
		nIndex = S_FAILED
		
		Dim i As Short
		If (m_cSelectedBrands.Count() > 0) Then
			
			'iterate through the collection
			For i = 1 To m_cSelectedBrands.Count()
				'check id
				'UPGRADE_WARNING: Couldn't resolve default property of object m_cSelectedBrands().getGeneric. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object m_cSelectedBrands(i).GetKitID. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object m_cSelectedBrands().GetBrandID. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (objBrand.GetID() = m_cSelectedBrands.Item(i).GetBrandID()) And (objKit.GetID() = m_cSelectedBrands.Item(i).GetKitID) And fGeneric = m_cSelectedBrands.Item(i).getGeneric() Then
					'found it, so get the index and return
					nIndex = i
					Exit For
				End If
			Next 
		End If
		
		FindBrand = nIndex
		
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
	Public Sub Load(ByRef strQuantificationID As String)
		
		'Store the QuantificationID for Later
		m_strQuantificationID = strQuantificationID
		
		'set up input parameters for stored procudure
		Dim aParams(1, 1) As Object
		
		'UPGRADE_WARNING: Couldn't resolve default property of object aParams(0, 0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(0, 0) = New Guid(strQuantificationID)
		'UPGRADE_WARNING: Couldn't resolve default property of object aParams(0, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(0, 1) = DbType.Guid
		
		'create a collection to hold the db objects that will be returned
		Dim cSelBrandDBs As Collection
		
		'go get this quantification's brands
		Dim objSelBrandColl As New ProQdb.SelectedBrandDBCollection
		Dim objSelBrandDB As New ProQdb.SelectedBrandDB
		cSelBrandDBs = objSelBrandColl.GetCollection(SP_GET_SELECTEDBRAND_BY_QUANT, aParams)
		
		'for each one of the selbrands, create a mid-tier obj and
		'put it in our collection of selected brands
        For Each objSelBrandDB In cSelBrandDBs
            Dim objSelBrand As New SelectedBrand
            objSelBrand.LoadFromObject(objSelBrandDB)
            m_cSelectedBrands.Add(objSelBrand)
            'UPGRADE_NOTE: Object objSelBrand may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            objSelBrand = Nothing
        Next objSelBrandDB
		
	End Sub
	
	
	
	'+
	'Save()
	'
	'pushes the data out to the db
	'
	'lbailey
	'27 june 2002
	'-
	Public Sub Save()
		Dim objSelectedBrand As SelectedBrand
		For	Each objSelectedBrand In m_cSelectedBrands
			objSelectedBrand.Save()
		Next objSelectedBrand
	End Sub
	
	
	'+
	'AddBrand()
	'
	'checks to see whether brand is already present in the collection.
	'if it isn't, it adds it.
	'if it is, then it accumlates the qtys.
	'
	'returns the index of the inserted brand
	'
	'lbailey
	'1 june 2002
	'-
	Public Function AddBrand(ByRef objSelectedBrand As SelectedBrand) As Short
		
		'is brand already present?
        Dim intIndex As Short
        'Dim objBrand As New ProtocolBrand

        intIndex = FindBrand(objSelectedBrand.GetBrand(), objSelectedBrand.GetKit(), objSelectedBrand.getGeneric())
        'objBrand.Load(objSelectedBrand.GetBrandID())

		If (intIndex <> S_FAILURE) Then
			'we found it, so let's accumulate
            'UPGRADE_WARNING: Couldn't resolve default property of object objSelectedBrand. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            AccumulateBrand(objSelectedBrand)
		Else
			'we can't find it, so add it
            m_cSelectedBrands.Add(objSelectedBrand)
		End If
		
		'get the index of the new brand and return it
		AddBrand = FindBrand(objSelectedBrand.GetBrand(), objSelectedBrand.GetKit(), objSelectedBrand.getGeneric())
		
	End Function
	
	
	
	'+
	'AccumulateBrand()
	'
	'if the brand doesn't already exist, this func will add it.
	'returns the new total count for this brand
	'
	'lbailey
	'2 june 2002
	'-
    Public Function AccumulateBrand(ByRef objProtocolBrand As Object) As Integer

        'see if brand is already here
        Dim nIndex As Short
        nIndex = FindBrand(objProtocolBrand.GetBrand(), objProtocolBrand.GetKit, objProtocolBrand.getGeneric)

        'if not, then create and add it
        Dim objSelectedBrand As New SelectedBrand
        If (nIndex = S_FAILURE) Then

            objSelectedBrand.Create()
            objSelectedBrand.SetBrand(objProtocolBrand.GetBrand())
            objSelectedBrand.SetCount(objProtocolBrand.GetCount())
            objSelectedBrand.SetKit(objProtocolBrand.GetKit)
            objSelectedBrand.SetKitCost(objProtocolBrand.GetKit.GetCost)
            objSelectedBrand.SetQuantificationID(m_strQuantificationID)
            objSelectedBrand.SetGeneric(objProtocolBrand.getGeneric)
            objSelectedBrand.SetGenericCode(objProtocolBrand.getGenericCode)
            'objSelectedBrand.SetQuantificationID =
            nIndex = AddBrand(objSelectedBrand)
        Else
            'now add the counts
            'UPGRADE_WARNING: Couldn't resolve default property of object m_cSelectedBrands().Add. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            m_cSelectedBrands.Item(nIndex).Add(objProtocolBrand.GetCount())
        End If

        'return the new total
        'UPGRADE_WARNING: Couldn't resolve default property of object m_cSelectedBrands().GetCount. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        AccumulateBrand = m_cSelectedBrands.Item(nIndex).GetCount()

    End Function
	
	
	
	'+
	'GetBrand()
	'
	'returns the specified brand object
	'
	'lbailey
	'5 june 2002
	'-
	Public Function GetBrand(ByRef strID As String) As SelectedBrand
		
		Dim objReturn As New SelectedBrand
		Dim objSelectedBrand As SelectedBrand
		
		'loop through the collection, finding the one with this id
		For	Each objSelectedBrand In m_cSelectedBrands
			If (strID = objSelectedBrand.GetID()) Then
				'set the return value
				objReturn = objSelectedBrand
				'leave this loop
				Exit For
			End If
		Next objSelectedBrand
		
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
	
	
	
	'+
	'GetTypeTotal()
	'
	'looks for all of the brands that do or do not require lab work
	'(as indicated by the inval flag) and totals up the number of
	'samples that will be tested within or without the lab, as
	'appropriate
	'
	'lbailey
	'26 june 2002
	'-
	Public Function GetTypeTotal(ByRef fIsLab As Boolean) As Integer
		
		'create and init retval
		Dim lTypeTotal As Integer
		lTypeTotal = 0
		
		'rip through selbrands, crunching the total for this type
		Dim objSelBrand As SelectedBrand
		For	Each objSelBrand In m_cSelectedBrands
			If (objSelBrand.GetBrand().IsLab() = fIsLab) Then
				lTypeTotal = lTypeTotal + objSelBrand.GetCount()
			End If
		Next objSelBrand
		
		GetTypeTotal = lTypeTotal
		
	End Function
	
	
	
	'+
	'GetTypeBrands()
	'
	'looks for all of the brands that do or do not require lab work
	'(as indicated by the inval flag) and cooks up a collection of them,
	'putting the percentage of their type-total in each object as we
	'go.
	'
	'lbailey
	'26 june 2002
	'-
	Public Function GetTypeBrands(ByRef fIsLab As Boolean) As Collection
		
		'get the number of tests of this type
		Dim lTypeTotal As Integer
		lTypeTotal = GetTypeTotal(fIsLab)
		
		'go find each test of this type, and its percentage, and put em
		'in a collection
		Dim sngPercent As Single
		Dim cTypeBrands As New Collection
		Dim objSelBrand As SelectedBrand
		For	Each objSelBrand In m_cSelectedBrands
			If (objSelBrand.GetBrand().IsLab() = fIsLab) Then
				'calculate the percentage
				If (lTypeTotal = 0) Then
					sngPercent = 0
				Else
					sngPercent = objSelBrand.GetCount() / lTypeTotal
				End If
				objSelBrand.SetPercent(sngPercent)
				
				'add this brand to the collection
				cTypeBrands.Add(objSelBrand, objSelBrand.GetID())
			End If
		Next objSelBrand
		
		GetTypeBrands = cTypeBrands
		
	End Function
	
	
	
	'+
	'standard accessors
	'
	'lbailey
	'14 june 2002
	'-
	Public Function GetSelectedBrands() As Collection
		GetSelectedBrands = m_cSelectedBrands
	End Function
	
	Public Function GetQuantificationID() As String
		GetQuantificationID = m_strQuantificationID
	End Function
	
	Public Sub SetQuantificationID(ByRef strID As String)
		m_strQuantificationID = strID
	End Sub
End Class