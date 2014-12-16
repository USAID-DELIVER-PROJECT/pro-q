Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("SelectedBrand_NET.SelectedBrand")> Public Class SelectedBrand
	'+
	'SelectedBrand.cls
	'
	'this class represents a selected brands and its quantities
	'
	'8 june 2002
	'-
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'private properties
	
	'dbobj holding brands information
	Private m_objSelectedBrandDB As New ProQdb.SelectedBrandDB
	'Object for the Brand information
	Private m_objBrand As New Brand
	Private m_objKit As New Kit
	'the number of samples to be tested with this brand
	Private m_lCount As Integer
	'the percentage of the samples that will be tested with
	'this brand
	Private m_sngPercent As Single
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' public methods
	
	'+
	'Create()
	'
	'creates a new selected brand object, and a new quantity object for
	'it
	'
	'lbailey
	'8 june 2002
	'-
	Public Function Create() As String
		
		'create the selected brand object
		Dim strBrandID As String
		
		strBrandID = m_objSelectedBrandDB.Create()
		
		Create = strBrandID
		
	End Function
	
	
	'+
	'Load()
	'
	'go get the selected brand indicated by the ID passed in, and then
	'load the associated brand objectgo
	'
	'lbailey
	'8 june 2002
	'-
	Public Sub Load(ByRef strID As String)
		
		'load the selected brand
		m_objSelectedBrandDB.Load(strID)
		
		'Load the BrandObject
		m_objBrand.Load(m_objSelectedBrandDB.GetBrandID)
		
		'Load the Kit Object
		m_objKit.Load(m_objSelectedBrandDB.GetKitID)
		
		
	End Sub
	
	
	'+
	'Load()
	'
	'get a reference to the selected brand db obj passed in, and then
	'load the associated brand objectgo
	'
	'lbailey
	'8 june 2002
	'-
	Public Sub LoadFromObject(ByRef objSelectedBrandDB As ProQdb.SelectedBrandDB)
		m_objSelectedBrandDB = objSelectedBrandDB
		
		'Load the BrandObject
		m_objBrand.Load(m_objSelectedBrandDB.GetBrandID)
		
		'Load the Kit Object
		m_objKit.Load(m_objSelectedBrandDB.GetKitID)
		
	End Sub
	
	
	'+
	'Delete()
	'
	'deletes the data from the backend
	'
	'lbailey
	'8 june 2002
	'-
	Public Sub Delete()
		m_objSelectedBrandDB.Delete()
	End Sub
	
	
	
	'+
	'Update()
	'
	'pushes the data to the backend
	'
	'lbailey
	'8 june 2002
	'-
	Public Sub Update()
		m_objSelectedBrandDB.Update()
	End Sub
	
	Public Sub Save()
		'Call GetSBName to Make sure the generic name is set.
		GetSBName()
		Update()
	End Sub
	
	
	'+
	'Add()
	'
	'increments the count by the specified amount
	'
	'lbailey
	'26 june 2002
	'-
	Public Sub Add(ByRef lCount As Integer)
		lCount = lCount + GetCount
		SetCount(lCount)
	End Sub
	
	
	
	'+
	'Subtract()
	'
	'decrements the amount by the specified amount, without going
	'below zero
	'
	'lbailey
	'26 june 2002
	'-
	Public Sub Subtract(ByRef lCount As Integer)
		
		'check for negative sums
		If (lCount > GetCount()) Then
			Err.Description = ERROR_SB_CANT_SUBTRACT
			Exit Sub
		End If
		'do the math and write it out
		lCount = GetCount - lCount
		SetCount(lCount)
		
	End Sub
	
	Public Function GetTestsPerKit() As Short
		GetTestsPerKit = m_objKit.GetTestsPerKit()
	End Function
	
	
	
	'+
	'Teminate()
	'
	'releases references to objects
	'
	'lbailey
	'26 june 2002
	'-
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object m_objSelectedBrandDB may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_objSelectedBrandDB = Nothing
		'UPGRADE_NOTE: Object m_objBrand may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_objBrand = Nothing
		'UPGRADE_NOTE: Object m_objKit may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_objKit = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	
	
	'+
	'standard manipulator functions (this time we have 2 backends
	'
	'lbailey
	'16 may 2002
	'-
	
	'selected brands...............
	Public Sub SetID(ByRef strID As String)
		'Use this one carefully.
		m_objSelectedBrandDB.SetID(strID)
	End Sub
	
	Public Sub SetBrandID(ByRef strBrandID As String)
		m_objSelectedBrandDB.SetBrandID(strBrandID)
		m_objBrand.Load(strBrandID)
	End Sub
	
	Public Sub SetKitID(ByRef strKitID As String)
		m_objSelectedBrandDB.SetKitID(strKitID)
		m_objKit.Load(strKitID)
	End Sub
	Public Sub SetQuantificationID(ByRef strID As String)
		m_objSelectedBrandDB.SetQuantificationID(strID)
	End Sub
	
	
	
	'Public Sub SetTestsPerKit(ByRef intTestsPerKit As Integer)
	'm_objSelectedBrandDB.SetTestsPerKit intTestsPerKit
	'End Sub
	
	Public Sub SetKitCost(ByRef dblKitCost As Double)
		m_objSelectedBrandDB.SetKitCost(CDbl(dblKitCost))
	End Sub
	
	Public Sub SetCount(ByRef lCount As Integer)
		m_objSelectedBrandDB.SetCount(lCount)
	End Sub
	
	Public Sub SetBrand(ByRef objBrand As Brand)
		m_objBrand = objBrand
		SetBrandID(objBrand.GetID)
	End Sub
	Public Sub SetKit(ByRef objKit As Kit)
		m_objKit = objKit
		SetKitID(objKit.GetID)
	End Sub
	
	Public Sub SetPercent(ByRef sngPercent As Single)
		m_sngPercent = sngPercent
	End Sub
	
	Public Sub SetGeneric(ByRef fIsGeneric As Boolean)
		m_objSelectedBrandDB.SetGeneric(fIsGeneric)
	End Sub
	Public Sub SetGenericCode(ByRef strGenericCode As String)
		m_objSelectedBrandDB.SetGenericCode(strGenericCode)
	End Sub
	
	'+
	'standard accessor functions
	'
	'lbailey
	'16 may 2002
	'-
	
	'selected brands...............
	
	Public Function GetID() As String
		GetID = m_objSelectedBrandDB.GetID()
	End Function
	
	Public Function GetBrandID() As String
		GetBrandID = m_objBrand.GetID()
	End Function
	
	Public Function GetKitID() As String
		GetKitID = m_objSelectedBrandDB.GetKitID()
	End Function
	
	Public Function GetQuantificationID() As String
		GetQuantificationID = m_objSelectedBrandDB.GetQuantificationID
	End Function
	
	Public Function GetKitCost() As Double
		GetKitCost = m_objSelectedBrandDB.GetKitCost()
	End Function
	
	Public Function GetCount() As Integer
		GetCount = m_objSelectedBrandDB.GetCount()
	End Function
	
	Public Function GetBrand() As Brand
		GetBrand = m_objBrand
	End Function
	
	Public Function GetKit() As Kit
		GetKit = m_objKit
	End Function
	
	Public Function GetPercent() As Single
		GetPercent = m_sngPercent
	End Function
	
	Public Function GetSBName() As String
		
		Dim strType As String
		
		
		If getGeneric() = False Then
			GetSBName = m_objBrand.GetName & " - " & m_objKit.GetTestsPerKit()
		Else
			Select Case m_objBrand.GetType_Renamed()
				Case 1
					strType = "Elisa"
				Case 2
					strType = "Rapid"
				Case 3
					strType = "Blot"
				Case Else
					strType = ""
			End Select
			
			GetSBName = "Generic " & strType & " " & getGenericCode() & " - " & m_objKit.GetTestsPerKit()
		End If
		
	End Function
	Public Function getGeneric() As Boolean
		getGeneric = m_objSelectedBrandDB.getGeneric
	End Function
	Public Function getGenericCode() As String
		
		Dim strCode As String
		Dim strQuantificationID As String
		Dim objQuantDB As New ProQdb.QuantificationDB
        Dim rst As New DataSet 'ADODB.Recordset
		Dim aParams(1, 1) As Object
        Dim fFound As Boolean
        Dim i As Integer

		fFound = False
		strCode = m_objSelectedBrandDB.getGenericCode
		
		'GoTo Proc_exit
		
		If strCode <> "" Then
			getGenericCode = strCode
		Else
			' Have to generate the code based on the selected Quantification/Aggregation.
			strQuantificationID = m_objSelectedBrandDB.GetQuantificationID
			
			'UPGRADE_WARNING: Couldn't resolve default property of object aParams(0, 0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            aParams(0, 0) = New Guid(strQuantificationID)
			'UPGRADE_WARNING: Couldn't resolve default property of object aParams(0, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            aParams(0, 1) = DbType.Guid
			
            'rst.CursorType = ADODB.CursorTypeEnum.adOpenStatic
			rst = objQuantDB.ReturnRS("qselProtocolBrandsbyAggregationQuant", aParams)
			
			With rst
                'If (.BOF And .EOF) Then
                'strCode = "ZZ"
                'Else
                'If Not (.BOF And .EOF) Then .MoveFirst
                For i = 0 To .Tables("qselProtocolBrandsbyAggregationQuant").Rows.Count - 1 'Do Until fFound Or .EOF
                    If .Tables("qselProtocolBrandsbyAggregationQuant").Rows(i).item("brandID") = GetBrandID() Then
                        fFound = True
                        Exit For
                        'Else
                        '.MoveNext()
                    End If
                next 'Loop
                If fFound Then 'Not .EOF Then
                    'Record found, use its Generic Code
                    If .Tables("qselProtocolBrandsbyAggregationQuant").Rows(i).Item("strGenericCode") = "" Then
                        strCode = FindNextCode(rst)
                    Else
                        strCode = .Tables("qselProtocolBrandsbyAggregationQuant").Rows(i).Item("strGenericCode")
                    End If
                Else
                    strCode = FindNextCode(rst)
                End If
                'End If
                m_objSelectedBrandDB.SetGenericCode(strCode)
                Save()
            End With
            'rst.Close()
        End If

Proc_Exit:
        getGenericCode = strCode
		
	End Function
	
    Function FindNextCode(ByRef rst As DataSet) As String

        Dim i As Short
        Dim x As Short
        Dim j As Integer
        Dim strtempCode As String
        Dim fFound As Boolean
        'Dim rstLocal As New adodb.Recordset

        'rstLocal.CursorType = adOpenDynamic
        'Set rstLocal = rst
        fFound = False
        With rst

            For i = 90 To 65 Step -1
                strtempCode = Chr(i) & Chr(i)
                fFound = False

                'If Not (.BOF And .EOF) Then .MoveFirst()
                For j = 0 To .Tables("qselProtocolBrandsbyAggregationQuant").Rows.Count - 1 'Do Until fFound Or .EOF
                    If .Tables("qselProtocolBrandsbyAggregationQuant").Rows(j).Item("strGenericCode") = strtempCode Then
                        fFound = True
                        Exit For
                        'Else
                        '.MoveNext()
                    End If
                Next 'Loop

                If fFound = False Then
                    'Code Not Found, return the code
                    FindNextCode = strtempCode
                    GoTo Proc_Exit
                End If
            Next i

            'All Codes Found in ZZ-AA range so keep looking
            For i = 90 To 65 Step -1
                For x = 90 To 65 Step -1
                    strtempCode = Chr(i) & Chr(x)

                    fFound = False

                    'If Not (.BOF And .EOF) Then .MoveFirst()

                    For j = 0 To .Tables("qselProtocolBrandsbyAggregationQuant").Rows.Count - 1 'Do Until fFound Or .EOF
                        If .Tables("qselProtocolBrandsbyAggregationQuant").Rows(j).Item("strGenericCode") = strtempCode Then
                            fFound = True
                            Exit For
                            'Else
                            '.MoveNext()
                        End If
                    Next 'Loop

                    If fFound = False Then
                        'Code Not Found, return the code
                        FindNextCode = strtempCode
                        GoTo Proc_Exit
                    End If
                Next x
            Next i
        End With

Proc_Exit:
        Exit Function

    End Function
End Class