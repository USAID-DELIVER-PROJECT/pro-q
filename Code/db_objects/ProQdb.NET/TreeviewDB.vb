Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("TreeViewDB_NET.TreeViewDB")> Public Class TreeViewDB
	Implements System.Collections.IEnumerable
	'local variable to hold collection
	Private mCol As Collection
	Public Function FindNode(ByRef lngTreeviewCount As Object, ByRef RefID As Object, Optional ByRef QuantificationID As Object = Nothing) As String
		
		Dim colTemp As Collection
		Dim nodTemp As TreeNodeCollectionDB
		Dim i As Short
		
		If IsDbNull(QuantificationID) Then
			'Aggregate Function
			colTemp = mCol
			
			With colTemp
				For i = 0 To colTemp.Count()
					nodTemp = colTemp.Item(i)
                    If nodTemp.TreeviewCount = lngTreeviewCount Then
                        FindNode = nodTemp.Key
                        Exit For
                    Else
                    End If
				Next 
			End With
		Else
			'In a Quantification
			colTemp = mCol
			
			With colTemp
				For i = 1 To colTemp.Count()
					nodTemp = colTemp.Item(i)

                    If Not nodTemp.Quantification Is Nothing Then
                        If QuantificationID.Equals(nodTemp.Quantification.ToString()) Then
                            If nodTemp.TreeviewCount = lngTreeviewCount Then
                                FindNode = nodTemp.Key
                                Exit For
                            Else
                            End If
                        End If
                    End If

				Next 
			End With
		End If
		
		
	End Function
	Public Function FindNodeByForm(ByRef strForm As String, ByRef RefID As Object, Optional ByRef QuantificationID As Object = Nothing) As String
		
		Dim colTemp As Collection
		Dim nodTemp As TreeNodeCollectionDB
		Dim i As Short
		
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If IsDbNull(QuantificationID) Then
			'Aggregate Function
			colTemp = mCol
			
			With colTemp
				For i = 0 To colTemp.Count()
					nodTemp = colTemp.Item(i)
					If nodTemp.Form = strForm Then
                        FindNodeByForm = nodTemp.Key
						Exit For
					Else
					End If
				Next 
			End With
		Else
			'In a Quantification
			colTemp = mCol
			
			With colTemp
				For i = 1 To colTemp.Count()
					nodTemp = colTemp.Item(i)
					'UPGRADE_WARNING: Couldn't resolve default property of object QuantificationID. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object nodTemp.Quantification. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'

                    ' -- jleiner:  Compare string to string...
                    If Not nodTemp.Quantification = Nothing Then
                        If nodTemp.Quantification.ToString = QuantificationID Then
                            If nodTemp.Form = strForm Then
                                FindNodeByForm = nodTemp.Key
                                Exit For
                            Else
                            End If
                        End If
                    End If
				Next 
			End With
		End If
		
		
	End Function
	
    Public Function Add(ByRef TreeviewCount As Short, ByRef Key As String, ByRef Action As Object, ByRef Form As String, ByRef Label As String, ByRef Icon As Short, ByRef Quantification As Object, Optional ByRef sKey As String = "") As TreeNodeCollectionDB
        'create a new object
        Dim objNewMember As TreeNodeCollectionDB
        objNewMember = New TreeNodeCollectionDB


        'set the properties passed into the method
        objNewMember.TreeviewCount = TreeviewCount
        objNewMember.Key = Key
        'UPGRADE_WARNING: Couldn't resolve default property of object Action. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        objNewMember.Action = Action
        objNewMember.Form = Form
        objNewMember.Label = Label
        objNewMember.Icon = Icon
        'UPGRADE_WARNING: Couldn't resolve default property of object Quantification. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If IsDBNull(Quantification) = False Then
            objNewMember.Quantification = Quantification
        End If
        If Len(sKey) = 0 Then
            mCol.Add(objNewMember)
        Else
            mCol.Add(objNewMember, sKey)
        End If


        'return the object created
        Add = objNewMember
        'UPGRADE_NOTE: Object objNewMember may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        objNewMember = Nothing


    End Function
	
    Default Public ReadOnly Property Item(ByVal vntIndexKey As String) As TreeNodeCollectionDB
        Get
            'used when referencing an element in the collection
            'vntIndexKey contains either the Index or Key to the collection,
            'this is why it is declared as a Variant
            'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
            Item = mCol.Item(vntIndexKey)
        End Get
    End Property
	Public ReadOnly Property Count() As Integer
		Get
			'used when retrieving the number of elements in the
			'collection. Syntax: Debug.Print x.Count
			Count = mCol.Count()
		End Get
	End Property
	
	
	'UPGRADE_NOTE: NewEnum property was commented out. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="B3FC1610-34F3-43F5-86B7-16C984F0E88E"'
	'Public ReadOnly Property NewEnum() As stdole.IUnknown
		'Get
			'this property allows you to enumerate
			'this collection with the For...Each syntax
			'NewEnum = mCol._NewEnum
		'End Get
	'End Property
	
	Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
		'UPGRADE_TODO: Uncomment and change the following line to return the collection enumerator. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="95F9AAD0-1319-4921-95F0-B9D3C4FF7F1C"'
        GetEnumerator = mCol.GetEnumerator
	End Function
	
	
	Public Sub Remove(ByRef vntIndexKey As Object)
		'used when removing an element from the collection
		'vntIndexKey contains either the Index or Key, which is why
		'it is declared as a Variant
		'Syntax: x.Remove(xyz)
		
		
		mCol.Remove(vntIndexKey)
	End Sub
	
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		'creates the collection when this class is created
		mCol = New Collection
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'destroys collection when this class is terminated
		'UPGRADE_NOTE: Object mCol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mCol = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
End Class