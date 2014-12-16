Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("TreeNodeCollectionDB_NET.TreeNodeCollectionDB")> Public Class TreeNodeCollectionDB
	Public Key As String
	
	'local variable(s) to hold property value(s)
	Private mvarAction As Object 'local copy
	Private mvarForm As String 'local copy
	Private mvarLabel As String 'local copy
	Private mvarIcon As Short 'local copy
	Private mvarQuantification As Object 'local copy
	Private mvarTreeviewCount As Object 'local Copy
	
	
	
	
	Public Property TreeviewCount() As Object
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Action
			'UPGRADE_WARNING: IsObject has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			If IsReference(mvarTreeviewCount) Then
				TreeviewCount = mvarTreeviewCount
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarTreeviewCount. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object TreeviewCount. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				TreeviewCount = mvarTreeviewCount
			End If
		End Get
		Set(ByVal Value As Object)
			If IsReference(Value) And Not TypeOf Value Is String Then
				'used when assigning an Object to the property, on the left side of a Set statement.
				'Syntax: Set x.Action = Form1
				mvarTreeviewCount = Value
			Else
				'used when assigning a value to the property, on the left side of an assignment.
				'Syntax: X.Action = 5
				'UPGRADE_WARNING: Couldn't resolve default property of object vData. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarTreeviewCount. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarTreeviewCount = Value
			End If
		End Set
	End Property
    Public Property Quantification() As Object
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.Quantification
            'UPGRADE_WARNING: IsObject has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            If IsReference(mvarQuantification) Then
                Quantification = mvarQuantification
            Else
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarQuantification. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Quantification. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Quantification = mvarQuantification
            End If
        End Get
        Set(ByVal Value As Object)
            If IsReference(Value) Then 'And Not TypeOf Value Is String Then
                'used when assigning an Object to the property, on the left side of a Set statement.
                'Syntax: Set x.Quantification = Form1
                mvarQuantification = Value
            Else
                'used when assigning a value to the property, on the left side of an assignment.
                'Syntax: X.Quantification = 5
                'UPGRADE_WARNING: Couldn't resolve default property of object vData. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarQuantification. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                mvarQuantification = Value
            End If
        End Set
    End Property
	
	
	
	
	
	Public Property Icon() As Short
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Icon
			Icon = mvarIcon
		End Get
		Set(ByVal Value As Short)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Icon = 5
			mvarIcon = Value
		End Set
	End Property
	
	
	
	
	
	Public Property Label() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Label
			Label = mvarLabel
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Label = 5
			mvarLabel = Value
		End Set
	End Property
	
	
	
	
	
	Public Property Form() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Form
			Form = mvarForm
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Form = 5
			mvarForm = Value
		End Set
	End Property
	
	
	
	
	
	
	
	Public Property Action() As Object
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Action
			'UPGRADE_WARNING: IsObject has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			If IsReference(mvarAction) Then
				Action = mvarAction
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarAction. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Action. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Action = mvarAction
			End If
		End Get
		Set(ByVal Value As Object)
			If IsReference(Value) And Not TypeOf Value Is String Then
				'used when assigning an Object to the property, on the left side of a Set statement.
				'Syntax: Set x.Action = Form1
				mvarAction = Value
			Else
				'used when assigning a value to the property, on the left side of an assignment.
				'Syntax: X.Action = 5
				'UPGRADE_WARNING: Couldn't resolve default property of object vData. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarAction. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarAction = Value
			End If
		End Set
	End Property
End Class