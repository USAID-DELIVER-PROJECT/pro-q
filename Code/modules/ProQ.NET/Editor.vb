Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("Editor_NET.Editor")> Public Class Editor
	
	'editor.cls
	'
	'the editor uses the edit rules to validate values.
	'
	'lbailey
	'25 may 2002
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'constants
	
	'result codes
	Private Const S_OK As Short = 0
	Private Const S_FAILED As Short = -1
	
	'data type names
	Private Const DATATYPE_STRING As String = "STRING"
	Private Const DATATYPE_DATE As String = "DATE"
	Private Const DATATYPE_CURRENCY As String = "CURRENCY"
	Private Const DATATYPE_BOOLEAN As String = "BOOLEAN"
	Private Const DATATYPE_BYTE As String = "BYTE"
	Private Const DATATYPE_INTEGER As String = "INTEGER"
	Private Const DATATYPE_LONG As String = "LONG"
	Private Const DATATYPE_DOUBLE As String = "DOUBLE"
	Private Const DATATYPE_SINGLE As String = "SINGLE"
	Private Const DATATYPE_VARIANT As String = "VARIANT"
	'Private Const DATATYPE_SHORT = "SINGLE"
	
	'error strings
	Private Const ERROR_NONE As String = "There is no error."
	Private Const ERROR_INVALID_EDIT_RULE As String = "Invalid Edit Rule was used."
	
	Private Const ERROR_STRING_LENGTH As String = "String is too long."
	Private Const ERROR_STRING_EMPTY As String = "Please enter a value to continue."
	
	'Currency
	Private Const ERROR_CURRENCY As String = "Please enter a number with two decimal places to continue."
	Private Const ERROR_CURRENCY_NEGATIVE As String = "Please enter a non-negative number with two decimal places to continue." & vbCrLf & vbCrLf & vbCrLf & "**A non-negative number is any positive number including zero."
	Private Const ERROR_CURRENCY_NONPOSITIVE As String = "Please enter a positive number with two decimal places to continue." & vbCrLf & vbCrLf & vbCrLf & "**A positive number is any positive number excluding zero."
	'Numbers
	Private Const ERROR_NUMBER As String = "Please enter a number to continue."
	Private Const ERROR_NUMBERTOBIG As String = "Please enter a number less than 1,000,000,000,000 to continue."
	Private Const ERROR_NUMBER_NONPOSITIVE As String = "Please enter a positive number to continue." & vbCrLf & vbCrLf & vbCrLf & "**A positive number is any positive number excluding zero."
	Private Const ERROR_NUMBER_NONPOSITIVE_DAYS_PER_YEAR As String = "Please enter a positive number to continue." & vbCrLf & "NOTE: There is a maximum of 366 days in a year." & vbCrLf & vbCrLf & vbCrLf & "**A positive number is any positive number excluding zero."
	Private Const ERROR_NUMBER_NONNEGATIVE As String = "Please enter a negative number to continue." & vbCrLf & vbCrLf & vbCrLf & "**A negative number is any negative number excluding zero."
	Private Const ERROR_NUMBER_POSITIVE As String = "Please enter a non-positive number to continue." & vbCrLf & vbCrLf & vbCrLf & "**A non-positive number is any negative number including zero."
	Private Const ERROR_NUMBER_NEGATIVE As String = "Please enter a non-negative number to continue." & vbCrLf & vbCrLf & vbCrLf & "**A non-negative number is any positive number including zero."
	Private Const ERROR_NUMBER_ZERO As String = "Please enter a non-zero number to continue."
	
	'Integers
	Private Const ERROR_INTEGER As String = "Please enter an integer to continue."
	Private Const ERROR_INTEGER_NONPOSITIVE As String = "Please enter a positive integer to continue." & vbCrLf & vbCrLf & vbCrLf & "**A positive integer is any positive integer excluding zero."
	Private Const ERROR_INTEGER_NONNEGATIVE As String = "Please enter a negative integer to continue." & vbCrLf & vbCrLf & vbCrLf & "**A negative integer is any negative integer excluding zero."
	Private Const ERROR_INTEGER_POSITIVE As String = "Please enter a non-positive integer to continue." & vbCrLf & vbCrLf & vbCrLf & "***A non-positive integer is any negative integer including zero."
	Private Const ERROR_INTEGER_NEGATIVE As String = "Please enter a non-negative integer to continue." & vbCrLf & vbCrLf & vbCrLf & "**A non-negative integer is any positive integer including zero."
	Private Const ERROR_INTEGER_ZERO As String = "Please enter a non-zero integer to continue."
	
	'Percents
	Private Const ERROR_PERCENT As String = "Please enter a percentage to continue. (i.e. enter 12.45 for 12.45%)"
	Private Const ERROR_PERCENT_NONPOSITIVE As String = "Please enter a positive percentage to continue. (i.e. enter 12.45 for 12.45%)" & vbCrLf & vbCrLf & vbCrLf & "**A positive percentage is any positive percentage excluding zero."
	Private Const ERROR_PERCENT_NONNEGATIVE As String = "Please enter a negative percentage to continue. (i.e. enter -12.45 for -12.45%)" & vbCrLf & vbCrLf & vbCrLf & "**A negative percentage is any negative percentage excluding zero."
	Private Const ERROR_PERCENT_POSITIVE As String = "Please enter a non-positive percentage to continue. (i.e. enter -12.45 for -12.45%)" & vbCrLf & vbCrLf & vbCrLf & "**A non-positive percentage is any negative percentage including zero."
	Private Const ERROR_PERCENT_NEGATIVE As String = "Please enter a non-negative percentage to continue. (i.e. enter 12.45 for 12.45%)" & vbCrLf & vbCrLf & vbCrLf & "**A non-negative percentage is any positive percentage including zero."
	Private Const ERROR_PERCENT_ZERO As String = "Please enter a non-zero percentage to continue. (i.e. enter 12.45 for 12.45%)"
	Private Const ERROR_PERCENT_OVER100 As String = "Please enter a percentage less than or equal to 100% to continue. (i.e. enter 12.45 for 12.45%)"
	Private Const ERROR_PERCENT_POSITIVE_OVER100 As String = "Please enter a positive percentage less than or equal to 100% to continue. (i.e. enter 12.45 for 12.45%)"
	
	'Dates
	Private Const ERROR_DATE As String = "Please enter a date to continue."
	Private Const ERROR_DATE_PAST As String = "Please enter a future date to continue."
	Private Const ERROR_DATE_FUTURE As String = "Please enter a past date to continue."
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'private properties
	
	Private m_strError As String
	Private m_lResult As Integer
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'public properties
	
	'none
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' private methods
	
	
	'EditRule1()
	'
	'is the inval a number?
	'
	'lbailey
	'25 may 2002
	'
	Function EditRule1(ByRef xVal As Object) As Integer
		
		'see if the variant type is not any of the numeric types
		Dim strVal As String
		'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		strVal = UCase(TypeName(xVal))
		If (strVal <> DATATYPE_INTEGER) And (strVal <> DATATYPE_LONG) And (strVal <> DATATYPE_DOUBLE) Then
			If Not IsNumeric(xVal) Then
				'variant is no numeric type, so set error data
				m_lResult = S_FAILED
				m_strError = ERROR_NUMBER
			End If
		End If
		
		'return the result code
		EditRule1 = m_lResult
		
	End Function
	
	
	'EditRule2()
	'
	'is the inval a date?
	'
	'lbailey
	'25 may 2002
	'
	Function EditRule2(ByRef xVal As Object) As Integer
		
		'see if the variant type is a date
		'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If (TypeName(xVal) <> DATATYPE_DATE) Then
            If Not IsDate(xVal) Then
                'variant is no numeric type, so set error data
                m_lResult = S_FAILED
                m_strError = ERROR_DATE
            End If

        End If

        'return the result code
        EditRule2 = m_lResult
		
	End Function
	
	
	
	'EditRule3()
	'
	'is the inval a non-zero number?
	'
	'lbailey
	'25 may 2002
	'
	Function EditRule3(ByRef xVal As Object) As Integer
		
		'first let's see if it's a number
		If (EditRule1(xVal) = S_FAILED) Then
			'if it isn't even a number, then we can quit now
			m_lResult = S_FAILED
			m_strError = ERROR_NUMBER_ZERO
		Else
			'it is a number, so let's check for zero-ness
			'UPGRADE_WARNING: Couldn't resolve default property of object xVal. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (CDbl(xVal) = 0) Then
				'uh-oh, it's zero...
				m_lResult = S_FAILED
				m_strError = ERROR_NUMBER_ZERO
			End If
		End If
		
		'return the result code
		EditRule3 = m_lResult
		
	End Function
	
	'EditRule4()
	'
	'is the inval a future date?
	'
	'lbailey
	'25 may 2002
	'
	Function EditRule4(ByRef xVal As Object) As Integer
		
		'first let's see if it's a date
		If (EditRule2(xVal) = S_FAILED) Then
			'if it isn't even a date, then we can quit now
			m_lResult = S_FAILED
		Else
			'it is a date, so let's check for future-ness
			'UPGRADE_WARNING: Couldn't resolve default property of object xVal. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If (CDate(xVal) <= Now) Then
                'uh-oh, it's not in the future...
                m_lResult = S_FAILED
                m_strError = ERROR_DATE_PAST
            End If
		End If
		
		'return the result code
		EditRule4 = m_lResult
		
	End Function
	
	
	
	'EditRule5()
	'
	'is the inval a past date?
	'
	'lbailey
	'25 may 2002
	'
	Function EditRule5(ByRef xVal As Object) As Integer
		
		'first let's see if it's a date
		If (EditRule2(xVal) = S_FAILED) Then
			'if it isn't even a date, then we can quit now
			m_lResult = S_FAILED
		Else
			'it is a date, so let's check for past-ness
			'UPGRADE_WARNING: Couldn't resolve default property of object xVal. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If (CDate(xVal) >= Now) Then
                'uh-oh, it's not in the past...
                m_lResult = S_FAILED
                m_strError = ERROR_DATE_FUTURE
            End If
		End If
		
		'return the result code
		EditRule5 = m_lResult
		
	End Function
	
	
	
	'EditRule6()
	'
	'is the inval currency?
	'
	'lbailey
	'25 may 2002
	'
	Function EditRule6(ByRef xVal As Object) As Integer
		
		'see if the variant type is currency
		'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If (TypeName(xVal) <> DATATYPE_CURRENCY) Then
            If EditRule1(xVal) = S_FAILED Then
                'variant is not a number, so set error data
                m_lResult = S_FAILED
                m_strError = ERROR_CURRENCY
                'ElseIf CDbl(xVal) >= 1000000000000# Then
                '    m_lResult = S_FAILED
                '    m_strError = ERROR_NUMBERTOBIG
                'UPGRADE_WARNING: Couldn't resolve default property of object xVal. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ElseIf System.Math.Round(CDbl(xVal), 2) <> System.Math.Round(CDbl(xVal), 4) Then
                'variant is not a number, so set error data
                m_lResult = S_FAILED
                m_strError = ERROR_CURRENCY
            End If
        End If
		
		'return the result code
		EditRule6 = m_lResult
		
	End Function
	'EditRule7()
	'
	'is the inval an integer?
	'
	'lbailey
	'25 may 2002
	'
	Function EditRule7(ByRef xVal As Object) As Integer
		'see if the variant type is not any of the numeric types
		Dim strVal As String
		'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		strVal = UCase(TypeName(xVal))
		If (strVal <> DATATYPE_INTEGER) And (strVal <> DATATYPE_LONG) Then
			If (EditRule1(xVal) = S_FAILED) Then
				'variant is not a number, so set error data
				m_lResult = S_FAILED
				m_strError = ERROR_INTEGER
				'UPGRADE_WARNING: Couldn't resolve default property of object xVal. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf CDbl(xVal) >= 1000000000000# Then 
				m_lResult = S_FAILED
				m_strError = ERROR_NUMBERTOBIG
				'UPGRADE_WARNING: Couldn't resolve default property of object xVal. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ElseIf Round_TSB(CDbl(xVal), 0) <> Round_TSB(CDbl(xVal), 2) Then
                'variant is not an integer, so set error data
                m_lResult = S_FAILED
                m_strError = ERROR_INTEGER
			End If
		End If
		
		'return the result code
		EditRule7 = m_lResult
		
	End Function
	'EditRule8()
	'
	'is the inval a percentage >100?
	'
	'lblankenship
	'25 sept 2002
	'
	Function EditRule8(ByRef xVal As Object) As Integer
		
		'see if the variant type is not any of the numeric types
		If EditRule18(xVal) = S_FAILED Then
			'variant is not a number, so set error data
			m_lResult = S_FAILED
			If m_strError <> ERROR_NUMBERTOBIG Then
				m_strError = ERROR_PERCENT_OVER100
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object xVal. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf CDbl(xVal) > 100 Then 
			'variant is not a valid percent, so set error data
			m_lResult = S_FAILED
			m_strError = ERROR_PERCENT_OVER100
		End If
		
		'return the result code
		EditRule8 = m_lResult
		
	End Function
	'EditRule9()
	'
	'is the inval a positive number? (does not include 0)
	'
	'lblankenship
	'25 sept 2002
	'
	Function EditRule9(ByRef xVal As Object) As Integer
		
		'see if the variant type is not any of the numeric types
		If EditRule1(xVal) = S_FAILED Then
			'variant is not a number, so set error data
			m_lResult = S_FAILED
			m_strError = ERROR_NUMBER_NONPOSITIVE
			'UPGRADE_WARNING: Couldn't resolve default property of object xVal. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf CDbl(xVal) <= 0 Then  'check if negative number including 0
			'variant is not a positive number, so set error data
			m_lResult = S_FAILED
			m_strError = ERROR_NUMBER_NONPOSITIVE
		End If
		
		'return the result code
		EditRule9 = m_lResult
		
	End Function
	'EditRule10()
	'
	'is the inval a negative number? (does not include 0)
	'
	'lblankenship
	'25 sept 2002
	'
	Function EditRule10(ByRef xVal As Object) As Integer
		
		'see if the variant type is not any of the numeric types
		If EditRule1(xVal) = S_FAILED Then
			'variant is not a number, so set error data
			m_lResult = S_FAILED
			m_strError = ERROR_NUMBER_NONNEGATIVE
			'UPGRADE_WARNING: Couldn't resolve default property of object xVal. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf CDbl(xVal) >= 0 Then  'check if positive number including 0
			'variant is not a negative number, so set error data
			m_lResult = S_FAILED
			m_strError = ERROR_NUMBER_NONNEGATIVE
		End If
		
		'return the result code
		EditRule10 = m_lResult
		
	End Function
	'EditRule11()
	'
	'is the inval a non-positive number? (includes 0)
	'
	'lblankenship
	'25 sept 2002
	'
	Function EditRule11(ByRef xVal As Object) As Integer
		
		'see if the variant type is not any of the numeric types
		If EditRule1(xVal) = S_FAILED Then
			'variant is not a number, so set error data
			m_lResult = S_FAILED
			m_strError = ERROR_NUMBER_POSITIVE
			'UPGRADE_WARNING: Couldn't resolve default property of object xVal. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf CDbl(xVal) > 0 Then  'check if positive number excluding 0
			'variant is not a nonpositive number, so set error data
			m_lResult = S_FAILED
			m_strError = ERROR_NUMBER_POSITIVE
		End If
		
		'return the result code
		EditRule11 = m_lResult
		
	End Function
	'EditRule12()
	'
	'is the inval a non-negative number? (includes 0)
	'
	'lblankenship
	'25 sept 2002
	'
	Function EditRule12(ByRef xVal As Object) As Integer
		
		'see if the variant type is not any of the numeric types
		If EditRule1(xVal) = S_FAILED Then
			'variant is not a number, so set error data
			m_lResult = S_FAILED
			m_strError = ERROR_NUMBER_NEGATIVE
			'UPGRADE_WARNING: Couldn't resolve default property of object xVal. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf CDbl(xVal) < 0 Then  'check if NEGATIVE number excluding 0
			'variant is not a nonnegative number, so set error data
			m_lResult = S_FAILED
			m_strError = ERROR_NUMBER_NEGATIVE
		End If
		
		'return the result code
		EditRule12 = m_lResult
		
	End Function
	'EditRule13()
	'
	'is the inval a positive integer? (does not include 0)
	'
	'lblankenship
	'25 sept 2002
	'
	Function EditRule13(ByRef xVal As Object) As Integer
		
		'see if the variant type is not an integer
		If EditRule7(xVal) = S_FAILED Then
			'variant is not an integer, so set error data
			m_lResult = S_FAILED
			If m_strError <> ERROR_NUMBERTOBIG Then
				m_strError = ERROR_INTEGER_NONPOSITIVE
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object xVal. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ElseIf Round_TSB(CDbl(xVal), 0) <= 0 Then  'check if negative integer including 0
            'variant is not a positive integer, so set error data
            m_lResult = S_FAILED
            m_strError = ERROR_INTEGER_NONPOSITIVE
		End If
		
		'return the result code
		EditRule13 = m_lResult
		
	End Function
	'EditRule14()
	'
	'is the inval a negative integer? (does not include 0)
	'
	'lblankenship
	'25 sept 2002
	'
	Function EditRule14(ByRef xVal As Object) As Integer
		
		'see if the variant type is not an integer
		If EditRule7(xVal) = S_FAILED Then
			'variant is not an integer, so set error data
			m_lResult = S_FAILED
			If m_strError <> ERROR_NUMBERTOBIG Then
				m_strError = ERROR_INTEGER_NONNEGATIVE
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object xVal. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ElseIf Round_TSB(CDbl(xVal), 0) >= 0 Then  'check if positive integer including 0
            'variant is not a negative integer, so set error data
            m_lResult = S_FAILED
            m_strError = ERROR_INTEGER_NONNEGATIVE
		End If
		
		'return the result code
		EditRule14 = m_lResult
		
	End Function
	'EditRule15()
	'
	'is the inval a non-positive integer? (includes 0)
	'
	'lblankenship
	'25 sept 2002
	'
	Function EditRule15(ByRef xVal As Object) As Integer
		
		'see if the variant type is not an integer
		If EditRule7(xVal) = S_FAILED Then
			'variant is not an integer, so set error data
			m_lResult = S_FAILED
			If m_strError <> ERROR_NUMBERTOBIG Then
				m_strError = ERROR_INTEGER_POSITIVE
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object xVal. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ElseIf Round_TSB(CDbl(xVal), 0) > 0 Then  'check if POSITIVE integer excluding 0
            'variant is not a nonpositive integer, so set error data
            m_lResult = S_FAILED
            m_strError = ERROR_INTEGER_POSITIVE
		End If
		
		'return the result code
		EditRule15 = m_lResult
		
	End Function
	'EditRule16()
	'
	'is the inval a non-negative integer? (includes 0)
	'
	'lblankenship
	'25 sept 2002
	'
	Function EditRule16(ByRef xVal As Object) As Integer
		
		'see if the variant type is not an integer
		If EditRule7(xVal) = S_FAILED Then
			'variant is not a integer, so set error data
			m_lResult = S_FAILED
			If m_strError <> ERROR_NUMBERTOBIG Then
				m_strError = ERROR_INTEGER_NEGATIVE
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object xVal. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ElseIf Round_TSB(CDbl(xVal), 0) < 0 Then  'check if NEGATIVE integer excluding 0
            'variant is not a nonnegative number, so set error data
            m_lResult = S_FAILED
            m_strError = ERROR_INTEGER_NEGATIVE
		End If
		
		'return the result code
		EditRule16 = m_lResult
		
	End Function
	
	'EditRule17()
	'
	'is the inval a non-zero integer?
	'
	'lblankenship
	'25 sept 2002
	'
	Function EditRule17(ByRef xVal As Object) As Integer
		
		'first let's see if it's a integer
		If (EditRule7(xVal) = S_FAILED) Then
			'if it isn't even a integer, then we can quit now
			m_lResult = S_FAILED
			If m_strError <> ERROR_NUMBERTOBIG Then
				m_strError = ERROR_INTEGER_ZERO
			End If
		Else
			'it is a integer, so let's check for zero-ness
			'UPGRADE_WARNING: Couldn't resolve default property of object xVal. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If (Round_TSB(CDbl(xVal), 0) = 0) Then
                'uh-oh, it's zero...
                m_lResult = S_FAILED
                m_strError = ERROR_INTEGER_ZERO
            End If
		End If
		
		'return the result code
		EditRule17 = m_lResult
		
	End Function
	
	'EditRule18()
	'
	'is the inval a percentage?
	'
	'lblankenship
	'25 sept 2002
	'
	Function EditRule18(ByRef xVal As Object) As Integer
		
		'see if the variant type is not any of the numeric types
		If EditRule1(xVal) = S_FAILED Then
			'variant is not a number, so set error data
			m_lResult = S_FAILED
			m_strError = ERROR_PERCENT
			'UPGRADE_WARNING: Couldn't resolve default property of object xVal. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf CDbl(xVal) >= 1000000000000# Then 
			m_lResult = S_FAILED
			m_strError = ERROR_NUMBERTOBIG
			'UPGRADE_WARNING: Couldn't resolve default property of object xVal. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ElseIf Round_TSB(CDbl(xVal), 2) <> Round_TSB(CDbl(xVal), 4) Then
            'variant is not a valid percent, so set error data
            m_lResult = S_FAILED
            m_strError = ERROR_PERCENT
		End If
		
		'return the result code
		EditRule18 = m_lResult
		
	End Function
	'EditRule19()
	'
	'is the inval a positive percent? (does not include 0)
	'
	'lblankenship
	'25 sept 2002
	'
	Function EditRule19(ByRef xVal As Object) As Integer
		
		'see if the variant type is not a percent
		If EditRule18(xVal) = S_FAILED Then
			'variant is not a percent, so set error data
			m_lResult = S_FAILED
			If m_strError <> ERROR_NUMBERTOBIG Then
				m_strError = ERROR_PERCENT_NONPOSITIVE
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object xVal. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf CDbl(xVal) <= 0 Then  'check if negative percent including 0
			'variant is not a positive percent, so set error data
			m_lResult = S_FAILED
			m_strError = ERROR_PERCENT_NONPOSITIVE
		End If
		
		'return the result code
		EditRule19 = m_lResult
		
	End Function
	'EditRule20()
	'
	'is the inval a negative percent? (does not include 0)
	'
	'lblankenship
	'25 sept 2002
	'
	Function EditRule20(ByRef xVal As Object) As Integer
		
		'see if the variant type is not an percent
		If EditRule18(xVal) = S_FAILED Then
			'variant is not an percent, so set error data
			m_lResult = S_FAILED
			If m_strError <> ERROR_NUMBERTOBIG Then
				m_strError = ERROR_PERCENT_NONNEGATIVE
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object xVal. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ElseIf Round_TSB(CDbl(xVal), 0) >= 0 Then  'check if positive percent including 0
            'variant is not a negative percent, so set error data
            m_lResult = S_FAILED
            m_strError = ERROR_PERCENT_NONNEGATIVE
		End If
		
		'return the result code
		EditRule20 = m_lResult
		
	End Function
	'EditRule21()
	'
	'is the inval a non-positive percent? (includes 0)
	'
	'lblankenship
	'25 sept 2002
	'
	Function EditRule21(ByRef xVal As Object) As Integer
		
		'see if the variant type is not an percent
		If EditRule18(xVal) = S_FAILED Then
			'variant is not an percent, so set error data
			m_lResult = S_FAILED
			If m_strError <> ERROR_NUMBERTOBIG Then
				m_strError = ERROR_PERCENT_POSITIVE
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object xVal. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ElseIf Round_TSB(CDbl(xVal), 0) > 0 Then  'check if negative percent excluding 0
            'variant is not a nonpositive percent, so set error data
            m_lResult = S_FAILED
            m_strError = ERROR_PERCENT_POSITIVE
		End If
		
		'return the result code
		EditRule21 = m_lResult
		
	End Function
	'EditRule22()
	'
	'is the inval a non-negative percent? (includes 0)
	'
	'lblankenship
	'25 sept 2002
	'
	Function EditRule22(ByRef xVal As Object) As Integer
		
		'see if the variant type is not an percent
		If EditRule18(xVal) = S_FAILED Then
			'variant is not a percent, so set error data
			m_lResult = S_FAILED
			If m_strError <> ERROR_NUMBERTOBIG Then
				m_strError = ERROR_PERCENT_NEGATIVE
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object xVal. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ElseIf Round_TSB(CDbl(xVal), 0) < 0 Then  'check if positive percent excluding 0
            'variant is not a nonnegative number, so set error data
            m_lResult = S_FAILED
            m_strError = ERROR_PERCENT_NEGATIVE
		End If
		
		'return the result code
		EditRule22 = m_lResult
		
	End Function
	
	'EditRule23()
	'
	'is the inval a non-zero percent?
	'
	'lblankenship
	'25 sept 2002
	'
	Function EditRule23(ByRef xVal As Object) As Integer
		
		'first let's see if it's a percent
		If (EditRule18(xVal) = S_FAILED) Then
			'if it isn't even a percent, then we can quit now
			m_lResult = S_FAILED
			If m_strError <> ERROR_NUMBERTOBIG Then
				m_strError = ERROR_PERCENT_ZERO
			End If
		Else
			'it is a percent, so let's check for zero-ness
			'UPGRADE_WARNING: Couldn't resolve default property of object xVal. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If (Round_TSB(CDbl(xVal), 0) = 0) Then
                'uh-oh, it's zero...
                m_lResult = S_FAILED
                m_strError = ERROR_PERCENT_ZERO
            End If
		End If
		
		'return the result code
		EditRule23 = m_lResult
		
	End Function
	'EditRule24()
	'
	'is the inval a positive percent <= 100?
	'
	'lblankenship
	'25 sept 2002
	'
	Function EditRule24(ByRef xVal As Object) As Integer
		
		'first let's see if it's a positive percent
		If (EditRule22(xVal) = S_FAILED) Then
			'if it isn't even a positive percent, then we can quit now
			m_lResult = S_FAILED
			m_strError = ERROR_PERCENT_POSITIVE_OVER100
		Else
			'it is a positive percent, so let's check if it's <= 100
			If (EditRule8(xVal) = S_FAILED) Then
				'uh-oh, it's >100
				m_lResult = S_FAILED
				m_strError = ERROR_PERCENT_POSITIVE_OVER100
			End If
		End If
		
		'return the result code
		EditRule24 = m_lResult
		
	End Function
	'EditRule25()
	'
	'is the inval a non-negative currecy? (includes 0)
	'
	'lblankenship
	'9 Oct 2002
	'
	Function EditRule25(ByRef xVal As Object) As Integer
		
		'see if the variant type is not any of the numeric types
		If EditRule6(xVal) = S_FAILED Then
			'variant is not currency, so set error data
			m_lResult = S_FAILED
			If m_strError <> ERROR_NUMBERTOBIG Then
				m_strError = ERROR_CURRENCY_NEGATIVE
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object xVal. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf CDbl(xVal) < 0 Then  'check if NEGATIVE number excluding 0
			'variant is not a nonnegative number, so set error data
			m_lResult = S_FAILED
			m_strError = ERROR_CURRENCY_NEGATIVE
		End If
		
		'return the result code
		EditRule25 = m_lResult
		
	End Function
	'EditRule26()
	'
	'is the inval not empty?
	'
	'lblankenship
	'9 Oct 2002
	'
	Function EditRule26(ByRef xVal As Object) As Integer
		
		'see if the variant type is empty or null
		'UPGRADE_WARNING: Couldn't resolve default property of object xVal. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If IsDbNull(xVal) Or xVal = "" Then
			'variant is empty or is null, so set error data
			m_lResult = S_FAILED
			m_strError = ERROR_STRING_EMPTY
		End If
		
		'return the result code
		EditRule26 = m_lResult
		
	End Function
	'EditRule27()
	'
	'is the inval a positive currency? (does not include 0)
	'
	'lblankenship
	'9 oct 2002
	'
	Function EditRule27(ByRef xVal As Object) As Integer
		
		'see if the variant type is not any of the numeric types
		If EditRule6(xVal) = S_FAILED Then
			'variant is not currency, so set error data
			m_lResult = S_FAILED
			If m_strError <> ERROR_NUMBERTOBIG Then
				m_strError = ERROR_CURRENCY_NONPOSITIVE
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object xVal. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf CDbl(xVal) <= 0 Then  'check if negative number including 0
			'variant is not a positive currency, so set error data
			m_lResult = S_FAILED
			m_strError = ERROR_CURRENCY_NONPOSITIVE
		End If
		
		'return the result code
		EditRule27 = m_lResult
		
	End Function
	'EditRule28()
	'
	'is the inval a positive number of days per year? (does not include 0)
	'
	'lblankenship
	'25 sept 2002
	'
	Function EditRule28(ByRef xVal As Object) As Integer
		
		'see if the variant type is not any of the numeric types
		If EditRule1(xVal) = S_FAILED Then
			'variant is not a number, so set error data
			m_lResult = S_FAILED
			m_strError = ERROR_NUMBER_NONPOSITIVE
			'UPGRADE_WARNING: Couldn't resolve default property of object xVal. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf CDbl(xVal) <= 0 Then  'check if negative number including 0
			'variant is not a positive number, so set error data
			m_lResult = S_FAILED
			m_strError = ERROR_NUMBER_NONPOSITIVE
			'UPGRADE_WARNING: Couldn't resolve default property of object xVal. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf CDbl(xVal) > 366 Then  'check if the number is greater than 366 (include leap year)
			m_lResult = S_FAILED
			m_strError = ERROR_NUMBER_NONPOSITIVE_DAYS_PER_YEAR
		End If
		
		'return the result code
		EditRule28 = m_lResult
		
	End Function
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' public methods
	
	
	
	'validate
	'
	'uses the specified edit rule id to determine which editing function
	'to pass the variant input value off to.  catches the result code
	'passed by the editor function and forwards it to the client.
	'
	'lbailey
	'25 may 2002
	'
    Function Validate(ByRef xInVal As Object, ByVal nEditRule As Short, Optional ByVal strLabel As String = "", Optional ByRef strORMsg As String = "", Optional ByRef fMessage As Object = True) As Integer

        'initialize (reset) result and error members
        'edit rule functions CANNOT reset, only set (or we can't cascade)
        m_strError = ERROR_NONE
        m_lResult = S_OK

        'switch to appropriate edit rule
        Select Case nEditRule
            Case 1
                'is a number
                m_lResult = EditRule1(xInVal)

            Case 2
                'is a date
                m_lResult = EditRule2(xInVal)

            Case 3
                'is a non-zero number
                m_lResult = EditRule3(xInVal)

            Case 4
                'is in the future
                m_lResult = EditRule4(xInVal)

            Case 5
                'is in the past
                m_lResult = EditRule5(xInVal)

            Case 6
                'is currency
                m_lResult = EditRule6(xInVal)

            Case 7
                'is an integer
                m_lResult = EditRule7(xInVal)

            Case 8
                'is a percent <= 100
                m_lResult = EditRule8(xInVal)

            Case 9
                'is a positive number
                m_lResult = EditRule9(xInVal)

            Case 10
                'is a negative number
                m_lResult = EditRule10(xInVal)

            Case 11
                'is a nonpositive number
                m_lResult = EditRule11(xInVal)

            Case 12
                'is a nonnegative number
                m_lResult = EditRule12(xInVal)

            Case 13
                'is a positive integer
                m_lResult = EditRule13(xInVal)

            Case 14
                'is a negative integer
                m_lResult = EditRule14(xInVal)

            Case 15
                'is a nonpositive integer
                m_lResult = EditRule15(xInVal)

            Case 16
                'is a nonnegative integer
                m_lResult = EditRule16(xInVal)

            Case 17
                'is a nonzero integer
                m_lResult = EditRule17(xInVal)

            Case 18
                'is a percent
                m_lResult = EditRule18(xInVal)

            Case 19
                'is a positive percent
                m_lResult = EditRule19(xInVal)

            Case 20
                'is a negative percent
                m_lResult = EditRule20(xInVal)

            Case 21
                'is a nonpositive percent
                m_lResult = EditRule21(xInVal)

            Case 22
                'is a nonnegative percent
                m_lResult = EditRule22(xInVal)

            Case 23
                'is a nonzero percent
                m_lResult = EditRule23(xInVal)

            Case 24
                'is a positive percent <=100
                m_lResult = EditRule24(xInVal)

            Case 25
                'is a nonnegative currency
                m_lResult = EditRule25(xInVal)

            Case 26
                'is not empty or null
                m_lResult = EditRule26(xInVal)
            Case 27
                'is positive currency
                m_lResult = EditRule27(xInVal)

            Case 28
                'is a positive number of day in a year
                m_lResult = EditRule28(xInVal)

            Case Else
                'user requested invalid edit rule
                m_lResult = S_FAILED
                m_strError = ERROR_INVALID_EDIT_RULE

        End Select

        'UPGRADE_WARNING: Couldn't resolve default property of object fMessage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If fMessage Then
            If strLabel <> "" Then
                strLabel = "The value you entered for " & strLabel & " is invalid.  "
            Else
                strLabel = "The value you entered is invalid.  "
            End If

            If m_lResult <> S_OK Then
                If strORMsg <> "" Then
                    MsgBox(strLabel & vbCrLf & strORMsg, MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Invalid Entry")
                Else
                    MsgBox(strLabel & vbCrLf & m_strError, MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Invalid Entry")
                End If
            End If
        End If

        'return the result code
        Validate = m_lResult


    End Function
	
	
	
	'GetErrorDescription()
	'
	'returns a description (string) of the reason why the validation failed.
	'
	'lbailey
	'25 may 2002
	'
	Function GetErrorDescription() As String
		
		'return the string
		GetErrorDescription = m_strError
		
	End Function
	
	Public Function ValuesDiffer(ByRef varValue1 As Object, ByRef varValue2 As Object) As Boolean
		' Comments  : determine if the values are different. Check for nulls 1st
		'
		' Parameters: varValue1 - 1st value to test
		'           : varvalue2 - the second value to test
		' Returns   : Boolean - True if different
		' Created   : 13-Jun-2002
		'-----------------------------------------------------------------------
		Dim fDiffer As Boolean
		
		fDiffer = False
		' Check for Nulls if one value is null and the other is not then exit
		
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If IsDbNull(varValue1) = True Then
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			If IsDbNull(varValue2) Then
				fDiffer = False
			Else
				fDiffer = True
			End If
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		ElseIf IsDbNull(varValue2) = True Then 
			fDiffer = True
			'UPGRADE_WARNING: Couldn't resolve default property of object varValue2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object varValue1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf varValue1 <> varValue2 Then 
			fDiffer = True
		Else
			fDiffer = False
		End If
		
		ValuesDiffer = fDiffer
		
	End Function
	
	Public Function UpdateDirtyFlag(ByRef fIsDirty As Boolean, ByRef varValue1 As Object, ByRef varValue2 As Object) As Boolean
		' Comments  : Detemine if the dirty flag should be set to true (if not already true)
		'           : Never change to false if True, if settofalse then compare and return
		' Parameters: fIsDirty - The presentvalue of the Dirty Flag
		'           : varValue1 - 1st value to test
		'           : varvalue2 - the second value to test
		' Returns   : Boolean - True, if Was Dirty.
		'           :           True, if Now dirty.
		'           :           False, if was false and is still False.
		' Created   : 13-Jun-2002
		'-----------------------------------------------------------------------
		
		If fIsDirty = True Then
			UpdateDirtyFlag = True
		Else
			UpdateDirtyFlag = ValuesDiffer(varValue1, varValue2)
		End If
		
	End Function
	Function Round_TSB(ByVal dblNumber As Double, ByVal intDecimals As Short) As Double
		' Comments  : Rounds a number to a specified number of decimal places (0.5 is rounded up).
		' Parameters: dblNumber - number to round
		'             intDecimals - number of decimal places to round to
		'                        (positive for right of decimal, negative for left)
		' Returns   : Rounded number
		' Modified   :
		' --------------------------------------------------------
		On Error GoTo PROC_ERR
		Dim dblFactor As Double
		Dim dblTemp As Double ' Temp var to prevent rounding problems in INT()
		
		dblFactor = 10 ^ intDecimals
		dblTemp = dblNumber * dblFactor + 0.5
		Round_TSB = Int(CDbl("" & dblTemp)) / dblFactor
		
		Exit Function
		
PROC_ERR: 
		MsgBox("The following error occured: " & ErrorToString())
		Resume Next
	End Function
End Class