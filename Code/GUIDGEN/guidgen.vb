Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("GUIDGen_NET.GUIDGen")> Public Class GUIDGen
	'GUIDGen
	'
	'Creates GUIDs on demand.  Returns them as a string.
	'
	'lbailey
	'9 may 2002
	
	'declare an interface to a windows api function
	'UPGRADE_WARNING: Structure GUID may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
    'Private Declare Function CoCreateGuid Lib "OLE32.DLL" (ByRef pGuid As GUID) As Long
	
	'declate the datatype used by the winapi func
    'Private Structure GUID 'Memory structure used by CoCreateGuid
    'Dim Data1 As Integer
    'Dim Data2 As Short
    'Dim Data3 As Short
    '<VBFixedArray(7)> Dim Data4() As Byte

    'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
    'Public Sub Initialize()
    ' ReDim Data4(7)
    'End Sub
    'End Structure

    'declare a const to be used to measure success
    Private Const S_OK As Short = 0 'Return value from CoCreateGuid



    'GetGUID()
    '
    'Uses the WinAPI to generate a GUID.  Returns the GUID as a string.
    '
    'lbailey
    '9 may 2002
    '
    Public Function GetGUID() As Object
        'Dim lResult As Integer
        'UPGRADE_WARNING: Arrays in structure lGuid may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
        'Dim lGuid As Guid
        'Dim strGuid As String
        'Dim strTemp As String
        'Dim intCtr As Short
        ' lResult = CoCreateGuid(lGuid)
        'If lResult = S_OK Then
        'strTemp = Hex(lGuid.Data1)
        'strGuid = "{" & New String("0", 8 - Len(strTemp)) & strTemp
        ' strTemp = Hex(lGuid.Data2)
        ' strGuid = strGuid & "-" & New String("0", 4 - Len(strTemp)) & strTemp
        'strTemp = Hex(lGuid.Data3)
        'strGuid = strGuid & "-" & New String("0", 4 - Len(strTemp)) & strTemp
        'strTemp = Hex(lGuid.Data4(0))
        'strGuid = strGuid & "-" & New String("0", 2 - Len(strTemp)) & strTemp
        'strTemp = Hex(lGuid.Data4(1))
        'strGuid = strGuid & New String("0", 2 - Len(strTemp)) & strTemp & "-"
        'For intCtr = 2 To 7
        'strTemp = Hex(lGuid.Data4(intCtr))
        'strGuid = strGuid & New String("0", 2 - Len(strTemp)) & strTemp
        'Next
        'strGuid = strGuid & "}"
        'UPGRADE_WARNING: Couldn't resolve default property of object GetGUID. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'GetGUID = strGuid
        'Else
        'UPGRADE_WARNING: Couldn't resolve default property of object GetGUID. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'GetGUID = "ERROR"
        'End If

        'getting error when upgraded...using new format to set guid
        GetGUID = CType((System.Guid.NewGuid), Guid).ToString()
    End Function
End Class