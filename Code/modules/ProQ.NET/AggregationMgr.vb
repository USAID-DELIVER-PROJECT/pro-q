Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("AggregationMgr_NET.AggregationMgr")> Public Class AggregationMgr
	
	' Aggregation Manager
	' Controls the functions on the Backend for returning lists of Aggregates etc.
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'' Public Functions
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	' Get the Aggregate List
    Public Function GetList() As DataSet 'ADODB.Recordset

        Dim objAggregation As New ProQdb.AggregationDB
        Dim strProc As String
        Dim aParams(,) As Object

        strProc = "qlkpAggregations"

        GetList = objAggregation.ReturnRS(strProc, aParams)

        'Set objAggregation = Nothing


    End Function
    Public Function GetCollection() As DataSet 'Collection

        Dim objAggCol As New ProQdb.AggregationCollectionDB
        Dim strProc As String
        Dim aParams(,) As Object

        strProc = "select * from qlkpAggregationsFull"

        GetCollection = objAggCol.GetCollection(strProc, aParams)

        'Set objAggregation = Nothing


    End Function
	
	
	
	' Create
	
	
	' Delete (From the DB)
	Public Function Delete(ByRef strID As String, Optional ByRef fAsk As Boolean = False) As Short
		
		On Error GoTo Proc_Err
		
		Dim strProc As String
		Dim objAggregation As New ProQdb.AggregationDB
		
		If fAsk = True Then
			If MsgBox("Are you sure you want to delete the selected quantification? ", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "DELETE?") = MsgBoxResult.No Then
				Delete = S_FAILED
				Exit Function
			End If
		End If
		
		'Delete the Quantification and Protocols related to the Quantification
		DeleteProtocols(strID)
		
		objAggregation.Delete(strID)
		'UPGRADE_NOTE: Object objAggregation may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objAggregation = Nothing
		
		Delete = S_OK
Proc_Exit: 
		On Error Resume Next
		Exit Function
		
Proc_Err: 
		Delete = Err.Number
		Resume Proc_Exit
	End Function
	
	Sub DeleteProtocols(ByRef strID As String)
		' Comments  : With the selected Aggregation ID, Scroll through each Quantification and
		'           : delete the protocol
		' Parameters: strID - the Aggregation ID
		' Created   : 21-Aug-2002 JSL
		'------------------------------------------------------------------------
		
		Dim objColQ As New ProQdb.QuantificationDBCollection
		
        Dim objDBQ As New ProQdb.QuantificationDB
        Dim mColQ As Collection
		Dim objProtocolDB As New ProQdb.ProtocolDB
		
		Dim strStoredProc As String
        Dim aParams(1, 1) As Object
        Dim i As Integer
		
		strStoredProc = SP_GET_QUANTIFICATIONS_BY_AGGREGATION
		
		'UPGRADE_WARNING: Couldn't resolve default property of object aParams(0, 0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(0, 0) = New Guid(strID)
		'UPGRADE_WARNING: Couldn't resolve default property of object aParams(0, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aParams(0, 1) = DbType.Guid
		
		mColQ = objColQ.GetCollection(strStoredProc, aParams)
		
        For Each objDBQ In mColQ
            'For i = 1 To mColQ.Count
            'm_objCurrentQuantification = m_colQuantifications.Add(CType((mCol.Tables("tblQuantification").Rows(i).Item("GuidID")), Guid).ToString()) 'objDB.GetID)
            'm_objCurrentQuantification.Load(CType((mCol.Tables("tblQuantification").Rows(i).Item("GuidID")), Guid).ToString()) 'objDB.GetID)
            objProtocolDB.Load(objDBQ.GetProtocolID)
            objProtocolDB.Delete()
            'Next
        Next objDBQ

        'UPGRADE_NOTE: Object objProtocolDB may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        objProtocolDB = Nothing
        'UPGRADE_NOTE: Object objColQ may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        objColQ = Nothing
        'UPGRADE_NOTE: Object mColQ may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        mColQ = Nothing





    End Sub
End Class