Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("ProgramDBCollection_NET.ProgramDBCollection")> Public Class ProgramDBCollection
	'+
	'ProgramDBCollection.cls
	'
	'manages groups of records
	'
	'jleiner
	'2 Aug 2002
	'-
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         private members
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	
	'whether or not the object has been constructed
    Private m_objConn As OleDbConnection
	Private m_objDBExec As New DBExec
	
	'the collection of Costs
	Private m_cDB As New Collection
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         public members
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	'hide everything from the user
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         private methods
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         public methods
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	
	'+
	'GetCollection()
	'
	'executes the specified stored procedure and returns the results in a
	'collection of objects
	'
	'lbailey
	'1 june 2002
	'-
    Public Function GetCollection(ByRef strStoredProc As String, ByRef aParams(,) As Object) As Collection

        On Error Resume Next

        'go get the recordset
        Dim objRS As DataSet
        Dim i As Integer

        m_objConn = GetConnection(DB_DSN)

        If aParams Is Nothing Then
            objRS = m_objDBExec.ReturnRSNoParams(strStoredProc, aParams)
        Else
            objRS = m_objDBExec.ReturnRS(m_objConn, strStoredProc, aParams)
        End If


        'for each record
        With objRS
            For i = 0 To .Tables(strStoredProc).Rows.Count - 1 'Do Until .EOF
                Dim objDB As New ProgramDB
                'put the object in the collection

                'stuff all of the record's properties into the object
                objDB.SetID(CType((.Tables(strStoredProc).Rows(i).Item("GuidID")), Guid).ToString()) '.Fields("GuidID").Value)
                objDB.SetName(.Tables(strStoredProc).Rows(i).Item("strName")) '.Fields("strName").Value)
                objDB.SetNotes(IIf(IsDBNull(.Tables(strStoredProc).Rows(i).Item("memNotes")), "", .Tables(strStoredProc).Rows(i).Item("memNotes")))

                'add the object to the collection
                m_cDB.Add(objDB)

                'UPGRADE_NOTE: Object objDB may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                objDB = Nothing

                'get the next record
                '.MoveNext()
            Next 'Loop
            '.Close()
        End With

        GetCollection = m_cDB
        'UPGRADE_NOTE: Object objRS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        objRS = Nothing
        'UPGRADE_NOTE: Object m_objConn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        m_objConn = Nothing
    End Function
    Public Function GetDataset(ByRef strStoredProc As String, ByRef aParams(,) As Object) As DataSet

        On Error Resume Next

        'go get the recordset
        Dim objRS As DataSet
        Dim i As Integer

        m_objConn = GetConnection(DB_DSN)

        If aParams Is Nothing Then
            objRS = m_objDBExec.ReturnRSNoParams(strStoredProc, aParams)
        Else
            objRS = m_objDBExec.ReturnRS(m_objConn, strStoredProc, aParams)
        End If

        GetDataset = objRS
        'UPGRADE_NOTE: Object objRS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        objRS = Nothing
        'UPGRADE_NOTE: Object m_objConn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        m_objConn = Nothing
    End Function
End Class