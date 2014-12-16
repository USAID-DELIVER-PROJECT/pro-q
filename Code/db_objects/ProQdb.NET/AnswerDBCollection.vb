Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("AnswerDBCollection_NET.AnswerDBCollection")> Public Class AnswerDBCollection
	'Answerdbcollection.cls
	'
	'manages groups of records
	'
	'lbailey
	'7 june 2002
	
	
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         private members
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	'the dsn used to connect to this db/table
    Private m_strDSN As String = DB_DSN
	'connection to the db
    Private m_objConn As OleDbConnection
	'sp & recordset functions
	Private m_objDBExec As New DBExec
	'the collection of brands
	Private m_cAnswers As New Collection
	
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'                         private methods
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	
	
	'Class_Initialize() (constructor)
	'
	'connect to the db and load the data into memory
	'
	'lbailey
	'1 june 2002
	'
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		
		'get a connection to the db
		'Set m_objConn = GetConnection(m_strDSN)
		
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	
	
	'Class_Terminate() (destructor)
	'
	'cleans up whatever needs to be cleaned up when this object is
	'released.
	'
	'lbailey
	'1 june 2002
	'
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		
		'release the connections to the db
		'UPGRADE_NOTE: Object m_objDBExec may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_objDBExec = Nothing
		
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	
	
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
	'12 june 2002
	'-
    Public Function GetCollection(ByRef strStoredProc As String, ByRef aParams(,) As Object) As Collection

        'go get the recordset
        Dim objRS As DataSet
        Dim strAnswer As String
        Dim i As Integer
        'Dim strMemo As Object

        m_objConn = GetConnection(m_strDSN)
        If aParams Is Nothing Then
            objRS = m_objDBExec.ReturnRSNoParams(strStoredProc, aParams)
        Else
            objRS = m_objDBExec.ReturnRS(m_objConn, strStoredProc, aParams)
        End If

        'for each record
        With objRS
            For i = 0 To .Tables(strStoredProc).Rows.Count - 1 'Do Until .EOF
                Dim objAnswerDB As New AnswerDB

                'create a new object
                strAnswer = IIf(IsDBNull(.Tables(strStoredProc).Rows(i).Item("Answer")), "", .Tables(strStoredProc).Rows(i).Item("Answer"))
                objAnswerDB.SetAnswer(strAnswer)
                objAnswerDB.SetLabel(.Tables(strStoredProc).Rows(i).Item("Label")) '.Fields("Label").Value)
                objAnswerDB.SetParentIndex(.Tables(strStoredProc).Rows(i).Item("ParentIndex")) '.Fields("ParentIndex").Value)
                objAnswerDB.SetActionID(.Tables(strStoredProc).Rows(i).Item("ActionID")) '.Fields("ActionID").Value)
                objAnswerDB.SetNotes(.Tables(strStoredProc).Rows(i).Item("notes")) '.Fields("notes").Value)
                objAnswerDB.SetIsPercent(.Tables(strStoredProc).Rows(i).Item("fPercent")) '.Fields("fPercent").Value)
                If .Tables(strStoredProc).Rows(i).Item("notes") = False Then
                    objAnswerDB.SetComment("")
                Else
                    objAnswerDB.SetComment(.Tables(strStoredProc).Rows(i).Item("comment"))
                End If

                'add the object to the collection
                m_cAnswers.Add(objAnswerDB)

                'release our ref to the obj
                'UPGRADE_NOTE: Object objAnswerDB may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                objAnswerDB = Nothing

                'get the next record
                '.MoveNext()
            Next 'Loop
        End With

        GetCollection = m_cAnswers
        'UPGRADE_NOTE: Object m_objConn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        m_objConn = Nothing
    End Function
End Class