VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParentChild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'=========================================================================================
'  ParentChild Class
'  Parent child method to load items from a database table
'=========================================================================================
'  Created By: Behrooz Sangani
'  Published Date: 7/9/2005
'  E-Mail: sangani@gmail.com
'  Legal Copyright: Behrooz Sangani © 7/9/2005
'  Free for personal use but do not remove the copyright notice.
'  For comercial use please contact the author.
'=========================================================================================
'
'Parent Child items in a database table is a well-known issue. Hierarchy in just
'one table of the database without complicated relations between tables
'can be obtained by this method. A sample for this is:
'
'   -- Employee Hierarchy --
'
' Field     ID      Name                ParentID
'           --      --------------      --------
'           1       Chris Davison       0
'           2       John Graves         1
'           3       Jerry Graves        1
'           4       Kate Jones          2
'           5       Johny Cash          3
'           6       Travis Pane         3
'           7       Cathy Charles       5
'
'
'   Chris Davison
'   |
'   -------- John Graves
'   |      |
'   |      -------- Kate Jones
'   |
'   -------- Jerry Graves
'          |
'          -------- Johny Cash
'          |      |
'          |      -------- Cathy Charles
'          |
'          -------- Travis Pane
'
'This class simplifies loading this type of data. Fast loading
'of nodes by filtering recordset data and flexibility to load
'any database using ADODB has been considered here.
'
'Please review the sample to see how to use this class
'
'=========================================================================================

'local variable(s) to hold property value(s)
Private CN As ADODB.Connection              'adodb connection object
Private RS As ADODB.Recordset               'main recordset object

Private mvarpkField As String               'Name of the primary key field (ie. "ID")
Private mvarpiField As String               'Name of the Parent ID field (ie. "PI")
Private mvarDatabasePath As String          'optional path to the database file
Private mvarChildFields As Recordset        'recordset which holds child fields of a parent field
Private mvarParentField As Recordset        'recordset which holds the parent field of a field
Private mvarOpenRS As Recordset             'last openned recordset with no filters
Private mvarError As String                 'last error string

'retuns the parent field of a field as a recordset
'returns an empty recordset if item is a top level element
'vField  =  numeric value (primary key value) of the field we want to get its parent
Public Property Get ParentField(vField As Double, Optional pkField As String, Optional piField As String) As Recordset

    On Error GoTo errHandler

    Dim ParentID As String    'Parent ID field value for the given field
    
    'load properties if no optional values are given
    If pkField = "" Then pkField = mvarpkField
    If piField = "" Then piField = mvarpiField

    'filter to get the parent id for the given field
    RS.Filter = pkField & "=" & CStr(vField)
    ParentID = RS.Fields(piField)
    
    'filter to get the parent field
    RS.Filter = pkField & "=" & ParentID

    'set the recordset to the filtered to parent recordset
    Set mvarParentField = RS

    Set ParentField = mvarParentField

    mvarError = ""  'no error
    Exit Property
errHandler:
    'report error
    mvarError = "Get ParentField::" & Err.Description

End Property
''''''''''''''''''''''''''''''''''

'returns the child fields of a field as a recordset
'vField  =  numeric value (primary key value) of the field we want to collect its children
Public Property Get ChildFields(vField As Double, Optional piField As String) As Recordset
Attribute ChildFields.VB_UserMemId = 0
    
    On Error GoTo errHandler
    
    'load properties if no optional values are given
    If piField = "" Then piField = mvarpiField

    'filter to get the child fields for the given field
    RS.Filter = piField & "=" & CStr(vField)
    
    'set the recordset to the filtered to children recordset
    Set mvarChildFields = RS

    Set ChildFields = mvarChildFields

    mvarError = ""  'no error
    Exit Property
errHandler:
    'report error
    mvarError = "Get ChildFields::" & Err.Description
    
End Property
''''''''''''''''''''''''''''''''''

'returns the opened recordset
Public Property Get OpenRS() As Recordset
    Set OpenRS = mvarOpenRS
End Property
''''''''''''''''''''''''''''''''''

'executes an sql string to open a recordset
' strSql   =    is the sql string to open an specific table
Public Function OpenRecordset(strSql As String, Optional CursorType As CursorTypeEnum = adOpenDynamic, Optional LockType As LockTypeEnum = adLockOptimistic) As Boolean

    On Error GoTo errHandler
    
    'close recordset if already open
    If RS.State = 1 Then RS.Close
    
    'open the recordset using given values
    RS.Open strSql, CN, CursorType, LockType

    'set the main recordset
    Set mvarOpenRS = RS

    OpenRecordset = True
    
    mvarError = "" 'no error
    Exit Function
errHandler:
    'report error
    mvarError = "OpenRecordset::" & Err.Description
    
End Function
''''''''''''''''''''''''''''''''''

'first of all connect to db
'   you can set a path to an access .mdb file or just write
'   your own connection string (ie. sql server connection)
Public Function Connect(Optional DatabasePath As String, Optional ConnectionString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= %DatabasePath% ;Persist Security Info=False") As Boolean

    On Error GoTo errHandler

    'Open the connection
    Set CN = New Connection
    Set RS = New Recordset
    
    'load the property if the optional value is empty
    If DatabasePath = "" Then DatabasePath = mvarDatabasePath

    ConnectionString = Replace(ConnectionString, "%DatabasePath%", DatabasePath)

    'open the database
    CN.Open ConnectionString

    CN.CursorLocation = adUseClient

    Connect = True
    
    mvarError = "" 'no error
    Exit Function
errHandler:
    'report error
    mvarError = "Connect::" & Err.Description

End Function
''''''''''''''''''''''''''''''''''

'path to the database to be opened
Public Property Let DatabasePath(ByVal vData As String)
    mvarDatabasePath = vData
End Property
Public Property Get DatabasePath() As String
    DatabasePath = mvarDatabasePath
End Property
''''''''''''''''''''''''''''''''''

'name of the parent id field (ie. "PI")
Public Property Let piField(ByVal vData As String)
    mvarpiField = vData
End Property
Public Property Get piField() As String
    piField = mvarpiField
End Property
''''''''''''''''''''''''''''''''''
 
'name of the primary key field (ie. "ID")
Public Property Let pkField(ByVal vData As String)
    mvarpkField = vData
End Property
Public Property Get pkField() As String
    pkField = mvarpkField
End Property
''''''''''''''''''''''''''''''''''

'last error string
Public Property Get LastError() As String
    LastError = mvarError
End Property
''''''''''''''''''''''''''''''''''

'close open connection and destroy objects
Private Sub Class_Terminate()

    If RS.State = 1 Then RS.Close
    If CN.State = 1 Then CN.Close

    Set RS = Nothing
    Set CN = Nothing

End Sub
