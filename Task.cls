VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Task"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarcontent As String
Private mvarpriority As String
'lokale Variable(n) zum Zuweisen der Eigenschaft(en)
Private mvarDescription As String 'lokale Kopie
Private mvarchecked As Boolean 'lokale Kopie
Private mvarlabels As Collection 'lokale Kopie
Private mvardue As TaskDue 'lokale Kopie
Private mvarproject_id As String 'lokale Kopie
'lokale Variable(n) zum Zuweisen der Eigenschaft(en)
Private mvarid As String 'lokale Kopie
Public Property Let id(ByVal vData As String)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.id = 5
    mvarid = vData
End Property


Public Property Get id() As String
Attribute id.VB_UserMemId = 0
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.id
    id = mvarid
End Property



Public Property Let project_id(ByVal vData As String)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.project_id = 5
    mvarproject_id = vData
End Property


Public Property Get project_id() As String
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.project_id
    project_id = mvarproject_id
End Property



Public Property Set due(ByVal vData As TaskDue)
'wird beim Zuweisen eines Objekts zu der Eigenschaft auf der linken Seite einer Set-Anweisung verwendet.
'Syntax: Set x.due = Form1
    Set mvardue = vData
End Property


Public Property Get due() As TaskDue
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.due
    Set due = mvardue
End Property



Public Property Set labels(ByVal vData As Collection)
'wird beim Zuweisen eines Objekts zu der Eigenschaft auf der linken Seite einer Set-Anweisung verwendet.
'Syntax: Set x.labels = Form1
    Set mvarlabels = vData
End Property


Public Property Get labels() As Collection
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.labels
    Set labels = mvarlabels
End Property



Public Property Let checked(ByVal vData As Boolean)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.checked = 5
    mvarchecked = vData
End Property


Public Property Get checked() As Boolean
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.checked
    checked = mvarchecked
End Property



Public Property Let Description(ByVal vData As String)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.Description = 5
    mvarDescription = vData
End Property


Public Property Get Description() As String
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.Description
    Description = mvarDescription
End Property




Public Property Let priority(ByVal vData As String)
    mvarpriority = vData
End Property


Public Property Get priority() As String
    priority = mvarpriority
End Property



Public Property Let Content(ByVal vData As String)
    mvarcontent = vData
End Property


Public Property Get Content() As String
    Content = mvarcontent
End Property

