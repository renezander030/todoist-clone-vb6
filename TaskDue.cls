VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TaskDue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarisRecurring As Boolean 'lokale Kopie
'lokale Variable(n) zum Zuweisen der Eigenschaft(en)
Private mvardueDate As String 'lokale Kopie
Public Property Let dueDate(ByVal vData As String)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.dueDate = 5
    mvardueDate = vData
End Property


Public Property Get dueDate() As String
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.dueDate
    dueDate = mvardueDate
End Property



Public Property Let isRecurring(ByVal vData As Boolean)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.isRecurring = 5
    mvarisRecurring = vData
End Property


Public Property Get isRecurring() As Boolean
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.isRecurring
    isRecurring = mvarisRecurring
End Property
