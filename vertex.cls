VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vertex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarxcoord As Single 'local copy
Private mvarycoord As Single 'local copy
Public Property Let ycoord(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ycoord = 5
    mvarycoord = vData
End Property


Public Property Get ycoord() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ycoord
    ycoord = mvarycoord
End Property



Public Property Let xcoord(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.xcoord = 5
    mvarxcoord = vData
End Property


Public Property Get xcoord() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.xcoord
    xcoord = mvarxcoord
End Property



