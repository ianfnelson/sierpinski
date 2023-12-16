Attribute VB_Name = "modMain"
Option Explicit
Dim colVertices As New Collection
Dim vertex As vertex
Const screenborder = 500
Dim i As Integer

Sub Main()
    
    ' show and maximise main form '
    frmMain.Show
    
    ' set position of vertices '
    Set colVertices = New Collection
    Call SetVertices(colVertices)
        
    ' plot vertices '
    DrawVertices
          
    ' start from top vertex, and pick one of the other two to move towards
    Dim curPoint As vertex
    Set curPoint = New vertex
    curPoint.xcoord = colVertices.Item(1).xcoord
    curPoint.ycoord = colVertices.Item(1).ycoord
    
    i = Int(Rnd(1) * 2) + 2
    
    ' keep moving from curPoint to halfway to a random vertex '
    Do
        curPoint.xcoord = (curPoint.xcoord + colVertices(i).xcoord) / 2
        curPoint.ycoord = (curPoint.ycoord + colVertices(i).ycoord) / 2
        Call PlotPoint(curPoint)
        DoEvents
        
        i = Int(Rnd(1) * 3) + 1
    Loop
          
End Sub

Sub SetVertices(colVertices As Collection)

    Dim sideWidth As Single
    sideWidth = (Screen.Height - (2 * screenborder)) * (2 / Sqr(3))
    
    ' top vertex '
    Set vertex = New vertex
    vertex.xcoord = (Screen.Width / 2)
    vertex.ycoord = screenborder
    colVertices.Add vertex
    
    ' bottom left vertex '
    Set vertex = New vertex
    vertex.xcoord = (Screen.Width / 2) - (sideWidth / 2)
    vertex.ycoord = Screen.Height - screenborder
    colVertices.Add vertex
    
    ' bottom right vertex '
    Set vertex = New vertex
    vertex.xcoord = (Screen.Width / 2) + (sideWidth / 2)
    vertex.ycoord = Screen.Height - screenborder
    colVertices.Add vertex

End Sub
Sub DrawVertices()

    For i = 1 To 3
    
        Call PlotPoint(colVertices.Item(i))
    
    Next i

End Sub


Sub PlotPoint(vertex)

    frmMain.PSet (vertex.xcoord, vertex.ycoord), RGB(255, 255, 255)

End Sub
