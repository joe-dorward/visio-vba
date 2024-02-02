Attribute VB_Name = "Shaper_v4_04"
' Shaper: v4.04
' Written by: Joe Dorward
' Started: 01/02/24
' ---------- ---------- ---------- ---------- ---------- ---------- ---------- ---------- ---------- ----------
Sub Main()

    Input_Filename = "t_shapes"
    Call Create_TemporaryCopy(Input_Filename)

    Call DeleteAll
    'Call Set_A4_Portrait

    X_Offset = to_inches(20)
    Y_Offset = to_inches(20)

    'MsgBox "Sizing, and orienting...", vbOKOnly, "Set_A3_Landscape()  "
    Call Set_A3_Landscape

    PageWidth = Get_PageWidth()
    PageHeight = Get_PageHeight()
    Debug.Print
    Debug.Print "= = = = = = = = = = = = = = = = = = = ="
    Debug.Print "Running Main()"
    Debug.Print "  PageWidth ="; PageWidth
    Debug.Print "  PageHeight ="; PageHeight

    Call Add_Shapes(Input_Filename, X_Offset, Y_Offset, PageWidth, PageHeight)

    Call Add_Connectors(Input_Filename)

    Call MoveShapes(Input_Filename)
    
    ActiveWindow.DeselectAll
    Call Delete_TemporaryCopy(Input_Filename)
End Sub
' ---------- ---------- ---------- ---------- ---------- ---------- ---------- ---------- ---------- ----------
Sub DeleteAll()
    ' delete all shapes

    If (Application.ActivePage.Shapes.Count > 0) Then
        MsgBox "Deleting all...", vbOKOnly, "DeleteAll()"
        Application.ActiveWindow.SelectAll
        Application.ActiveWindow.Selection.Delete
    End If
    
End Sub
Function to_inches(millimeters)
    ' converts numbers to inches
    to_inches = millimeters / 25.4
End Function
' ---------- ---------- ---------- ---------- ---------- ---------- ---------- ---------- ---------- ----------
Sub Set_A4_Portrait()
    ' set page to be A4 and Portrait

    Dim UndoScopeID1 As Long
    UndoScopeID1 = Application.BeginUndoScope("Page Setup")
    Application.ActivePage.Background = False
    Application.ActivePage.BackPage = ""
    Application.ActivePage.PageSheet.CellsSRC(visSectionObject, visRowPage, visPageWidth).FormulaU = "210 mm"
    Application.ActivePage.PageSheet.CellsSRC(visSectionObject, visRowPage, visPageHeight).FormulaU = "297 mm"
    Application.ActivePage.PageSheet.CellsSRC(visSectionObject, visRowPrintProperties, visPrintPropertiesPageOrientation).FormulaU = "0"
    Application.EndUndoScope UndoScopeID1, True
    
End Sub
Sub Set_A3_Landscape()
    ' set page to be A3 and Landscape

    Dim UndoScopeID1 As Long
    UndoScopeID1 = Application.BeginUndoScope("Page Setup")
    Application.ActivePage.Background = False
    Application.ActivePage.BackPage = ""
    Application.ActivePage.PageSheet.CellsSRC(visSectionObject, visRowPage, visPageWidth).FormulaU = "420 mm"
    Application.ActivePage.PageSheet.CellsSRC(visSectionObject, visRowPage, visPageHeight).FormulaU = "297 mm"
    Application.ActivePage.PageSheet.CellsSRC(visSectionObject, visRowPrintProperties, visPrintPropertiesPageOrientation).FormulaU = "1"
    Application.EndUndoScope UndoScopeID1, True
    
End Sub
' ---------- ---------- ---------- ---------- ---------- ---------- ---------- ---------- ---------- ----------
Function Get_PageWidth()
    Get_PageWidth = to_inches(Int(Left(Application.ActivePage.PageSheet.CellsSRC(visSectionObject, visRowPage, visPageWidth).FormulaU, 3)))
End Function
' ---------- ---------- ---------- ---------- ---------- ---------- ---------- ---------- ---------- ----------
Function Get_PageHeight()
    Get_PageHeight = to_inches(Int(Left(Application.ActivePage.PageSheet.CellsSRC(visSectionObject, visRowPage, visPageHeight).FormulaU, 3)))
End Function
' ---------- ---------- ---------- ---------- ---------- ---------- ---------- ---------- ---------- ----------
Sub Create_TemporaryCopy(Filename)
    ' create a temporary copy of the 'shapes' file, while skipping the first two lines of DITA XML

    Open Filename + ".dita" For Input As #1

    Line Input #1, FirstLine ' skip first-line
    Line Input #1, SecondLine ' skip second-line

    ' read rest of 'shapes' file
    Do Until EOF(1)
        Line Input #1, Linefromfile
        DITA_XML = DITA_XML + Linefromfile + vbNewLine
    Loop
    
    Close #1
    
    ' write temporary copy
    Open Filename + "_temp.dita" For Output As #2
    Print #2, DITA_XML
    Close #2
End Sub
' ---------- ---------- ---------- ---------- ---------- ---------- ---------- ---------- ---------- ----------
Sub Add_Shapes(Filename, X_Offset, Y_Offset, PageWidth, PageHeight)
    ' read shape values from 'shapes' file

    Dim XDoc As Object
    Set XDoc = CreateObject("MSXML2.DOMDocument")
    XDoc.async = False: XDoc.validateOnParse = False
    XDoc.Load (Filename + "_temp.dita")
    Set Rows = XDoc.selectnodes("//topic/body/simpletable/strow")

    Dim ShapeName As String
    Dim ShapeLeft As Double
    Dim ShapeTop As Double
    Dim ShapeWidth As Double
    Dim ShapeHeight As Double
    
    For Each Value In Rows
                
        ShapeName = Value.ChildNodes(0).Text
        ShapeColumn = Value.ChildNodes(1).Text
        ShapeRow = Value.ChildNodes(2).Text
        ShapeWidth = Value.ChildNodes(3).Text / 25.4
        ShapeHeight = Value.ChildNodes(4).Text / 25.4
        ShapeConnectsTo = Value.ChildNodes(5).Text
                    
        ShapeLeft = (X_Offset * ShapeColumn) + (ShapeWidth * (ShapeColumn - 1))
        ShapeTop = PageHeight - (ShapeHeight + Y_Offset) * ShapeRow
        
        MsgBox "Adding shapes..." + Chr(10) + "    " + ShapeName, vbOKOnly, "Add_Shapes()"
        Call Create_Shape(ShapeLeft, ShapeTop, ShapeLeft + ShapeWidth, ShapeTop + ShapeHeight, ShapeName)

    Next Value
    
    ActiveWindow.DeselectAll
    
    Set XDoc = Nothing
    
End Sub
' ---------- ---------- ---------- ---------- ---------- ---------- ---------- ---------- ---------- ----------
Sub Create_Shape(xStart, yStart, xEnd, yEnd, Content)
    ' create shapes on VISIO page

    Dim vsoShape As Visio.Shape
    Set vsoShape = ActivePage.DrawRectangle(xStart, yStart, xEnd, yEnd)
    vsoShape.Text = Content
    
    vsoShape.BringToFront

End Sub
' ---------- ---------- ---------- ---------- ---------- ---------- ---------- ---------- ---------- ----------
Sub Add_Connectors(Filename)
    ' read 'connector' values from 'shapes' file
    
    Dim XDoc As Object
    Set XDoc = CreateObject("MSXML2.DOMDocument")
    XDoc.async = False: XDoc.validateOnParse = False
    XDoc.Load (Filename + "_temp.dita")
    Set Rows = XDoc.selectnodes("//topic/body/simpletable/strow")
   
    ' get to-from shape-names
    For Each Shape In Rows
    
        From_Name = Shape.ChildNodes(0).Text
        To_Name = Shape.ChildNodes(5).Text
        
        If Not Len(To_Name) = 0 Then
            From_ID = Get_ID(From_Name)
            To_ID = Get_ID(To_Name)
            
            Debug.Print "- - - - - - - - - - - - - - - - - - - -"
            Debug.Print From_Name; "->"; To_Name; " (adding connection)"
            Debug.Print "  From_ID =" + Str(From_ID)
            Debug.Print "  To_ID =" + Str(To_ID)
                            
            MsgBox "Connecting shapes..." + Chr(10) + "    " + From_Name + "->" + To_Name, vbOKOnly, "Add_Connector()"
            Call Add_Connector(From_ID, To_ID)
        End If
        
    Next Shape
    Debug.Print "- - - - - - - - - - - - - - - - - - - -"
    
    ActiveWindow.DeselectAll
    
End Sub
' ---------- ---------- ---------- ---------- ---------- ---------- ---------- ---------- ---------- ----------
Function Get_ID(ShapeName)
    ' returns the ID of a named-shape

    For Each Shape In Application.ActiveWindow.Page.Shapes
    
        If (Shape.Text = ShapeName) Then
            Shape_ID = Shape.ID
        End If
            
    Next Shape
    
    Get_ID = Shape_ID
    
End Function
' ---------- ---------- ---------- ---------- ---------- ---------- ---------- ---------- ---------- ----------
Sub Add_Connector(FromID, ToID)
    ' create connectors on VISIO page

    ShapesCount = Application.ActiveWindow.Page.Shapes.Count

    Application.Windows.ItemEx(ActiveDocument.Name).Activate
    Application.ActiveWindow.Page.Drop Application.Documents.Item("BASIC_M.VSS").Masters.ItemU("Dynamic connector"), 3.700787, 8.070866

    Dim UndoScopeID1 As Long
    UndoScopeID1 = Application.BeginUndoScope("Size Object")
    Dim vsoCell1 As Visio.Cell
    Dim vsoCell2 As Visio.Cell
    Set vsoCell1 = Application.ActiveWindow.Page.Shapes.ItemFromID(ShapesCount + 1).CellsU("BeginX")
    Set vsoCell2 = Application.ActiveWindow.Page.Shapes.ItemFromID(FromID).CellsSRC(1, 1, 0)
    vsoCell1.GlueTo vsoCell2
    Application.EndUndoScope UndoScopeID1, True

    Dim UndoScopeID2 As Long
    UndoScopeID2 = Application.BeginUndoScope("Size Object")
    Dim vsoCell3 As Visio.Cell
    Dim vsoCell4 As Visio.Cell
    Set vsoCell3 = Application.ActiveWindow.Page.Shapes.ItemFromID(ShapesCount + 1).CellsU("EndX")
    Set vsoCell4 = Application.ActiveWindow.Page.Shapes.ItemFromID(ToID).CellsSRC(1, 1, 0)
    vsoCell3.GlueTo vsoCell4
    Application.EndUndoScope UndoScopeID2, True

    Dim UndoScopeID3 As Long
    UndoScopeID3 = Application.BeginUndoScope("Line Properties")
    Application.ActiveWindow.Page.Shapes.ItemFromID(ShapesCount + 1).CellsSRC(visSectionObject, visRowLine, visLineEndArrow).FormulaU = "13"
    Application.EndUndoScope UndoScopeID3, True
    
    ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemFromID(ShapesCount + 1), visSelect
    Application.ActiveWindow.Selection.SendToBack
    
    'Debug.Print Application.ActiveWindow.Page.Shapes.Count
End Sub
' ---------- ---------- ---------- ---------- ---------- ---------- ---------- ---------- ---------- ----------
Sub Delete_TemporaryCopy(Filename)
    ' delete a temporary copy of the 'shapes' file
    
    'MsgBox "Killing " + Filename + "_temp.dita"
    Kill Filename + "_temp.dita"
End Sub
' ---------- ---------- ---------- ---------- ---------- ---------- ---------- ---------- ---------- ----------
Sub MoveShapes(Filename)
    ' read 'move' values from 'shapes' file
    
    Dim XDoc As Object
    Set XDoc = CreateObject("MSXML2.DOMDocument")
    XDoc.async = False: XDoc.validateOnParse = False
    XDoc.Load (Filename + "_temp.dita")
    Set Rows = XDoc.selectnodes("//topic/body/simpletable/strow")
    
    ' get shape-names
    For Each Shape In Rows
    
        ShapeName = Shape.ChildNodes(0).Text
        Left_Right = Shape.ChildNodes(6).Text
        Up_Down = Shape.ChildNodes(7).Text

        If Len(Left_Right) > 0 Then
            Call Move_LeftRight(ShapeName, Left_Right)
        End If
        
        If Len(Up_Down) > 0 Then
            Call Move_UpDown(ShapeName, Up_Down)
        End If

    Next Shape
    
End Sub
Sub Move_LeftRight(ShapeName, MoveBy)
    ' left / right

    If (MoveBy > 0) Then
        LeftRight = " (right)"
    Else
        LeftRight = " (left)"
    End If
    
    MsgBox "Moving shapes..." + Chr(10) + "    " + ShapeName + LeftRight, vbOKOnly, "Move_Shape()"
    ActiveWindow.DeselectAll
    ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemFromID(Get_ID(ShapeName)), visSelect
    Application.ActiveWindow.Selection.Move (MoveBy / 25.4), 0#

End Sub
Sub Move_UpDown(ShapeName, MoveBy)
    ' up / down
    
    If (MoveBy > 0) Then
        UpDown = " (up)"
    Else
        UpDown = " (down)"
    End If
    
    MsgBox "Moving shapes..." + Chr(10) + "    " + ShapeName + UpDown, vbOKOnly, "Move_Shape()"
    ActiveWindow.DeselectAll
    ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemFromID(Get_ID(ShapeName)), visSelect
    Application.ActiveWindow.Selection.Move 0#, (MoveBy / 25.4)

End Sub

