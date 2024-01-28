' ---------- ---------- ---------- ---------- ---------- ---------- ---------- ---------- ---------- ----------
Sub Main()

    X_Offset = 20 / 25.4
    Y_Offset = 20 / 25.4

    Call Set_A3_Landscape
    
    PageWidth = Get_PageWidth() / 25.4
    Debug.Print "PageWidth ="; PageWidth
    PageHeight = Get_PageHeight() / 25.4
    Debug.Print "PageHeight ="; PageHeight
    
    Input_Filename = "t_shapes"
    
    Call Create_TemporaryCopy(Input_Filename)
    Call Add_Shapes(Input_Filename, X_Offset, Y_Offset, PageWidth, PageHeight)
    
    Call Add_Connectors(Input_Filename)
    
    Call Delete_TemporaryCopy(Input_Filename)
End Sub
' ---------- ---------- ---------- ---------- ---------- ---------- ---------- ---------- ---------- ----------
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
    Get_PageWidth = Int(Left(Application.ActivePage.PageSheet.CellsSRC(visSectionObject, visRowPage, visPageWidth).FormulaU, 3))
End Function
' ---------- ---------- ---------- ---------- ---------- ---------- ---------- ---------- ---------- ----------
Function Get_PageHeight()
    Get_PageHeight = Int(Left(Application.ActivePage.PageSheet.CellsSRC(visSectionObject, visRowPage, visPageHeight).FormulaU, 3))
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
    Set ROWS = XDoc.selectnodes("//topic/body/simpletable/strow")

    Dim ShapeName As String
    Dim ShapeLeft As Double
    Dim ShapeTop As Double
    Dim ShapeWidth As Double
    Dim ShapeHeight As Double
    
    For Each Value In ROWS
                
        ShapeName = Value.ChildNodes(0).Text
        ShapeColumn = Value.ChildNodes(1).Text
        ShapeRow = Value.ChildNodes(2).Text
        ShapeWidth = Value.ChildNodes(3).Text / 25.4
        ShapeHeight = Value.ChildNodes(4).Text / 25.4
        ShapeConnectsTo = Value.ChildNodes(5).Text
                    
        ShapeLeft = (X_Offset * ShapeColumn) + (ShapeWidth * (ShapeColumn - 1))
        ShapeTop = PageHeight - (ShapeHeight + Y_Offset) * ShapeRow
        
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
    Set ROWS = XDoc.selectnodes("//topic/body/simpletable/strow")

    ShapeCounter = 1
    ' get to-from shape-names
    Debug.Print "- - - - - - - - - - -"
    Debug.Print "Connecting:"
    
    For Each Shape In ROWS
    
        From_Name = Shape.ChildNodes(0).Text
        To_Name = Shape.ChildNodes(5).Text
        
        If Not (To_Name = "none") Then
            Debug.Print From_Name + "->" + To_Name
        
            From_ID = Get_ID(From_Name)
            Debug.Print "  From_ID =" + Str(From_ID)
            
            To_ID = Get_ID(To_Name)
            Debug.Print "  To_ID =" + Str(To_ID)
            
            Call Add_Connector(From_ID, To_ID)
        End If
        ShapeCounter = ShapeCounter + 1
    Next Shape
    Debug.Print "- - - - - - - - - - -"

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
