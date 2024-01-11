# Items-01.vsd
The following VBA code example presumes that you know enough about using and running VBA to use it as-is. This code will execute when its VISIO file opens, and 'expects' to 'find' the 'Items-01.xml' (above) file in the same folder.
' ---------- ---------- ---------- ---------- ---------- ---------- ---------- ---------- ---------- ----------
Private Sub Document_DocumentOpened(ByVal doc As IVDocument)
  
    Dim XDoc As Object
    Set XDoc = CreateObject("MSXML2.DOMDocument")
    XDoc.async = False: XDoc.validateOnParse = False
    XDoc.Load ("Items-01.xml")
    Set Items = XDoc.documentElement
    
    Dim Name As String
    Dim Left As Double
    Dim Bottom As Double
    Dim Width As Double
    Dim Height As Double
    
    For Each Item In Items.childNodes
        MsgBox ("ADD A SHAPE?")
      
        Name = Item.childNodes(0).Text
        Left = Item.childNodes(1).Text
        Bottom = Item.childNodes(2).Text
        Width = Item.childNodes(3).Text
        Height = Item.childNodes(4).Text
    
        Call Create_Shape(Left, Bottom, Left + Width, Bottom + Height, Name)
    Next Item
    
    Set XDoc = Nothing
End Sub
' ---------- ---------- ---------- ---------- ---------- ---------- ---------- ---------- ---------- ----------
Sub Create_Shape(xStart As Double, yStart As Double, xEnd As Double, yEnd As Double, Name As String)

    Dim vsoShape As Visio.Shape

    Set vsoShape = ActivePage.DrawRectangle(xStart, yStart, xEnd, yEnd)
    vsoShape.Text = Name

End Sub
