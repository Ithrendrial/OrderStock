VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub A3ColouredPaper_Click()
    Call UpdateFromButton("C10", "D10")
End Sub

Private Sub BubbleWrap_Click()
    Call UpdateFromButton("C37", "D37")
End Sub

Private Sub LaminatingPouches_Click()
    Call UpdateFromButton("C38", "D38")
End Sub

Private Sub Rubberbands_Click()
    Call UpdateFromButton("C39", "D39")
End Sub

Private Sub Staples_Click()
    Call UpdateFromButton("C40", "D40")
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = Me.ListObjects("Stock")
    On Error GoTo 0

    If Not tbl Is Nothing Then
        If Not Intersect(Target, tbl.DataBodyRange) Is Nothing Then
            Call GenerateOrder
        End If
    End If
End Sub

Sub UpdateFromButton(ByVal codeAddress As String, ByVal nameAddress As String)
    Dim matchingRow As ListRow
    Dim productCode As String
    Dim itemName As String
    Dim quantity As Double
    
    productCode = Range(codeAddress).Value
    itemName = Range(nameAddress).Value
    
    Call InitialiseVariables
    Set matchingRow = ItemExistsInOrder(productCode)
    
    If Not matchingRow Is Nothing Then
        quantity = matchingRow.Range(1, 3)
        quantity = quantity + 1
        Call UpdateQuantity(matchingRow, quantity)
    Else
        Call AddItemToOrder(productCode, itemName, 1)
    End If
End Sub
