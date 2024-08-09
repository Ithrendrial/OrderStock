Attribute VB_Name = "Module3"
Public tblStock As ListObject
Public tblOrder As ListObject
    
Sub InitialiseVariables()
    Set tblStock = ActiveSheet.ListObjects("Stock")
    Set tblOrder = ActiveSheet.ListObjects("Order")
End Sub


Sub GenerateOrder()
    Dim productCode As String
    Dim itemName As String
    Dim unitsPerItem As Double
    Dim goalStock As Double
    Dim actualStock As Double
    Dim minValue As Double
    Dim orderQuantity As Double
    Dim matchingOrderRow As ListRow
    
    Call InitialiseVariables
    Dim stockRow As ListRow

    For Each stockRow In tblStock.ListRows
        If Not IsEmpty(stockRow.Range(1, 4).Value) Then
            'Retrieve item information
            productCode = stockRow.Range(1, 1).Value
            itemName = stockRow.Range(1, 2).Value
            unitsPerItem = stockRow.Range(1, 3).Value
            goalStock = stockRow.Range(1, 4).Value
            actualStock = stockRow.Range(1, 5).Value + stockRow.Range(1, 6).Value 'Cabinet + backup stock
            minValue = stockRow.Range(1, 7).Value
            
            'Calculate quantity required for order
            orderQuantity = (goalStock - actualStock) / unitsPerItem
            orderQuantity = WorksheetFunction.RoundUp(orderQuantity, 0)
        End If
        
        'Check if item exists in the order already
        Set matchingOrderRow = ItemExistsInOrder(productCode)
    
    'Add, remove, or update items as required
        If Not matchingOrderRow Is Nothing Then
            If actualStock > minValue Then
                Call RemoveItemFromOrder(matchingOrderRow) 'Remove the item from the order if it no longer needs to be restocked
            Else
                Call UpdateQuantity(matchingOrderRow, orderQuantity) 'Update the quantity if the stocking has decreased since last order update
            End If
        Else
            If actualStock <= minValue Then
                Call AddItemToOrder(productCode, itemName, orderQuantity) 'Add item to order if it is not already included
            End If
        End If
    Next stockRow
End Sub

'Check if an item in the generated order exists with a matching product code
Function ItemExistsInOrder(productCode As String) As ListRow
    Dim matchingOrderRow As ListRow
    Set matchingOrderRow = Nothing

    For Each orderRow In tblOrder.ListRows
        If orderRow.Range(1, 1).Value = productCode Then
            Set matchingOrderRow = orderRow
            Exit For
        End If
    Next orderRow
    
    Set ItemExistsInOrder = matchingOrderRow
End Function

'Add a new item to the order with the parameters as the column values
Sub AddItemToOrder(productCode As String, itemName As String, quantity As Double)
    Dim newOrderRow As ListRow
    Set newOrderRow = tblOrder.ListRows.Add(AlwaysInsert:=True)
    
    newOrderRow.Range(1, 1).Value = productCode
    newOrderRow.Range(1, 2).Value = itemName
    newOrderRow.Range(1, 3).Value = quantity
End Sub

'Update the quanity to order of an existing item in the order
Sub UpdateQuantity(rowIndex As ListRow, quantity As Double)
    rowIndex.Range(1, 3).Value = quantity
End Sub

'Remove an item in the order
Sub RemoveItemFromOrder(orderRow As ListRow)
     If Not orderRow Is Nothing Then
        orderRow.Delete
    End If
End Sub
