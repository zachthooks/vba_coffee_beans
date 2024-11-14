' Calculate Total Sales per Order
Sub CalculateTotalSales()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("orders")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    For i = 2 To lastRow
        ' Assuming product prices are in column B of "products" sheet and linked by Product ID
        ws.Cells(i, "G").Value = ws.Cells(i, "E").Value * _
            Application.VLookup(ws.Cells(i, "D").Value, ThisWorkbook.Sheets("products").Range("A:B"), 2, False)
    Next i
End Sub

' Highlight Large Orders
Sub HighlightLargeOrders()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("orders")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, "E").Value > 10 Then
            ws.Rows(i).Interior.Color = RGB(255, 0, 0) ' Highlights in red
        End If
    Next i
End Sub

'Filter and Copy Orders by Date
Sub FilterOrdersByDate()
    Dim ws As Worksheet, newWs As Worksheet
    Set ws = ThisWorkbook.Sheets("orders")
    
    ' Add a new sheet for filtered data
    Set newWs = ThisWorkbook.Sheets.Add
    newWs.Name = "Filtered Orders"
    
    ' Copy header row
    ws.Rows(1).Copy newWs.Rows(1)
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ' Apply filter for dates within the specified range
    ws.Range("A1:E" & lastRow).AutoFilter Field:=2, Criteria1:=">=2021-01-01", Operator:=xlAnd, Criteria2:="<=2021-12-31"
    ws.Range("A2:E" & lastRow).SpecialCells(xlCellTypeVisible).Copy newWs.Range("A2")
    
    ' Turn off filter
    ws.AutoFilterMode = False
End Sub

' Generate Customer Summary
Sub CustomerSummary()
    Dim ws As Worksheet, summaryWs As Worksheet
    Set ws = ThisWorkbook.Sheets("orders")
    
    ' Create a new sheet for the summary
    Set summaryWs = ThisWorkbook.Sheets.Add
    summaryWs.Name = "Customer Summary"
    
    ' Set headers for the summary
    summaryWs.Cells(1, 1).Value = "Customer ID"
    summaryWs.Cells(1, 2).Value = "Total Quantity"
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = 2 To lastRow
        Dim customerID As String, quantity As Integer
        customerID = ws.Cells(i, "C").Value
        quantity = ws.Cells(i, "E").Value
        
        ' Update quantity per customer
        If dict.exists(customerID) Then
            dict(customerID) = dict(customerID) + quantity
        Else
            dict.Add customerID, quantity
        End If
    Next i
    
    ' Output summary data to the new sheet
    Dim row As Integer
    row = 2
    Dim Key As Variant
    For Each Key In dict.Keys
        summaryWs.Cells(row, 1).Value = Key
        summaryWs.Cells(row, 2).Value = dict(Key)
        row = row + 1
    Next Key
End Sub

' Automated Sales Report Generation
Sub GenerateSalesReport()
    Dim ws As Worksheet, reportWs As Worksheet
    Set ws = ThisWorkbook.Sheets("orders")
    
    ' Create a new sheet for the sales report
    Set reportWs = ThisWorkbook.Sheets.Add
    reportWs.Name = "Sales Report"
    
    ' Set up report headers
    reportWs.Cells(1, 1).Value = "Product ID"
    reportWs.Cells(1, 2).Value = "Total Quantity Sold"
    reportWs.Cells(1, 3).Value = "Total Sales Amount"
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Calculate quantity sold per product
    Dim i As Long
    For i = 2 To lastRow
        Dim productID As String, quantity As Double, salesAmount As Double
        productID = ws.Cells(i, "D").Value
        quantity = ws.Cells(i, "E").Value
        salesAmount = quantity * Application.VLookup(productID, ThisWorkbook.Sheets("products").Range("A:B"), 2, False)
        
        ' Update or add to dictionary
        If dict.exists(productID) Then
            dict(productID)(0) = dict(productID)(0) + quantity
            dict(productID)(1) = dict(productID)(1) + salesAmount
        Else
            dict.Add productID, Array(quantity, salesAmount)
        End If
    Next i
    
    ' Write report data
    Dim row As Integer
    row = 2
    Dim Key As Variant
    For Each Key In dict.Keys
        reportWs.Cells(row, 1).Value = Key
        reportWs.Cells(row, 2).Value = dict(Key)(0)
        reportWs.Cells(row, 3).Value = dict(Key)(1)
        row = row + 1
    Next Key
End Sub

' Price Update Automation
Sub UpdateProductPrices()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("products")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim newPrice As Double
    Dim productID As String
    
    ' Prompt user for Product ID and new price
    productID = InputBox("Enter Product ID to update price:")
    newPrice = InputBox("Enter new price for " & productID & ":")
    
    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, "A").Value = productID Then
            ws.Cells(i, "B").Value = newPrice
            MsgBox "Price updated successfully for " & productID
            Exit Sub
        End If
    Next i
    
    MsgBox "Product ID not found!"
End Sub

' Customer Loyalty Program Automation
Sub GenerateLoyalCustomerList()
    Dim ws As Worksheet, loyaltyWs As Worksheet
    Set ws = ThisWorkbook.Sheets("orders")
    
    ' Create a new sheet for loyal customers
    Set loyaltyWs = ThisWorkbook.Sheets.Add
    loyaltyWs.Name = "Loyal Customers"
    
    ' Set up headers
    loyaltyWs.Cells(1, 1).Value = "Customer ID"
    loyaltyWs.Cells(1, 2).Value = "Total Quantity Ordered"
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Calculate total quantity per customer
    Dim i As Long
    For i = 2 To lastRow
        Dim customerID As String, quantity As Double
        customerID = ws.Cells(i, "C").Value
        quantity = ws.Cells(i, "E").Value
        
        If dict.exists(customerID) Then
            dict(customerID) = dict(customerID) + quantity
        Else
            dict.Add customerID, quantity
        End If
    Next i
    
    ' Output customers with quantities above a certain threshold
    Dim row As Integer
    row = 2
    For Each Key In dict.Keys
        If dict(Key) > 50 Then ' Define loyalty threshold
            loyaltyWs.Cells(row, 1).Value = Key
            loyaltyWs.Cells(row, 2).Value = dict(Key)
            row = row + 1
        End If
    Next Key
End Sub

' Inventory Tracking and Low Stock Alerts
Sub CheckLowStock()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("products")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim lowStockThreshold As Integer
    lowStockThreshold = 10 ' Set low stock threshold
    
    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, "C").Value < lowStockThreshold Then
            MsgBox "Low stock alert for Product ID " & ws.Cells(i, "A").Value & ": Only " & ws.Cells(i, "C").Value & " units left."
        End If
    Next i
End Sub

' Monthly Sales Summary by Country
Sub MonthlySalesSummaryByCountry()
    Dim ws As Worksheet, custWs As Worksheet, summaryWs As Worksheet
    Set ws = ThisWorkbook.Sheets("orders")
    Set custWs = ThisWorkbook.Sheets("customers")
    
    ' Create a new sheet for the summary
    Set summaryWs = ThisWorkbook.Sheets.Add
    summaryWs.Name = "Monthly Sales by Country"
    
    ' Set headers for the summary
    summaryWs.Cells(1, 1).Value = "Month"
    summaryWs.Cells(1, 2).Value = "Country"
    summaryWs.Cells(1, 3).Value = "Total Sales"
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Calculate monthly sales per country
    Dim i As Long
    For i = 2 To lastRow
        Dim orderDate As Date, month As String, customerID As String, country As String, salesAmount As Double
        orderDate = ws.Cells(i, "B").Value
        month = Format(orderDate, "yyyy-mm")
        customerID = ws.Cells(i, "C").Value
        country = Application.VLookup(customerID, custWs.Range("A:B"), 2, False)
        salesAmount = ws.Cells(i, "E").Value * Application.VLookup(ws.Cells(i, "D").Value, ThisWorkbook.Sheets("products").Range("A:B"), 2, False)
        
        Dim key As String
        key = month & "-" & country
        
        ' Update or add to dictionary
        If dict.exists(key) Then
            dict(key) = dict(key) + salesAmount
        Else
            dict.Add key, salesAmount
        End If
    Next i
    
    ' Output summary data
    Dim row As Integer
    row = 2
    Dim Key As Variant
    For Each Key In dict.Keys
        summaryWs.Cells(row, 1).Value = Split(Key, "-")(0)
        summaryWs.Cells(row, 2).Value = Split(Key, "-")(1)
        summaryWs.Cells(row, 3).Value = dict(Key)
        row = row + 1
    Next Key
End Sub
