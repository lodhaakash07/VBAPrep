Public Type Company
    Name As String
    Employees As Integer
    City As String
    Year As Integer
    StockMarketPresence As Boolean
End Type

Private Sub cmdSubmit_Click()
    Dim myCompany As Company
    Dim ws As Worksheet
    Dim lastRow As Long

    'Get the user input
    myCompany.Name = txtName.Value
    myCompany.Employees = CInt(txtEmployees.Value)
    myCompany.City = txtCity.Value
    myCompany.Year = CInt(txtYear.Value)
    myCompany.StockMarketPresence = chkStockMarket.Value

    'Get a reference to your worksheet
    Set ws = Worksheets("Part C")

    'Find the last row in column A of the worksheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    'Export the information
    ws.Cells(lastRow + 1, 1).Value = myCompany.Name
    ws.Cells(lastRow + 1, 2).Value = myCompany.Employees
    ws.Cells(lastRow + 1, 3).Value = myCompany.City
    ws.Cells(lastRow + 1, 4).Value = myCompany.Year
    ws.Cells(lastRow + 1, 5).Value = IIf(myCompany.StockMarketPresence, "Yes", "No")

End Sub
