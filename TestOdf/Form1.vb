Public Class Form1

    Private Sub ButtonCreateOdf_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonCreateOdf.Click

        Dim ds As New DataSet
        Dim sql As String = "SELECT RegionID, RegionDescription FROM Regions;"
        sql &= "SELECT EmployeeID, LastName, FirstName, Title, TitleOfCourtesy, BirthDate, HireDate, Address, City, Region, PostalCode, Country, HomePhone, Extension, Photo, Notes, PhotoPath FROM PreviousEmployees;"
        sql &= "SELECT EmployeeID, LastName, FirstName, Title, TitleOfCourtesy, BirthDate, HireDate, Address, City, Region, PostalCode, Country, HomePhone, Extension, Photo, Notes, PhotoPath FROM Employees;"
        sql &= "SELECT CustomerID, CompanyName, ContactName, ContactTitle, Address, City, Region, PostalCode, Country, Phone, Fax FROM Customers;"
        sql &= "SELECT SupplierID, CompanyName, ContactName, ContactTitle, Address, City, Region, PostalCode, Country, Phone, Fax, HomePage FROM Suppliers;"
        sql &= "SELECT OrderID, CustomsDescription, ExciseTax FROM InternationalOrders;"
        sql &= "SELECT OrderID, ProductID, UnitPrice, Quantity, Discount FROM OrderDetails;"
        sql &= "SELECT TerritoryID, TerritoryDescription, RegionID FROM Territories;"
        sql &= "SELECT EmployeeID, TerritoryID FROM EmployeesTerritories;"
        sql &= "SELECT ProductID, ProductName, SupplierID, CategoryID, QuantityPerUnit, UnitPrice, UnitsInStock, UnitsOnOrder, ReorderLevel, Discontinued, DiscontinuedDate FROM Products;"
        sql &= "SELECT OrderID, CustomerID, EmployeeID, OrderDate, RequiredDate, ShippedDate, Freight, ShipName, ShipAddress, ShipCity, ShipRegion, ShipPostalCode, ShipCountry FROM Orders;"
        sql &= "SELECT CategoryID, CategoryName, Description, Picture FROM Categories;"

        Dim cmd As New System.Data.SQLite.SQLiteCommand
        Dim cn As New System.Data.SQLite.SQLiteConnection("Data Source=northwindEF.db;Version=3;")
        cmd.Connection = cn
        cmd.CommandType = CommandType.Text
        cmd.CommandText = sql

        Dim da As New System.Data.SQLite.SQLiteDataAdapter
        da.SelectCommand = cmd
        da.Fill(ds)

        ds.Tables(0).TableName = "Regions"
        ds.Tables(1).TableName = "PreviousEmployees"
        ds.Tables(2).TableName = "Employees"
        ds.Tables(3).TableName = "Customers"
        ds.Tables(4).TableName = "Suppliers"
        ds.Tables(5).TableName = "InternationalOrders"
        ds.Tables(6).TableName = "OrderDetails"
        ds.Tables(7).TableName = "Territories"
        ds.Tables(8).TableName = "EmployeesTerritories"
        ds.Tables(9).TableName = "Products"
        ds.Tables(10).TableName = "Orders"
        ds.Tables(11).TableName = "Categories"

        Dim thePath As String = "C:\Users\Peter\Desktop\testtheodf.odf"
        Dim theSpread As New OdsReaderWriter
        theSpread.WriteOdsFile(ds, thePath)
    End Sub
End Class
