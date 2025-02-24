Imports System
Imports System.Collections.Generic

Namespace MasterDetailExample

    Friend Class DataHelper

        Public Shared Function CreateData() As SupplierCollection
            Dim suppliers As SupplierCollection = New SupplierCollection()
            Dim supplier As Supplier = New Supplier("Exotic Liquids")
            suppliers.Add(supplier)
            supplier.Add(CreateProduct(supplier.SupplierID, "Chai"))
            supplier.Add(CreateProduct(supplier.SupplierID, "Chang"))
            supplier.Add(CreateProduct(supplier.SupplierID, "Aniseed Syrup"))
            supplier = New Supplier("New Orleans Cajun Delights")
            suppliers.Add(supplier)
            supplier.Add(CreateProduct(supplier.SupplierID, "Chef Anton's Cajun Seasoning"))
            supplier.Add(CreateProduct(supplier.SupplierID, "Chef Anton's Gumbo Mix"))
            supplier = New Supplier("Grandma Kelly's Homestead")
            suppliers.Add(supplier)
            supplier.Add(CreateProduct(supplier.SupplierID, "Grandma's Boysenberry Spread"))
            supplier.Add(CreateProduct(supplier.SupplierID, "Uncle Bob's Organic Dried Pears"))
            supplier.Add(CreateProduct(supplier.SupplierID, "Northwoods Cranberry Sauce"))
            Return suppliers
        End Function

        Private Shared random As Random = New Random(5)

        Public Shared Function CreateProduct(ByVal supplierID As Integer, ByVal productName As String) As Product
            Dim product As Product = New Product(supplierID, productName)
            product.OrderDetails.AddRange(New OrderDetail() {New OrderDetail(product.ProductID, GetRandomString(), random.Next(0, 100)), New OrderDetail(product.ProductID, GetRandomString(), random.Next(0, 100)), New OrderDetail(product.ProductID, GetRandomString(), random.Next(0, 100))})
            Return product
        End Function

        Public Shared Function CreateFakeDataSource() As List(Of Integer)
            Dim result As List(Of Integer) = New List(Of Integer)()
            result.Add(0)
            Return result
        End Function

        Public Shared Function GetRandomString() As String
            Dim path As String = IO.Path.GetRandomFileName()
            path = path.Replace(".", "")
            Return path
        End Function
    End Class
End Namespace
