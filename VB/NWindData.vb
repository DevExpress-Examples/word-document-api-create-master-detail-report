Imports System
Imports System.Collections.Generic

Namespace MasterDetailExample
    Public Class NWindData
        Public Property Customers As List(Of Customer) = New List(Of Customer)()

        Public Class Customer
            Public Property CustomerId As String
            Public Property CompanyName As String
            Public Property ContactName As String
            Public Property Country As String
            Public Property Address As String
            Public Property City As String
            Public Property Phone As String
            Public Property Orders As List(Of [Order])

            Public Sub New(customerId As String, companyName As String, contactName As String, country As String, address As String, city As String, phone As String, orders As List(Of [Order]))
                Me.CustomerId = customerId
                Me.CompanyName = companyName
                Me.ContactName = contactName
                Me.Country = country
                Me.Address = address
                Me.City = city
                Me.Phone = phone
                Me.Orders = orders
            End Sub
        End Class

        Public Class [Order]
            Public Property OrderID As String
            Public Property OrderDate As DateTime?
            Public Property ShipCountry As String
            Public Property Freight As Double
            Public Property OrderDetails As List(Of OrderDetails)

            Public Sub New(orderID As String, orderDate As DateTime?, shipCountry As String, freight As Double, orderDetails As List(Of OrderDetails))
                Me.OrderID = orderID
                Me.OrderDate = orderDate
                Me.ShipCountry = shipCountry
                Me.Freight = freight
                Me.OrderDetails = orderDetails
            End Sub
        End Class

        Public Class OrderDetails
            Public Property ProductId As Integer
            Public Property ProductName As String
            Public Property Quantity As Integer
            Public Property UnitPrice As Decimal
            Public Property Discount As Double

            Public Sub New(productId As Integer, productName As String, quantity As Integer, unitPrice As Decimal, discount As Double)
                Me.ProductId = productId
                Me.ProductName = productName
                Me.Quantity = quantity
                Me.UnitPrice = unitPrice
                Me.Discount = discount
            End Sub
        End Class
    End Class
End Namespace