Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Collections
Imports System.ComponentModel
Imports System.Linq

Namespace MasterDetailExample

	Public Class SupplierCollection
		Inherits ArrayList
		Implements ITypedList
		Private Function GetItemProperties(ByVal listAccessors() As PropertyDescriptor) As PropertyDescriptorCollection Implements ITypedList.GetItemProperties
			If listAccessors IsNot Nothing AndAlso listAccessors.Length > 0 Then
				Dim listAccessor As PropertyDescriptor = listAccessors(listAccessors.Length - 1)
				If listAccessor.PropertyType.Equals(GetType(ProductCollection)) Then
					Return TypeDescriptor.GetProperties(GetType(Product))
				ElseIf listAccessor.PropertyType.Equals(GetType(OrderDetailCollection)) Then
					Return TypeDescriptor.GetProperties(GetType(OrderDetail))
				End If
			End If
			Return TypeDescriptor.GetProperties(GetType(Supplier))
		End Function
		Private Function GetListName(ByVal listAccessors() As PropertyDescriptor) As String Implements ITypedList.GetListName
			Return "Suppliers"
		End Function
	End Class

	Public Class Supplier
		Private Shared nextID As Integer = 0
		Private id As Integer
		Private name As String
		Private products_Renamed As New ProductCollection()

		Public ReadOnly Property Products() As ProductCollection
			Get
				Return products_Renamed
			End Get
		End Property
		Public ReadOnly Property SupplierID() As Integer
			Get
				Return id
			End Get
		End Property
		Public ReadOnly Property CompanyName() As String
			Get
				Return name
			End Get
		End Property

		Public Sub New(ByVal name As String)
			Me.name = name

			Me.id = nextID
			nextID += 1
		End Sub
		Public Sub Add(ByVal product As Product)
			products_Renamed.Add(product)
		End Sub
	End Class

	Public Class ProductCollection
		Inherits ArrayList
		Implements ITypedList
		Private Function GetItemProperties(ByVal listAccessors() As PropertyDescriptor) As PropertyDescriptorCollection Implements ITypedList.GetItemProperties
			Return TypeDescriptor.GetProperties(GetType(Product))
		End Function
		Private Function GetListName(ByVal listAccessors() As PropertyDescriptor) As String Implements ITypedList.GetListName
			Return "Products"
		End Function
	End Class

	Public Class Product
		Private Shared nextID As Integer = 0

		Private orderDetails_Renamed As New OrderDetailCollection()
		Private suppID As Integer
		Private prodID As Integer
		Private name As String

		Public ReadOnly Property SupplierID() As Integer
			Get
				Return suppID
			End Get
		End Property
		Public ReadOnly Property ProductID() As Integer
			Get
				Return prodID
			End Get
		End Property
		Public ReadOnly Property ProductName() As String
			Get
				Return name
			End Get
		End Property
		Public ReadOnly Property OrderDetails() As OrderDetailCollection
			Get
				Return orderDetails_Renamed
			End Get
		End Property

		Public Sub New(ByVal suppID As Integer, ByVal name As String)
			Me.suppID = suppID
			Me.name = name

			Me.prodID = nextID
			nextID += 1
		End Sub
	End Class

	Public Class OrderDetailCollection
		Inherits ArrayList
		Implements ITypedList
		Private Function GetItemProperties(ByVal listAccessors() As PropertyDescriptor) As PropertyDescriptorCollection Implements ITypedList.GetItemProperties
			Return TypeDescriptor.GetProperties(GetType(OrderDetail))
		End Function
		Private Function GetListName(ByVal listAccessors() As PropertyDescriptor) As String Implements ITypedList.GetListName
			Return "OrderDetails"
		End Function
	End Class

	Public Class OrderDetail
		Private prodID As Integer
		Private orderID_Renamed As String
		Private quantity_Renamed As Short
		Public ReadOnly Property ProductID() As Integer
			Get
				Return prodID
			End Get
		End Property
		Public ReadOnly Property OrderID() As String
			Get
				Return orderID_Renamed
			End Get
		End Property
		Public ReadOnly Property Quantity() As Short
			Get
				Return quantity_Renamed
			End Get
		End Property

		Public Sub New(ByVal prodID As Integer, ByVal orderID As String, ByVal quantity As Integer)
			Me.prodID = prodID
			Me.orderID_Renamed = orderID
			Me.quantity_Renamed = Convert.ToInt16(quantity)
		End Sub
	End Class
End Namespace
