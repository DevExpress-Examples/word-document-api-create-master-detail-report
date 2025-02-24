Imports System.Collections.Generic
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native

Namespace MasterDetailExample

    Friend Class MergeProcessor

        Private mainRichEdit As RichEditDocumentServer = New RichEditDocumentServer()

        Private suppllierRichEdit As RichEditDocumentServer = New RichEditDocumentServer()

        Private productRichEdit As RichEditDocumentServer = New RichEditDocumentServer()

        Private ordersRichEdit As RichEditDocumentServer = New RichEditDocumentServer()

        Private resultRichEdit As RichEditDocumentServer = New RichEditDocumentServer()

        Private ds As SupplierCollection

        Private dataDetailedForProducts As ProductCollection

        Private dataDetailedForOrders As OrderDetailCollection

        Private Shared fakeDataSource As List(Of Integer) = DataHelper.CreateFakeDataSource()

        Private supplierID As Integer = -1

        Private productID As Integer = -1

        Public Sub New()
            ' Subscribe to the CalculateDocumentVariable event that triggers the master-detail report generation
            AddHandler resultRichEdit.CalculateDocumentVariable, New CalculateDocumentVariableEventHandler(AddressOf resultRichEdit_CalculateDocumentVariable)
            ' Load main template
            mainRichEdit.LoadDocument("main.rtf")
            ' Create project's data source
            ds = DataHelper.CreateData()
            ' Load templates and specify data sources for RichEdit servers.
            suppllierRichEdit.LoadDocument("supplier.rtf")
            suppllierRichEdit.Options.MailMerge.DataSource = ds
            productRichEdit.LoadDocument("detail.rtf")
            productRichEdit.Options.MailMerge.DataSource = ds
            productRichEdit.Options.MailMerge.DataMember = "Products"
            ordersRichEdit.LoadDocument("detaildetail.rtf")
            ordersRichEdit.Options.MailMerge.DataSource = ds
            ordersRichEdit.Options.MailMerge.DataMember = "Products.OrderDetails"
        End Sub

'#Region "#startmailmerge"
        ' Start the process by merging main template into the document contained within the resultRichEdit server.
        Public Sub Start()
            ' Since the main template contains no merge fields requiring no merge data, provide a fake data source.
            ' Otherwise mailmerge will not start.
            mainRichEdit.Options.MailMerge.DataSource = fakeDataSource
            ' Trigger the multistage process. After the first mailmerge the CalculateDocumentVariable event
            'for the resultRichEdit server fires.
            mainRichEdit.MailMerge(resultRichEdit.Document)
            resultRichEdit.SaveDocument("result.docx", DocumentFormat.OpenXml)
        End Sub

'#End Region  ' #startmailmerge
'#Region "#secondstage"
        ' Second stage. For each Supplier ID create a detailed document that will be inserted in place of the DOCVARIABLE field.
        Private Sub resultRichEdit_CalculateDocumentVariable(ByVal sender As Object, ByVal e As CalculateDocumentVariableEventArgs)
            If Equals(e.VariableName, "Supplier") Then
                ' Create a text engine to process a document after the mail merge.
                Dim richServerMaster As RichEditDocumentServer = New RichEditDocumentServer()
                ' Provide a procedure for further processing
                AddHandler richServerMaster.CalculateDocumentVariable, AddressOf richServerMaster_CalculateDocumentVariable
                ' Create a merged document using the Supplier template. The document will contain DOCVARIABLE fields with ProductID arguments. 
                ' The CalculateDocumentVariable event for the richServerMaster fires.
                suppllierRichEdit.MailMerge(richServerMaster)
                RemoveHandler richServerMaster.CalculateDocumentVariable, AddressOf richServerMaster_CalculateDocumentVariable
                ' Return the document to insert.
                e.Value = richServerMaster
                ' Required to use e.Value. Otherwise it will be ignored.
                e.Handled = True
            End If
        End Sub

'#End Region  ' #secondstage
'#Region "#thirdstage"
        ' Third stage. For each Product ID create a detailed document that will be inserted in place of the DOCVARIABLE field.
        Private Sub richServerMaster_CalculateDocumentVariable(ByVal sender As Object, ByVal e As CalculateDocumentVariableEventArgs)
            Dim currentSupplierID As Integer = GetID(e.Arguments(0).Value)
            If currentSupplierID = -1 Then Return
            If supplierID <> currentSupplierID Then
                ' Get data source that contains products for the specified supplier.
                dataDetailedForProducts = GetProductsDataFilteredbySupplier(currentSupplierID)
                supplierID = currentSupplierID
            End If

            If Equals(e.VariableName, "Product") Then
                ' Create a text engine to process a document after the mail merge.
                Dim richServerDetail As RichEditDocumentServer = New RichEditDocumentServer()
                ' Specify data source for mail merge.
                Dim options As MailMergeOptions = productRichEdit.CreateMailMergeOptions()
                options.DataSource = dataDetailedForProducts
                ' Specify that the resulting table should be joined with the header table.
                ' Do not specify this option if calculated fields are not within table cells.
                options.MergeMode = MergeMode.JoinTables
                ' Provide a procedure for further processing.
                AddHandler richServerDetail.CalculateDocumentVariable, AddressOf richServerDetail_CalculateDocumentVariable
                ' Create a merged document using the Product template. The document will contain DOCVARIABLE fields with OrderID arguments. 
                ' The CalculateDocumentVariable event for the richServerDetail fires.
                productRichEdit.MailMerge(options, richServerDetail)
                RemoveHandler richServerDetail.CalculateDocumentVariable, AddressOf richServerDetail_CalculateDocumentVariable
                ' Return the document to insert.
                e.Value = richServerDetail
                ' This setting is required for inserting e.Value into the source document. Otherwise it will be ignored.
                e.Handled = True
            End If
        End Sub

'#End Region  ' #thirdstage
'#Region "#fourthstage"
        ' Fourth stage. For each Order ID create a detailed document that will be inserted in place of the DOCVARIABLE field.
        ' This is the final stage and the Product.Orders template does not contain DOCVARIABLE fields. So, further processing is not required.
        Private Sub richServerDetail_CalculateDocumentVariable(ByVal sender As Object, ByVal e As CalculateDocumentVariableEventArgs)
            Dim currentProductID As Integer = GetID(e.Arguments(0).Value)
            If currentProductID = -1 Then Return
            If productID <> currentProductID Then
                ' Get data source that contains orders for the specified product. 
                ' The data source is obtained from the data already filtered by supplier.
                dataDetailedForOrders = GetOrderDataFilteredbyProductAndSupplier(currentProductID)
                productID = currentProductID
            End If

            If Equals(e.VariableName, "OrderDetails") Then
                Dim richServerDetailDetail As RichEditDocumentServer = New RichEditDocumentServer()
                Dim options As MailMergeOptions = ordersRichEdit.CreateMailMergeOptions()
                options.DataSource = dataDetailedForOrders
                options.MergeMode = MergeMode.JoinTables
                ordersRichEdit.MailMerge(options, richServerDetailDetail)
                e.Value = richServerDetailDetail
                e.Handled = True
            End If
        End Sub

'#End Region  ' #fourthstage
'#Region "Helper Methods"
        Protected Friend Overridable Function GetProductsDataFilteredbySupplier(ByVal supplierID As Integer) As ProductCollection
            Dim products As ProductCollection = New ProductCollection()
            For Each s As Supplier In ds
                If s.SupplierID = supplierID Then
                    products.AddRange(s.Products)
                End If
            Next

            Return products
        End Function

        Protected Friend Overridable Function GetOrderDataFilteredbyProductAndSupplier(ByVal productID As Integer) As OrderDetailCollection
            Dim orders As OrderDetailCollection = New OrderDetailCollection()
            For Each p As Product In dataDetailedForProducts
                If p.ProductID = productID Then
                    orders.AddRange(p.OrderDetails)
                End If
            Next

            Return orders
        End Function

        Protected Friend Overridable Function GetID(ByVal value As String) As Integer
            Dim result As Integer
            If Integer.TryParse(value, result) Then Return result
            Return -1
        End Function
'#End Region  ' Helper Methods
    End Class
End Namespace
