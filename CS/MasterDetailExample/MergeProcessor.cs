using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;

namespace MasterDetailExample
{
    class MergeProcessor
    {

        RichEditDocumentServer mainRichEdit = new RichEditDocumentServer();
        RichEditDocumentServer suppllierRichEdit = new RichEditDocumentServer();
        RichEditDocumentServer productRichEdit = new RichEditDocumentServer();
        RichEditDocumentServer ordersRichEdit = new RichEditDocumentServer();
        RichEditDocumentServer resultRichEdit = new RichEditDocumentServer();        
        
        SupplierCollection ds;
        ProductCollection dataDetailedForProducts;
        OrderDetailCollection dataDetailedForOrders;
        static List<int> fakeDataSource = DataHelper.CreateFakeDataSource();
        int supplierID = -1;
        int productID = -1;

        public MergeProcessor()
        {
           // Subscribe to the CalculateDocumentVariable event that triggers the master-detail report generation
            resultRichEdit.CalculateDocumentVariable += new CalculateDocumentVariableEventHandler(resultRichEdit_CalculateDocumentVariable);

            // Load main template
            mainRichEdit.LoadDocument("main.rtf");

            // Create project's data source
            ds = DataHelper.CreateData();

            // Load templates and specify data sources for RichEdit servers.

            suppllierRichEdit.LoadDocument("supplier.rtf");
            suppllierRichEdit.Options.MailMerge.DataSource = ds;
            
            productRichEdit.LoadDocument("detail.rtf");
            productRichEdit.Options.MailMerge.DataSource = ds;
            productRichEdit.Options.MailMerge.DataMember = "Products";
            
            ordersRichEdit.LoadDocument("detaildetail.rtf");
            ordersRichEdit.Options.MailMerge.DataSource = ds;
            ordersRichEdit.Options.MailMerge.DataMember = "Products.OrderDetails";

        }

        #region #startmailmerge
        // Start the process by merging main template into the document contained within the resultRichEdit server.
        public void Start()
        {
            // Since the main template contains no merge fields requiring no merge data, provide a fake data source.
            // Otherwise mailmerge will not start.
            mainRichEdit.Options.MailMerge.DataSource = fakeDataSource;
            // Trigger the multistage process. After the first mailmerge the CalculateDocumentVariable event
            //for the resultRichEdit server fires.
            mainRichEdit.MailMerge(resultRichEdit.Document);
            resultRichEdit.SaveDocument("result.docx", DocumentFormat.OpenXml);
        }
        #endregion #startmailmerge

        #region #secondstage
        // Second stage. For each Supplier ID create a detailed document that will be inserted in place of the DOCVARIABLE field.
        void resultRichEdit_CalculateDocumentVariable(object sender, CalculateDocumentVariableEventArgs e)
        {
            if (e.VariableName == "Supplier") {
                // Create a text engine to process a document after the mail merge.
                RichEditDocumentServer richServerMaster = new RichEditDocumentServer();
                // Provide a procedure for further processing
                richServerMaster.CalculateDocumentVariable += richServerMaster_CalculateDocumentVariable;
                // Create a merged document using the Supplier template. The document will contain DOCVARIABLE fields with ProductID arguments. 
                // The CalculateDocumentVariable event for the richServerMaster fires.
                suppllierRichEdit.MailMerge(richServerMaster);
                richServerMaster.CalculateDocumentVariable -= richServerMaster_CalculateDocumentVariable;                
                // Return the document to insert.
                e.Value = richServerMaster;
                // Required to use e.Value. Otherwise it will be ignored.
                e.Handled = true;
            }
        }
        #endregion #secondstage
        #region #thirdstage
        // Third stage. For each Product ID create a detailed document that will be inserted in place of the DOCVARIABLE field.
        void richServerMaster_CalculateDocumentVariable(object sender, CalculateDocumentVariableEventArgs e)
        {
            int currentSupplierID = GetID(e.Arguments[0].Value);
            if (currentSupplierID == -1)
                return;

            if (supplierID != currentSupplierID) {
                // Get data source that contains products for the specified supplier.
                dataDetailedForProducts = GetProductsDataFilteredbySupplier(currentSupplierID);
                supplierID = currentSupplierID;
            }

            if (e.VariableName == "Product") {
                // Create a text engine to process a document after the mail merge.
                RichEditDocumentServer richServerDetail = new RichEditDocumentServer();
                // Specify data source for mail merge.
                MailMergeOptions options = productRichEdit.CreateMailMergeOptions();
                options.DataSource = dataDetailedForProducts;
                // Specify that the resulting table should be joined with the header table.
                // Do not specify this option if calculated fields are not within table cells.
                options.MergeMode = MergeMode.JoinTables;
                // Provide a procedure for further processing.
                richServerDetail.CalculateDocumentVariable += richServerDetail_CalculateDocumentVariable;
                // Create a merged document using the Product template. The document will contain DOCVARIABLE fields with OrderID arguments. 
                // The CalculateDocumentVariable event for the richServerDetail fires.
                productRichEdit.MailMerge(options, richServerDetail);
                richServerDetail.CalculateDocumentVariable -= richServerDetail_CalculateDocumentVariable;
                // Return the document to insert.
                e.Value = richServerDetail;
                // This setting is required for inserting e.Value into the source document. Otherwise it will be ignored.
                e.Handled = true;
            }
        }
        #endregion #thirdstage
        #region #fourthstage
        // Fourth stage. For each Order ID create a detailed document that will be inserted in place of the DOCVARIABLE field.
        // This is the final stage and the Product.Orders template does not contain DOCVARIABLE fields. So, further processing is not required.
        void richServerDetail_CalculateDocumentVariable(object sender, CalculateDocumentVariableEventArgs e)
        {
            int currentProductID = GetID(e.Arguments[0].Value);
            if (currentProductID == -1)
                return;

            if (productID != currentProductID) {
                // Get data source that contains orders for the specified product. 
                // The data source is obtained from the data already filtered by supplier.
                dataDetailedForOrders = GetOrderDataFilteredbyProductAndSupplier(currentProductID);
                productID = currentProductID;
            }
            
            if (e.VariableName == "OrderDetails") {

                RichEditDocumentServer richServerDetailDetail = new RichEditDocumentServer();
                MailMergeOptions options = ordersRichEdit.CreateMailMergeOptions();
                options.DataSource = dataDetailedForOrders;
                options.MergeMode = MergeMode.JoinTables;
                ordersRichEdit.MailMerge(options, richServerDetailDetail);
                e.Value = richServerDetailDetail;
                e.Handled = true;
            }
        }
        #endregion #fourthstage
        #region Helper Methods
        protected internal virtual ProductCollection GetProductsDataFilteredbySupplier(int supplierID)
        {
            ProductCollection products = new ProductCollection();

            foreach (Supplier s in ds) {
                if (s.SupplierID == supplierID) {
                    products.AddRange(s.Products);
                }
            }

            return products;
        }

        protected internal virtual OrderDetailCollection GetOrderDataFilteredbyProductAndSupplier(int productID)
        {
            OrderDetailCollection orders = new OrderDetailCollection();

            foreach (Product p in dataDetailedForProducts) {
                if (p.ProductID == productID) {
                    orders.AddRange(p.OrderDetails);
                }
            }
            return orders;
        }
        protected internal virtual int GetID(string value)
        {
            int result;
            if (Int32.TryParse(value, out result))
                return result;
            return -1;
        }
        #endregion Helper Methods
    

    }
}
