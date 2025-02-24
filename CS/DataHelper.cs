using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MasterDetailExample
{
    class DataHelper
    {

        public static  SupplierCollection CreateData()
        {
            SupplierCollection suppliers = new SupplierCollection();

            Supplier supplier = new Supplier("Exotic Liquids");
            suppliers.Add(supplier);
            supplier.Add(CreateProduct(supplier.SupplierID, "Chai"));
            supplier.Add(CreateProduct(supplier.SupplierID, "Chang"));
            supplier.Add(CreateProduct(supplier.SupplierID, "Aniseed Syrup"));

            supplier = new Supplier("New Orleans Cajun Delights");
            suppliers.Add(supplier);
            supplier.Add(CreateProduct(supplier.SupplierID, "Chef Anton's Cajun Seasoning"));
            supplier.Add(CreateProduct(supplier.SupplierID, "Chef Anton's Gumbo Mix"));

            supplier = new Supplier("Grandma Kelly's Homestead");
            suppliers.Add(supplier);
            supplier.Add(CreateProduct(supplier.SupplierID, "Grandma's Boysenberry Spread"));
            supplier.Add(CreateProduct(supplier.SupplierID, "Uncle Bob's Organic Dried Pears"));
            supplier.Add(CreateProduct(supplier.SupplierID, "Northwoods Cranberry Sauce"));

            return suppliers;
        }

        static Random random = new Random(5);

        public static  Product CreateProduct(int supplierID, string productName)
        {
            Product product = new Product(supplierID, productName);

            product.OrderDetails.AddRange(new OrderDetail[] { 
                new OrderDetail(product.ProductID, GetRandomString(), random.Next(0, 100)), 
                new OrderDetail(product.ProductID, GetRandomString(), random.Next(0, 100)),
                new OrderDetail(product.ProductID, GetRandomString(), random.Next(0, 100)) });

            return product;
        }

        public static List<int> CreateFakeDataSource()
        {
            List<int> result = new List<int>();
            result.Add(0);
            return result;
        }

        public static string GetRandomString()
        {
            string path = System.IO.Path.GetRandomFileName();
            path = path.Replace(".", ""); 
            return path;
        }
    }
}
