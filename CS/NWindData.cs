namespace MasterDetailExample
{
    public class NWindData
    {
        public List<Customer> Customers { get; set; } = new();

        public record Customer(
            string CustomerId,
            string CompanyName,
            string ContactName,
            string Country,
            string Address,
            string City,
            string Phone,
            List<Order> Orders);

        public record Order(
            string OrderID,
            DateTime? OrderDate,
            string ShipCountry,
            double Freight,
            List<OrderDetails> OrderDetails);

        public record OrderDetails(
            int ProductId,
            string ProductName,
            int Quantity,
            decimal UnitPrice,
            double Discount);
    }
}
