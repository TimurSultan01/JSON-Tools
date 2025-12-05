using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;

namespace JSON_Tools.Models
{
    public class RootObject
    {
        public List<Order> Orders { get; set; }
    }

    public class Order
    {
        public int OrderId { get; set; }
        public string Customer { get; set; }
        public DateTime Created { get; set; }
        public decimal Amount { get; set; }
    }
}
