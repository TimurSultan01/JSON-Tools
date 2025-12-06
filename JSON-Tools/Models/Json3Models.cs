using System;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace JSON_Tools.Models
{
    public class Json3Root
    {
        [JsonProperty("orders")]
        public List<Json3Order> Orders { get; set; }
    }

    public class Json3Order
    {
        [JsonProperty("orderId", Required = Required.Always)]
        public int OrderId { get; set; }
        [JsonProperty("customer", Required = Required.Always)]
        public string Customer { get; set; }
        [JsonProperty("created", Required = Required.Always)]
        public DateTime Created { get; set; }
        [JsonProperty("status", Required = Required.Always)]
        public string Status { get; set; }
        [JsonProperty("salesRep", Required = Required.Always)] 
        public string SalesRep { get; set; }
        [JsonProperty("delivery", Required = Required.Always)]
        public Json3Delivery Delivery { get; set; }
        [JsonProperty("items", Required = Required.Always)]
        public List<Json3Item> Items { get; set; }

        [JsonIgnore]
        public decimal TotalAmount => Items?.Sum(i => i.ItemTotal) ?? 0;
    }

    public class Json3Item
    {
        [JsonProperty("sku", Required = Required.Always)]
        public string Sku { get; set; }
        [JsonProperty("qty", Required = Required.Always)]
        public int Qty { get; set; }
        [JsonProperty("price", Required = Required.Always)]
        public decimal Price { get; set; }

        [JsonIgnore]
        public decimal ItemTotal => Qty * Price;
    }
    public class Json3Delivery
    {
        [JsonProperty("address", Required = Required.Always)]
        public string Address { get; set; }
        [JsonProperty("deliveryDate")]
        public DateTime? DeliveryDate { get; set; }
    }
}
