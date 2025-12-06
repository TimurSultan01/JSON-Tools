using System;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace JSON_Tools.Models
{
    public class Json2Root
    {
        [JsonProperty("orders")]
        public List<Json2Order> Orders { get; set; }
    }

    public class Json2Order
    {
        [JsonProperty("orderId", Required = Required.Always)]
        public int OrderId { get; set; }
        [JsonProperty("customer", Required = Required.Always)]
        public string Customer { get; set; }
        [JsonProperty("created", Required = Required.Always)]
        public DateTime Created { get; set; }
        [JsonProperty("items", Required = Required.Always)]
        public List<Json2Item> Items { get; set; }

        [JsonIgnore]
        public decimal TotalAmount => Items?.Sum(i => i.ItemTotal) ?? 0;
    }

    public class Json2Item
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
}
