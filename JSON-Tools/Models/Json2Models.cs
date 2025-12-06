using System;
using System.Collections.Generic;
using System.Linq;
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
        [JsonProperty("orderId")]
        public int OrderId { get; set; }
        [JsonProperty("customer")]
        public string Customer { get; set; }
        [JsonProperty("created")]
        public DateTime Created { get; set; }
        [JsonProperty("items")]
        public List<Json2Item> Items { get; set; }

        [JsonIgnore]
        public decimal TotalAmount => Items?.Sum(i => i.ItemTotal) ?? 0;
    }

    public class Json2Item
    {
        [JsonProperty("sku")]
        public string Sku { get; set; }
        [JsonProperty("qty")]
        public int Qty { get; set; }
        [JsonProperty("price")]
        public decimal Price { get; set; }

        [JsonIgnore]
        public decimal ItemTotal => Qty * Price;
    }
}
