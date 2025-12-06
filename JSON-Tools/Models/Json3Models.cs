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
        [JsonProperty("orderId")]
        public int OrderId { get; set; }
        [JsonProperty("customer")]
        public string Customer { get; set; }
        [JsonProperty("created")]
        public DateTime Created { get; set; }
        [JsonProperty("status")]
        public string Status { get; set; }
        [JsonProperty("salesRep")]
        public string SalesRep { get; set; }
        [JsonProperty("delivery")]
        public Json3Delivery Delivery { get; set; }
        [JsonProperty("items")]
        public List<Json3Item> Items { get; set; }
    }

    public class Json3Item
    {
        [JsonProperty("sku")]
        public string Sku { get; set; }
        [JsonProperty("qty")]
        public int Qty { get; set; }
        [JsonProperty("price")]
        public decimal Price { get; set; }
    }
    public class Json3Delivery
    {
        [JsonProperty("address")]
        public string Address { get; set; }
        [JsonProperty("deliveryDate")]
        public DateTime? DeliveryDate { get; set; }
    }
}
