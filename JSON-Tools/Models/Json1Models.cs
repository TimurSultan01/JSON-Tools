using System;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace JSON_Tools.Models
{
    public class Json1Root
    {
        public List<Json1Order> Orders { get; set; }
    }

    public class Json1Order
    {
        [JsonProperty("OrderId")]
        public int OrderId { get; set; }
        [JsonProperty("customer")]
        public string Customer { get; set; }
        [JsonProperty("created")]
        public DateTime Created { get; set; }
        [JsonProperty("amount")]
        public decimal Amount { get; set; }
    }
}
