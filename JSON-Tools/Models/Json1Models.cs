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
        [JsonProperty("OrderId", Required = Required.Always)]
        public int OrderId { get; set; }
        [JsonProperty("customer", Required = Required.Always)]
        public string Customer { get; set; }
        [JsonProperty("created", Required = Required.Always)]
        public DateTime Created { get; set; }
        [JsonProperty("amount", Required = Required.Always)]
        public decimal Amount { get; set; }
    }
}
