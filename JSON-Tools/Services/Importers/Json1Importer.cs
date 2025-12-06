using System.Collections.Generic;
using Newtonsoft.Json;
using JSON_Tools.Models;

namespace JSON_Tools.Services.Importers
{
    public class Json1Importer : IOrderImporter
    {
        public bool CanHandle(string json)
        {
            return json.Contains("\"amount\"") && !json.Contains("\"items\"");
        }
        public object Import(string json)
        {
            var root = JsonConvert.DeserializeObject<Json1Root>(json);
            return root?.Orders ?? new List<Json1Order>();
        }
    }
}
