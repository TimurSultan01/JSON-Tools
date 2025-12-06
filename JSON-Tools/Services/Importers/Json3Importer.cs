using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json;
using JSON_Tools.Models;

namespace JSON_Tools.Services.Importers
{
    public class Json3Importer : IOrderImporter
    {
        public bool CanHandle(string json)
        {
            return json.Contains("\"status\"") && json.Contains("\"salesRep\"");
        }

        public object Import(string json)
        {
            var root = JsonConvert.DeserializeObject<Json3Root>(json);

            return root?.Orders ?? new List<Json3Order>();
        }
    }
}
