using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json;
using JSON_Tools.Models;
using System.Runtime.InteropServices;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace JSON_Tools.Services.Importers
{
    public class Json2Importer : IOrderImporter
    {
        public bool CanHandle(string json)
        {
            return json.Contains("\"items\"") && !json.Contains("\"status\"");
        }

        public object Import(string json)
        {
            var root = JsonConvert.DeserializeObject<Json2Root>(json);
            return root?.Orders ?? new List<Json2Order>();
        }
    }
}
