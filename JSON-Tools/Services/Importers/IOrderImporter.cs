using JSON_Tools.Models;
using Newtonsoft.Json;
using System.Collections.Generic;

namespace JSON_Tools.Services.Importers
{
    public interface IOrderImporter
    {
        bool CanHandle(string json);
        object Import(string json);
    }
}
