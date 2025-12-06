using System.Collections.Generic;
using JSON_Tools.Models;

namespace JSON_Tools.Services.Importers
{
    public interface IOrderImporter
    {
        bool CanHandle(string json);
        object Import(string json);
    }
}
