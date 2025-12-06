using System;
using System.Collections.Generic;
using System.Linq;
using JSON_Tools.Models;

namespace JSON_Tools.Services.Importers
{
    public class OrderImportDispatcher
    {
        private readonly List<IOrderImporter> _importers;
        public OrderImportDispatcher()
        {
            _importers = new List<IOrderImporter>
            {
                new Json3Importer(),
                new Json2Importer(),
                new Json1Importer()
            };
        }
        public object Import(string json)
        {
            var importer = _importers.FirstOrDefault(i => i.CanHandle(json));
            if (importer == null)
                throw new NotSupportedException("Unbekanntes JSON-Format.");
            
            return importer.Import(json);
        }
    }
}
