using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;
using JSON_Tools.Models;
using JSON_Tools.Services.Importers;


namespace JSON_Tools.Services
{
    public class JsonImportService
    {
        private readonly OrderImportDispatcher _dispatcher = new OrderImportDispatcher();

        public object LoadOrders(string filePath)
        {
           var json = File.ReadAllText(filePath);
            return _dispatcher.Import(json);
        }
    }
}
