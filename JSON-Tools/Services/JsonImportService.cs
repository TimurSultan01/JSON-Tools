using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;
using JSON_Tools.Models;


namespace JSON_Tools.Services
{
    public class JsonImportService
    {
        public List<Order> LoadOrders(string filePath)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException("Die angegebene Datei wurde nicht gefunden.", filePath);
            
            try
            {
                string jsonContent = File.ReadAllText(filePath);

                var rootObject = JsonConvert.DeserializeObject<RootObject>(jsonContent);
                
                if (rootObject?.Orders == null)
                    throw new InvalidDataException("Die JSON-Daten enthalten keine gültigen Bestellungen.");

                return rootObject.Orders;
            }
            catch (JsonException ex)
            {
                throw new InvalidDataException("Fehler beim Laden der JSON-Daten.", ex);
            }
        }
    }
}
