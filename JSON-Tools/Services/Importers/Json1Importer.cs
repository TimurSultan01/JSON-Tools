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
            var settings = new JsonSerializerSettings
            {
                MissingMemberHandling = MissingMemberHandling.Error
            };

            try
            {
                var root = JsonConvert.DeserializeObject<Json1Root>(json, settings);
                return root?.Orders ?? new List<Json1Order>();
            }
            catch (JsonSerializationException ex)
            {
                throw new InvalidOperationException(TranslateError(ex));
            }
        }

        private string TranslateError(JsonSerializationException ex)
        {
            if (ex.Message.Contains("Could not find member"))
            {
                var field = ex.Message.Split('\'')[1];
                return $"Das Feld '{field}' ist in diesem Format nicht erlaubt.";
            }
            if (ex.Message.Contains("Required property"))
            {
                var field = ex.Message.Split('\'')[1];
                return $"Das Pflichtfeld '{field}' fehlt komplett.";
            }
            if (ex.Message.Contains("Error converting value"))
            {
                return $"Ein Feld hat den falschen Datentyp (z.B. Text statt Zahl).\nDetails: {ex.Path}";
            }

            return $"Strukturfehler: {ex.Message}";
        }
    }
}