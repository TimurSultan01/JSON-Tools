using Xunit;
using System;
using System.Collections.Generic;
using Newtonsoft.Json;
using JSON_Tools.Services.Importers;
using JSON_Tools.Services;
using JSON_Tools.Models;

namespace JSON_Tools.Tests
{
    public class ImportTests
    {
        // ==========================================
        // TESTS FOR Invalid JSON
        // ==========================================
        [Fact]
        public void Importer_InvalidJson_ShouldThrowException()
        {
            // Arrange: Malformed JSON
            string json = @"
            {
                ""orders"": [
                    { ""orderId"": 1, ""customer"": ""Kunde"", ""amount"": 100.0, ""created"": ""2023-01-01"" 
                ]
            "; // Missing closing braces
            var importer = new OrderImportDispatcher();
            // Act & Assert
            var ex = Assert.Throws<Newtonsoft.Json.JsonReaderException>(() => importer.Import(json));
        }

        [Fact]
        public void Importer_NullJson_ShouldThrowException()
        {
            // Arrange: Null JSON
            string json = null;
            var importer = new OrderImportDispatcher();
            // Act & Assert
            var ex = Assert.Throws<NullReferenceException>(() => importer.Import(json));
        }

        [Fact]
        public void Importer_UnknownFormat_ShouldThrowException()
        {
            // Arrange: JSON that doesn't conform to any known format
            string json = @"
            {
                ""orders"": [
                    { ""orderId"": 1, ""kunde"": ""Test Kunde"", ""price"": 100.50,},
                ]
            }";
            var importer = new OrderImportDispatcher();
            // Act & Assert
            var ex = Assert.Throws<NotSupportedException>(() => importer.Import(json));
            Assert.Contains("Unbekanntes JSON-Format", ex.Message);
        }

        // ==========================================
        // TESTS FOR FORMAT 1
        // ==========================================
        [Fact]
        public void Json1_ValidJson_ShouldImportCorrectly()
        {
            // Arrange
            string json = @"
            {
                ""orders"": [
                    { ""orderId"": 1, ""customer"": ""Test Kunde"", ""amount"": 100.50,  ""created"": ""2023-01-01"" },
                ]
            }";
            var importer = new Json1Importer();

            // Act
            var result = (List<Json1Order>)importer.Import(json);

            // Assert
            Assert.NotNull(result);
            Assert.Single(result);
            Assert.Equal("Test Kunde", result[0].Customer);
            Assert.Equal(100.50m, result[0].Amount);
        }

        [Fact]
        public void Json1_WrongDataType_ShouldThrowException()
        {
            // Arrange: 'amount' is text instead of number
            string json = @"
            {
                ""orders"": [
                    { ""orderId"": 1, ""customer"": ""Kunde"", ""amount"": """", ""created"": ""2023-01-01"" }
                ]
            }";
            var importer = new Json1Importer();

            // Act & Assert
            // We expect an InvalidOperationException
            var ex = Assert.Throws<InvalidOperationException>(() => importer.Import(json));
            // The error message should indicate a type mismatch'
            Assert.Contains("Datentyp", ex.Message);
        }

        [Fact]
        public void Json1_ExtraField_ShouldThrowException()
        {
            // Arrange: Adding an unexpected field 'UnbekanntesFeld'
            string json = @"
            {
                ""orders"": [
                    { 
                        ""orderId"": 1, 
                        ""customer"": ""Kunde"", 
                        ""amount"": 10.0, 
                        ""created"": ""2023-01-01"",
                        ""UnbekanntesFeld"": ""Hack"" 
                    }
                ]
            }";
            var importer = new Json1Importer();

            // Act & Assert
            var ex = Assert.Throws<InvalidOperationException>(() => importer.Import(json));
            // The error message should indicate that extra fields are not allowed
            Assert.Contains("nicht erlaubt", ex.Message);
        }


        // ==========================================
        // TESTS FOR FORMAT 2
        // ==========================================
        [Fact]
        public void Json2_ValidJson_ShouldImportItems()
        {
            // Arrange
            string json = @"
            {
                ""orders"": [
                    { 
                        ""orderId"": 200, 
                        ""customer"": ""Kunde 2"", 
                        ""created"": ""2023-05-01"",
                        ""items"": [
                            { ""sku"": ""A100"", ""qty"": 5, ""price"": 10.0 },
                            { ""sku"": ""B200"", ""qty"": 2, ""price"": 20.0 }
                        ]
                    }
                ]
            }";
            var importer = new Json2Importer();

            // Act
            var result = (List<Json2Order>)importer.Import(json);

            // Assert
            Assert.NotNull(result);
            Assert.Single(result);
            Assert.Equal(2, result[0].Items.Count); // Load items correctly
            Assert.Equal("A100", result[0].Items[0].Sku);
        }

        [Fact]
        public void Json2_MissingRequiredList_ShouldThrow()
        {
            // Arrange: items missing
            string json = @"
            {
                ""orders"": [
                    { ""orderId"": 200, ""customer"": ""Kunde 2"", ""created"": ""2023-05-01"" }
                ]
            }";
            var importer = new Json2Importer();

            // Act & Assert
            var ex = Assert.Throws<InvalidOperationException>(() => importer.Import(json));
            // should mention missing 'items'
            Assert.Contains("items", ex.Message.ToLower());
        }

        [Fact]
        public void Json2_QtyAsFloat_ShouldThrow_IfIntRequired()
        {
            // Arrange: Qty is a float instead of int
            string json = @"
            {
                ""orders"": [
                    { 
                        ""orderId"": 1, ""customer"": ""Kunde"", ""created"": ""2023-01-01"",
                        ""items"": [ { ""sku"": ""X"", ""qty"": 5.5, ""price"": 10.0 } ]
                    }
                ]
            }";
            var importer = new Json2Importer();

            // Act & Assert
            var ex = Assert.Throws<JsonReaderException>(() => importer.Import(json));
        }


        // ==========================================
        // TESTS FOR FORMAT 3
        // ==========================================
        [Fact]
        public void Json3_ValidComplexStructure_ShouldImport()
        {
            // Arrange
            string json = @"
            {
                ""orders"": [
                    { 
                        ""orderId"": 300, 
                        ""customer"": ""Firma AG"", 
                        ""created"": ""2023-10-01"",
                        ""status"": ""shipped"",
                        ""salesRep"": ""Max"",
                        ""delivery"": {
                            ""address"": ""Musterstraße 1"",
                            ""deliveryDate"": ""2023-10-05""
                        },
                        ""items"": []
                    }
                ]
            }";
            var importer = new Json3Importer();

            // Act
            var result = (List<Json3Order>)importer.Import(json);

            // Assert
            Assert.NotNull(result);
            Assert.Single(result);
            Assert.Equal("Musterstraße 1", result[0].Delivery.Address);
            Assert.Equal("shipped", result[0].Status);
        }

        [Fact]
        public void Json3_MissingNestedObject_ShouldThrow()
        {
            // Arrange: 'delivery' object is missing
            string json = @"
            {
                ""orders"": [
                    { 
                        ""orderId"": 300, 
                        ""customer"": ""Firma AG"", 
                        ""status"": ""shipped"",
                        ""salesRep"": ""Max"",
                        ""created"": ""2023-10-01"",
                        ""items"": []
                    }
                ]
            }";
            var importer = new Json3Importer();

            // Act & Assert
            var ex = Assert.Throws<InvalidOperationException>(() => importer.Import(json));
            // The error message should indicate that 'delivery' is missing
            Assert.Contains("delivery", ex.Message.ToLower());
        }

        [Fact]
        public void Json3_InvalidStructure_ShouldFail()
        {
            // Arrange: 'delivery' is a string instead of an object
            string json = @"
            {
                ""orders"": [
                    { 
                        ""orderId"": 300, 
                        ""customer"": ""Firma AG"", 
                        ""delivery"": ""Ich bin kein Objekt"", 
                        ""status"": ""new"", ""salesRep"": ""x"", ""created"": ""2023-01-01"", ""items"": []
                    }
                ]
            }";
            var importer = new Json3Importer();

            // Act & Assert
            var ex = Assert.Throws<InvalidOperationException>(() => importer.Import(json));
            // The error message should indicate an issue with 'delivery'
            Assert.Contains("delivery", ex.Message);
        }
    }
}
