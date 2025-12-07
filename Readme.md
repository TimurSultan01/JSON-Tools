# JSON Order Manager Tool

Eine robuste WPF-Anwendung zum Importieren, Validieren und Konvertieren verschiedener JSON-Bestellformate in Excel. Die Anwendung wurde entwickelt, um strikte Datenintegrität sicherzustellen und benutzerfreundliche Fehlermeldungen bei Format- oder Logikfehlern bereitzustellen.

## Erste Schritte & Nutzung

__Installation (Für Entwickler)__
1. Repository klonen.
2. Projekt in Visual Studio öffnen.
3. NuGet-Pakete wiederherstellen.

__Sofortiger Start (Für Endbenutzer)__
1. Wechseln Sie in den Ordner Builds/v1.0.0.
2. Starten Sie die App direkt über JSON_Tools.exe.

## Technologien & Abhängigkeiten
- __Framework__: .NET 10.0
- __JSON Parsing__: Newtonsoft.Json (v13.0.x)
- __MVVM__: CommunityToolkit.Mvvm (v8.4.x)
- __Office Export__: ClosedXML (v0.105.x
- __Office Export__: EPPlus (5.8.x)

## Architektur 
Das Projekt folgt dem __Model-View-ViewModel (MVVM)__ Entwurfsmuster, um eine klare Trennung zwischen der Benutzeroberfläche (View), der Geschäftslogik (ViewModel) und den Daten (Model/Services) zu gewährleisten.

__Wichtige Komponenten__:

* *View (MainWindow.xaml)*: Zuständig für die Anzeige der Daten und die Interaktion des Benutzers.

* ViewModel (MainViewModel): Hält den Zustand der Anwendung und steuert die Logik über Commands.

* *Models (Json1Order, etc.):* Die Datenstrukturen, die direkt die JSON-Schemata abbilden
* *Services (JsonImportService, ExcelExportService)*: Enthalten die Logik wie das Lesen der Datei, das Parsen, die Validierung und den Export.

* *OrderImportDispatcher*: Dient als Vermittler, um basierend auf den Schlüsselwörtern im JSON (CanHandle Methode) den korrekten Importer (z.B. Json3Importer) auszuwählen.

## Unterstützte JSON-Formate
Das Tool erkennt automatisch das Format der hochgeladenen Datei, indem es nach charakteristischen Feldern sucht.

### Format 1: Einfache Bestellung
Minimalistische Struktur, ohne Artikellisten.
* **Beispielstruktur:**
    ```json
    {
      "orderId": 2000,
      "customer": "Softline GmbH",
      "amount": 4683.16,
      "created": "2023-07-01"
    }
    ```

### Format 2: Bestellung mit Artikeln
Bestellung mit geschachtelter Liste von Artikeln (`items`).
* **Beispielstruktur:**
    ```json
    {
      "orderId": 3000,
      "customer": "GlobalTech GmbH",
      "created": "2023-08-01",
      "items": [
        { "sku": "Z300", "qty": 7, "price": 519.5 }
      ]
    }
    ```

### Format 3: Komplexe Bestellung (Logistik)
Detaillierte Bestellung inklusive Logistik-, Status- und Vertreterinformationen.
* **Beispielstruktur:**
    ```json
    {
      "orderId": 4000,
      "customer": "Innotech AG",
      "created": "2023-10-01",
      "status": "shipped",
      "salesRep": "M. Wagner",
      "delivery": {
        "address": "Musterstraße 12",
        "deliveryDate": "2023-10-04"
      },
      "items": [...]
    }
    ```
