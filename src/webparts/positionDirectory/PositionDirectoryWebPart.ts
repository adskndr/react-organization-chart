{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx/client-side-web-part-manifest.schema.json",
  "id": "3f7a1c2e-9b4d-4e6a-8c2f-1a2b3c4d5e6f",
  "alias": "PositionDirectoryWebPart",
  "componentType": "WebPart",
  "supportsFullBleed": true,

  "version": "*",
  "manifestVersion": 2,

  "requiresCustomScript": false,
  "supportedHosts": ["SharePointWebPart"],

  "preconfiguredEntries": [{
    "groupId": "5c03119e-3074-46fd-976b-c60198311f70",
    "group": { "default": "SPFx - Custom WebParts" },
    "title": { "default": "Positions-Verzeichnis" },
    "description": { "default": "Zeigt alle Personen mit einer bestimmten Position/Jobtitel, unabhängig von Vorgesetzten" },
    "officeFabricIconFontName": "ContactList",
    "properties": {
      "jobTitles": "Lernende, Lernender",
      "groupByDepartment": true,
      "showActionsBar": false
    }
  }]
}
