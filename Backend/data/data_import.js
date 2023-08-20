const XLSX = require('xlsx');
const fs = require('fs');
const { MongoClient } = require('mongodb');

// Konfiguration für MongoDB
const mongoURL = 'mongodb://127.0.0.1:27017'; // Passen Sie die URL entsprechend an
const dbName = 'YACOID'; // Geben Sie den Namen Ihrer Datenbank ein

// Array mit den Informationen für die einzelnen Sammlungen und ihre Spalten
const collectionsConfig = [
  {
    collectionName: 'authors',
    excelFileConfigs: [
      {
        excelFilePath: './34_DefsOfIntelligence_AGISIsurvey.xlsx',
        sheetName: 'MI and HI Defs AGISI',
        columnsToExtract: [
          { name: 'Id.', column: 0 },
          { name: 'Author(s)', column: 3 },
        ],
      },
	  {
        excelFilePath: './71_DefsOfIntelligence_LeggHutter2007_extendedTable.xlsx',
        sheetName: 'Legg-Hutter Defs',
        columnsToExtract: [
          { name: 'Id.', column: 0 },
          { name: 'Author(s)', column: 3 }
        ]
      },
	  {
        excelFilePath: './125_humanIntelligence_suggestedDefs_AGISIsurvey.xlsx',
        sheetName: 'HI suggested Defs',
        columnsToExtract: [
          { name: 'Id.', column: 0 },
          { name: 'Author(s)', column: 3 }
        ]
      },
	  {
        excelFilePath: './213_machineIntelligence_suggestedDefs_AGISIsurvey.xlsx',
        sheetName: 'MI suggested Defs',
        columnsToExtract: [
          { name: 'Id.', column: 0 },
          { name: 'Author(s)', column: 3 }
        ]
      },
      // Weitere Konfigurationen für weitere Spalten aus der ersten Datei
      // {
      //   excelFilePath: 'Pfad zur ersten Datei',
      //   sheetName: 'Name des Arbeitsblatts',
      //   columnsToExtract: [
      //     { name: 'Spaltenname1', column: 'Spaltenbuchstabe1' },
      //     { name: 'Spaltenname2', column: 'Spaltenbuchstabe2' },
      //     // Weitere Spalten, falls benötigt
      //   ]
      // },
    ]
  },
  {
    collectionName: 'definitions',
    excelFileConfigs: [
      {
        excelFilePath: './34_DefsOfIntelligence_AGISIsurvey.xlsx',
        sheetName: 'MI and HI Defs AGISI',
        columnsToExtract: [
          { name: 'ID', column: 0 },
          { name: 'Category', column: 1 },
		  { name: 'Definition', column: 2 }
        ]
      },
	  {
        excelFilePath: './71_DefsOfIntelligence_LeggHutter2007_extendedTable.xlsx',
        sheetName: 'Legg-Hutter Defs',
        columnsToExtract: [
          { name: 'Id.', column: 0 },
          { name: 'Category', column: 1 },
		  { name: 'Definition', column: 2 }
        ]
      },
	  {
        excelFilePath: './125_humanIntelligence_suggestedDefs_AGISIsurvey.xlsx',
        sheetName: 'HI suggested Defs',
        columnsToExtract: [
          { name: 'Id.', column: 0 },
          { name: 'Category', column: 1 },
		  { name: 'Definition', column: 2 }
        ]
      },
	  {
        excelFilePath: './213_machineIntelligence_suggestedDefs_AGISIsurvey.xlsx',
        sheetName: 'MI suggested Defs',
        columnsToExtract: [
          { name: 'Id.', column: 0 },
          { name: 'Category', column: 1 },
		  { name: 'Definition', column: 2 }
        ]
      },
    ]
  },
  {
    collectionName: 'sources',
    excelFileConfigs: [
      {
        excelFilePath: './34_DefsOfIntelligence_AGISIsurvey.xlsx',
        sheetName: 'MI and HI Defs AGISI',
        columnsToExtract: [
          { name: 'ID', column: 0 },
          { name: 'Source', column: 4 },
		  { name: 'Year', column: 5 }
        ]
      },
	  {
        excelFilePath: './71_DefsOfIntelligence_LeggHutter2007_extendedTable.xlsx',
        sheetName: 'Legg-Hutter Defs',
        columnsToExtract: [
          { name: 'Id.', column: 0 },
          { name: 'Source', column: 4 },
		  { name: 'Year', column: 5 }
        ]
      },
	  {
        excelFilePath: './125_humanIntelligence_suggestedDefs_AGISIsurvey.xlsx',
        sheetName: 'HI suggested Defs',
        columnsToExtract: [
          { name: 'Id.', column: 0 },
          { name: 'Source', column: 4 },
		  { name: 'Year', column: 5 }
        ]
      },
	  {
        excelFilePath: './213_machineIntelligence_suggestedDefs_AGISIsurvey.xlsx',
        sheetName: 'MI suggested Defs',
        columnsToExtract: [
          { name: 'Id.', column: 0 },
          { name: 'Source', column: 4 },
		  { name: 'Year', column: 5 }
        ]
      },
    ]
  },
];

// Funktion zum Einlesen und Verarbeiten der Excel-Datei
async function processExcelFile(excelFileConfig) {
  const { excelFilePath, sheetName, columnsToExtract } = excelFileConfig;

  const workbook = XLSX.readFile(excelFilePath);
  const worksheet = workbook.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(worksheet, {
    header: 1, // Wir deaktivieren die Kopfzeile
    range: 1, // Wir überspringen die Kopfzeile
    raw: false, // Wir setzen raw auf false, um Werte statt Zell-Objekte zu erhalten
    dateNF: 'yyyy-mm-dd', // Setzen Sie dieses Format entsprechend Ihrer Datumsangaben
    defval: '', // Leere Zellen erhalten einen leeren String statt undefined
    blankrows: false, // Leere Zeilen ignorieren
    rawNumbers: false, // Werte in Rohform, um Typen zu erhalten
  });

  // Transformieren der Daten in das gewünschte Format für die MongoDB
  const transformedData = data.map(row => {
    const document = {};
    for (const { name, column } of columnsToExtract) {
      // Prüfen, ob die Spalte in der aktuellen Zeile vorhanden ist
      if (row[column] !== undefined) {
        document[name] = row[column];
      } else {
        document[name] = null; // Setzen Sie den Wert auf null, wenn die Spalte fehlt oder leer ist
      }
    }
    return document;
  });

  return transformedData;
}

// Funktion zum Einfügen von Daten in eine MongoDB-Sammlung
async function insertDataIntoCollection(collectionName, data) {
  const client = new MongoClient(mongoURL);
  try {
    await client.connect();
    const db = client.db(dbName);
    const collection = db.collection(collectionName);
    await collection.insertMany(data);
    console.log(`Daten erfolgreich in Sammlung ${collectionName} eingefügt.`);
  } catch (err) {
    console.error(`Fehler beim Einfügen der Daten in Sammlung ${collectionName}:`, err);
  } finally {
    client.close();
  }
}

// Funktion zum Löschen aller Daten aus einer MongoDB-Sammlung
async function deleteDataFromCollection(collectionName) {
  const client = new MongoClient(mongoURL);
  try {
    await client.connect();
    const db = client.db(dbName);
    const collection = db.collection(collectionName);
    await collection.deleteMany({});
    console.log(`Alle Daten aus Sammlung ${collectionName} wurden gelöscht.`);
  } catch (err) {
    console.error(`Fehler beim Löschen der Daten aus Sammlung ${collectionName}:`, err);
  } finally {
    client.close();
  }
}

// Funktion zum Verarbeiten aller Konfigurationen und Einfügen der Daten in die Sammlungen
async function processAllCollections() {
  for (const { collectionName, excelFileConfigs } of collectionsConfig) {
    await deleteDataFromCollection(collectionName); // Daten vor dem Einfügen löschen
    for (const excelFileConfig of excelFileConfigs) {
      const data = await processExcelFile(excelFileConfig);
      await insertDataIntoCollection(collectionName, data);
    }
  }
}

processAllCollections();
