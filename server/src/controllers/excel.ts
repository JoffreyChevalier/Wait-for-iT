import Excel from 'exceljs';
import path from 'path';

export const exportCountriesFile = async (): Promise<void> => {
  // Interface de ce que l'on souhaite dans le fichier excel (colonnes)
  interface Country {
    name: string;
    countryCode: string;
    capital: string;
    phoneIndicator: number;
  }

  // Data fictive
  const countries: Country[] = [
    { name: 'France', capital: 'Paris', countryCode: 'FR', phoneIndicator: 33 },

    { name: 'Japan', capital: 'Tokyo', countryCode: 'JP', phoneIndicator: 81 },
  ];

  try {
    // Créer un nouveau fichier excel
    const workbook = new Excel.Workbook();

    // Créer une feuille dans le fichier excel
    const worksheet = workbook.addWorksheet('Countries List');

    // Définir le nom des colonnes de notre feuille excel
    worksheet.columns = [
      { key: 'name', header: 'Name' },
      { key: 'countryCode', header: 'Country Code' },
      { key: 'capital', header: 'Capital' },
      { key: 'phoneIndicator', header: 'International Direct Dialling' },
    ];

    // Envoyer la data dans notre feuille excel
    countries.forEach((item) => {
      worksheet.addRow(item);
    });

    // Mettre en page notre fichier excel (bonus)
    worksheet.columns.forEach((sheetColumn) => {
      sheetColumn.font = {
        size: 12,
      };
      sheetColumn.width = 30;
    });
    // Met en gras et augmente la taille de la première ligne
    worksheet.getRow(1).font = {
      bold: true,
      size: 13,
    };

    // Générer le fichier excel
    const exportPath = path.resolve(
      '/Users/sandra/Desktop/WCS/Wait-for-it/Wait-for-iT/server/src/static',
      'countries.xlsx'
    );
    await workbook.xlsx.writeFile(exportPath);
  } catch (e) {
    console.log(e);
  }
};
