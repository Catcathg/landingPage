// Fonction pour rechercher un produit par SKU et renvoyer les données
function searchProductBySku(sku) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('final');
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {  // Ignorer la première ligne (en-tête)
        if (data[i][0] == sku) {
            return {
                sku: data[i][0],
                nom: data[i][1],
                effets: data[i][2],
                icone1: data[i][3],
                icone4: data[i][4],
                utilisation: data[i][5],
                precaution: data[i][6],
                composition: data[i][7],
                codebarre: data[i][8],
                nomchinois: data[i][9],
                couleur: data[i][10],
                compo_title: data[i][11],
                brand: data[i][12],
                lot: data[i][13],
                form: data[i][14]
            };
        }
    }
    return null;  // Retourne null si aucun produit n'est trouvé
}



// Fonction pour charger la page HTML
function doGet() {
    return HtmlService.createHtmlOutputFromFile('index')
        .setTitle('Générateur Etiquette CD')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
