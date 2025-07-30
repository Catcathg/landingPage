function doGet(e) {
    // Récupérer le paramètre de page ou utiliser "Index" par défaut
    var page = e.parameter.page || "Index";
    var allowedPages = ["Index", "Preview-NON", "Preview-OUI"]; // Liste des pages autorisées

    // Vérifier si la page demandée est autorisée
    if (!allowedPages.includes(page)) {
        page = "Index"; // Si la page n'existe pas, retourner à la page par défaut
    }

    // Créer le template à partir de la page
    var template = HtmlService.createTemplateFromFile(page);

    // Évaluer le template et configurer les options importantes pour le responsive
    var html = template.evaluate()
        .setTitle("Testez votre énergie féminine")
        .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    return html;
}
