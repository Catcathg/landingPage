//version avec page blanche avant la couverture de fin 

function doGet() {
    return HtmlService.createHtmlOutputFromFile('Page')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .setTitle('Générateur PDF - Interface Web');
}

function generatePdfFromSheets() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName("configuration");

    // Vérification de l'existence de la feuille config
    if (!configSheet) {
        throw new Error("La feuille 'configuration' n'existe pas");
    }
    // === 新插入的函数 ===
    // 插入在 const categories = [ 之前

    // 函数1: 提取分类数据
    function extractCategoriesFromConfig(configSheet) {
        const categoriesRange = configSheet.getRange("G2:G20");
        const imagesRange = configSheet.getRange("H2:H20");
        const dateEdition = configSheet.getRange("N2");
        const categoriesData = categoriesRange.getValues();
        const imagesData = imagesRange.getValues();
        const categories = [];

        categoriesData.forEach((row, index) => {
            const categoryName = row[0];
            const imageUrl = imagesData[index] ? imagesData[index][0] : "";

            if (categoryName && categoryName.toString().trim() !== "") {
                categories.push({
                    name: categoryName.toString().trim(),
                    imageCell: "H" + (2 + index), // H2, H3, H4, etc.
                    sheetName: categoryName.toString().trim()
                });
            }
        });

        return {
            categories: categories,
            dateEdition: dateEdition.getValue()
        };
    }

    // 函数2: 提取表头数据
    function extractHeadersFromConfig(configSheet) {
        const headersRange = configSheet.getRange("J2:J20");
        const headersData = headersRange.getValues();
        const headers = [];

        headersData.forEach((row) => {
            const headerName = row[0];
            if (headerName && headerName.toString().trim() !== "") {
                headers.push(headerName.toString().trim());
            }
        });

        return headers;
    }

    // 函数3: 提取列比例数据
    function extractColumnProportionsFromConfig(configSheet) {
        const proportionsRange = configSheet.getRange("J2:K20");
        const proportionsData = proportionsRange.getValues();
        const proportions = {};

        proportionsData.forEach((row) => {
            const columnName = row[0];
            const proportion = row[1];

            if (columnName && columnName.toString().trim() !== "" &&
                proportion !== "" && !isNaN(proportion)) {
                proportions[columnName.toString().trim()] = parseFloat(proportion);
            }
        });

        return proportions;
    }

    const configData = extractCategoriesFromConfig(configSheet);
    const categories = configData.categories;

    const expectedHeaders = extractHeadersFromConfig(configSheet);

    // Récupération des paramètres de configuration avec vérification
    const itemsPerPage = Number(configSheet.getRange("B2").getValue()) || 25;
    const colorEven = configSheet.getRange("B3").getValue() || "#ffffff";
    const colorOdd = configSheet.getRange("B4").getValue() || "#F0F0F0";
    const fontSize = Number(configSheet.getRange("B5").getValue()) || 6;

    const imageCover = configSheet.getRange("E2").getValue();
    const imageEnd = configSheet.getRange("E3").getValue();

    // Vérification des URLs d'images
    if (!imageCover || !imageEnd) {
        throw new Error("Les URLs des images de couverture et de fin ne sont pas définies");
    }

    // Cache pour les images
    const imageCache = {};

    // Fonction pour récupérer les images de manière sécurisée
    function fetchImageSafely(url) {
        if (!url) return null;

        try {
            if (!imageCache[url]) {
                console.log(`Chargement de l'image: ${url}`);
                imageCache[url] = UrlFetchApp.fetch(url).getBlob();
            }
            return imageCache[url];
        } catch (error) {
            console.error(`Erreur lors du chargement de l'image: ${url}`, error);
            return null;
        }
    }

    // Fonction pour redimensionner et ajouter une image
    function addResizedImage(imageBlob, maxWidth = 500, maxHeight = 700, centered = false) {
        if (!imageBlob) return null;

        try {
            const image = body.appendImage(imageBlob);
            image.setWidth(maxWidth);
            image.setHeight(maxHeight);

            if (centered) {
                const parent = image.getParent();
                if (parent.getType() === DocumentApp.ElementType.PARAGRAPH) {
                    parent.asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
                }
            }

            return image;
        } catch (error) {
            console.error("Erreur lors de l'ajout de l'image redimensionnée:", error);
            return null;
        }
    }

    // Fonction pour créer les en-têtes de tableau adaptatifs
    function createTableHeaders(table, headers, colWidths) {
        const headRow = table.appendTableRow();

        const availableHeaders = expectedHeaders.filter(header =>
            headers.findIndex(col =>
                col && col.toString().trim().toLowerCase() === header.trim().toLowerCase()
            ) !== -1
        );

        availableHeaders.forEach((header, index) => {
            const originalIndex = headers.findIndex(col =>
                col && col.toString().trim().toLowerCase() === header.trim().toLowerCase()
            );
            let label = originalIndex !== -1 ? headers[originalIndex] : header;

            // Remplacer "PRIX PUBLIC TTC" par "PRIX TTC" pour qu'il tienne sur une ligne
            if (label.toString().toUpperCase().includes("PRIX PUBLIC TTC")) {
                label = "PRIX TTC";
            }

            const cell = headRow.appendTableCell(label.toString());
            cell.setFontSize(fontSize + 1); // Header un peu plus grand
            cell.setFontFamily('Poppins');

            // Mettre les headers en gras
            const textElement = cell.editAsText();
            textElement.setBold(0, label.toString().length - 1, true);

            cell.setPaddingTop(1);
            cell.setPaddingBottom(6); // Padding bottom plus important pour créer plus d'espace
            cell.setPaddingLeft(3);
            cell.setPaddingRight(2);
            cell.setVerticalAlignment(DocumentApp.VerticalAlignment.TOP);

            if (index < colWidths.length) {
                cell.setWidth(colWidths[index]);
            }
        });

        // Ligne d'espacement entre headers et données
        const spacerRow = table.appendTableRow();
        spacerRow.setMinimumHeight(1);

        availableHeaders.forEach(() => {
            const spacerCell = spacerRow.appendTableCell("");
            spacerCell.setBackgroundColor("#ffffff");
            spacerCell.setFontSize(1);
            spacerCell.setPaddingTop(0);
            spacerCell.setPaddingBottom(0);
        });

        return availableHeaders;
    }

    let doc = DocumentApp.create("Catalogue Calebasse");
    const docId = doc.getId();
    let body = doc.getBody();

    // Configuration des marges optimisées
    body.setMarginBottom(0);
    body.setMarginLeft(0);
    body.setMarginRight(60);

    // Configuration de la police par défaut pour tout le document
    const style = {};
    style[DocumentApp.Attribute.FONT_FAMILY] = 'Poppins';
    style[DocumentApp.Attribute.FONT_SIZE] = fontSize;
    body.setAttributes(style);

    // Système de comptage séquentiel
    let sequentialPageCounter = 0;
    let operationCount = 0;
    let isDocClosed = false;

    // Fonction pour incrémenter le compteur de page de manière sécurisée
    function incrementPageCounter(reason = "") {
        sequentialPageCounter++;
        console.log(`Page ${sequentialPageCounter} ajoutée ${reason ? '(' + reason + ')' : ''}`);
    }

    // Fonction pour sauvegarder le document périodiquement
    function saveIfNeeded() {
        if (operationCount >= 15) {
            try {
                doc.saveAndClose();
                isDocClosed = true;
                Utilities.sleep(1000);
                operationCount = 0;
            } catch (error) {
                console.error("Erreur lors de la sauvegarde:", error);
            }
        }
        if (isDocClosed) {
            try {
                doc = DocumentApp.openById(docId);
                body = doc.getBody();
                isDocClosed = false;
            } catch (error) {
                console.error("Erreur lors de la réouverture:", error);
                throw error;
            }
        }
    }

    // Fonction pour ajouter du contenu de manière sécurisée
    function appendSafe(callback) {
        try {
            operationCount++;
            saveIfNeeded();
            callback();
        } catch (error) {
            console.error("Erreur lors de l'ajout de contenu:", error);
            throw error;
        }
    }

    // === AJOUT DE LA PAGE DE COUVERTURE ===
    const coverBlob = fetchImageSafely(imageCover);
    if (coverBlob) {
        // Configuration des marges spécifiques pour la couverture
        body.setMarginTop(0);
        body.setMarginBottom(0);
        body.setMarginLeft(0);
        body.setMarginRight(0);

        appendSafe(() => addResizedImage(coverBlob, 700, 950));

        // Ajouter la date d'édition sous l'image de couverture
        appendSafe(() => {
            const editionParagraph = body.appendParagraph("          "); //ajouter espace pour décaler à droite la date catalogue

            // Ajouter la date en italique
            const dateText = editionParagraph.appendText(configData.dateEdition || "");
            dateText.setFontSize(20);
            dateText.setFontFamily('Poppins');
            dateText.setItalic(true);

            // Paramètres du paragraphe
            editionParagraph.setSpacingBefore(0);
            editionParagraph.setSpacingAfter(10);
        });

        appendSafe(() => body.appendPageBreak());
        incrementPageCounter("couverture");

        // Remettre les marges normales après la couverture
        body.setMarginTop(10);
        body.setMarginBottom(0);
        body.setMarginLeft(0);
        body.setMarginRight(60);

        // Ajouter une page blanche après la couverture
        appendSafe(() => {
            const blankPageParagraph = body.appendParagraph(" ");
            blankPageParagraph.setFontSize(1);
            blankPageParagraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
            blankPageParagraph.setSpacingBefore(0);
            blankPageParagraph.setSpacingAfter(0);
        });
        appendSafe(() => body.appendPageBreak());
        incrementPageCounter("page blanche après couverture");
    }

    // === FILTRAGE DES CATÉGORIES VALIDES ===
    const validCategories = categories.filter(cat => {
        const sheet = ss.getSheetByName(cat.sheetName);
        if (!sheet) {
            console.warn(`Feuille '${cat.sheetName}' non trouvée, ignorée`);
            return false;
        }

        const data = sheet.getDataRange().getValues();
        if (data.length <= 1) {
            console.warn(`Aucune donnée dans la feuille '${cat.sheetName}', ignorée`);
            return false;
        }

        const rows = data.slice(1).filter(row => row.some(cell => cell !== ""));
        if (rows.length === 0) {
            console.warn(`Aucune ligne de données dans la feuille '${cat.sheetName}', ignorée`);
            return false;
        }

        console.log(`Catégorie '${cat.name}' validée avec ${rows.length} lignes de données`);
        return true;
    });

    if (validCategories.length === 0) {
        throw new Error("Aucune catégorie avec des données valides trouvée");
    }

    console.log(`${validCategories.length} catégorie(s) valide(s) trouvée(s) sur ${categories.length}`);

    // === TRAITEMENT DES CATÉGORIES ===
    for (const [catIndex, cat] of validCategories.entries()) {
        console.log(`Traitement de la catégorie: ${cat.name}`);

        const sheet = ss.getSheetByName(cat.sheetName);
        const data = sheet.getDataRange().getValues();
        const headers = data[0];
        const rows = data.slice(1).filter(row => row.some(cell => cell !== ""));

        // Ajout de l'image de catégorie
        const imageUrl = configSheet.getRange(cat.imageCell).getValue();
        if (imageUrl) {
            const catImgBlob = fetchImageSafely(imageUrl);
            if (catImgBlob) {
                appendSafe(() => addResizedImage(catImgBlob, 780, 1100, true));
                //appendSafe(() => body.appendPageBreak());
                incrementPageCounter(`image catégorie ${cat.name}`);
            }
        }

        // Traitement des données par chunks
        for (let i = 0; i < rows.length; i += itemsPerPage) {
            const chunk = rows.slice(i, i + itemsPerPage);
            incrementPageCounter(`données ${cat.name} chunk ${Math.floor(i / itemsPerPage) + 1}`);

            // Titre de la catégorie avec espacement minimal
            appendSafe(() => {
                const titleParagraph = body.appendParagraph(cat.name.toUpperCase());

                titleParagraph.setFontSize(14); // Nom de catégorie encore plus grand
                titleParagraph.setFontFamily('Poppins');

                // Mettre le nom de catégorie en gras
                const titleTextElement = titleParagraph.editAsText();
                titleTextElement.setBold(0, cat.name.length - 1, true);

                titleParagraph.setSpacingBefore(2);
                titleParagraph.setSpacingAfter(8);

                if (sequentialPageCounter % 2 === 0) {
                    // Page paire : aligné à gauche avec marge
                    titleParagraph.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
                    titleParagraph.setIndentStart(17);
                    titleParagraph.setIndentEnd(0);
                } else {
                    // Page impaire : aligné à droite, parfaitement aligné avec la colonne RÉFÉRENCE
                    titleParagraph.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
                    titleParagraph.setIndentStart(0);
                    titleParagraph.setIndentEnd(-55);
                }
            });

            // Création du tableau optimisé pour 25 lignes
            appendSafe(() => {
                const table = body.appendTable();
                table.setBorderWidth(0);

                // Proportions ajustées pour que tout tienne sur une ligne
                const columnProportions = extractColumnProportionsFromConfig(configSheet);

                const availableHeaders = expectedHeaders.filter(header =>
                    headers.findIndex(col =>
                        col && col.toString().trim().toLowerCase() === header.trim().toLowerCase()
                    ) !== -1
                );

                const totalProportions = availableHeaders.reduce((sum, header) => {
                    return sum + (columnProportions[header] || 1);
                }, 0);

                const totalWidth = 600;
                const colWidths = availableHeaders.map(header => {
                    const proportion = columnProportions[header] || 1;
                    return Math.round((proportion / totalProportions) * totalWidth);
                });

                const headersUsed = createTableHeaders(table, headers, colWidths);
                const headerIndexes = availableHeaders.map(col =>
                    headers.findIndex(h => h && h.toString().trim().toLowerCase() === col.trim().toLowerCase())
                );

                // Lignes de données avec hauteur fixe
                const fixedRowHeight = 28.8; // HAUTEUR FIXE pour toutes les lignes (ajustée)
                let totalDataRowHeight = 0;

                chunk.forEach((row, index) => {
                    const tr = table.appendTableRow();

                    // Hauteur fixe pour toutes les lignes
                    tr.setMinimumHeight(fixedRowHeight);
                    totalDataRowHeight += fixedRowHeight;

                    const rowColor = index % 2 === 0 ? colorOdd : colorEven;

                    headerIndexes.forEach((headerIndex, colIdx) => {
                        let value = headerIndex !== -1 ? row[headerIndex] : "";
                        const rawText = value ? value.toString().trim() : "";
                        const cell = tr.appendTableCell();
                        const header = availableHeaders[colIdx];
                        const cellParagraph = cell.getChild(0).asParagraph();

                        // Formatage selon le type de colonne
                        switch (header) {
                            case "RÉFÉRENCE":
                                cell.setText(rawText.toUpperCase());
                                break;
                            case "NOM":
                                if (rawText) {
                                    const formatted = rawText.charAt(0).toUpperCase() + rawText.slice(1).toLowerCase();
                                    cell.setText(formatted);
                                    const textElement = cell.editAsText();
                                    textElement.setBold(0, formatted.length - 1, true);
                                }
                                break;
                            case "NOM LATIN":
                                if (rawText) {
                                    const formatted = rawText.charAt(0).toUpperCase() + rawText.slice(1).toLowerCase();
                                    cell.setText(formatted);
                                    const textElement = cell.editAsText();
                                    textElement.setItalic(0, formatted.length - 1, true);
                                }
                                break;
                            case "PRIX PRO HT":
                            case "PRIX PUBLIC TTC":
                                if (rawText) {
                                    const price = parseFloat(rawText);
                                    if (!isNaN(price)) {
                                        const priceText = "€" + price.toFixed(2);
                                        cell.setText(priceText);
                                        const textElement = cell.editAsText();
                                        textElement.setBold(0, priceText.length - 1, true);
                                    } else {
                                        cell.setText(rawText);
                                    }
                                } else {
                                    cell.setText("");
                                }
                                break;
                            case "TVA":
                                if (rawText) {
                                    const percentage = parseFloat(rawText);
                                    if (!isNaN(percentage)) {
                                        const finalPercentage = percentage < 1 ? percentage * 100 : percentage;
                                        cell.setText(finalPercentage.toFixed(2) + "%");
                                    } else {
                                        cell.setText(rawText);
                                    }
                                } else {
                                    cell.setText("");
                                }
                                break;
                            case "REMISE":
                                if (rawText !== null && rawText !== undefined && rawText !== "") {
                                    const percentage = parseFloat(rawText);
                                    if (!isNaN(percentage)) {
                                        const finalPercentage = percentage < 1 ? percentage * 100 : percentage;
                                        cell.setText(Math.round(finalPercentage) + "%");
                                    } else {
                                        cell.setText(rawText);
                                    }
                                } else {
                                    cell.setText("0%"); // ou "" selon ce que vous préférez
                                }
                                break;
                            default:
                                cell.setText(rawText.toLowerCase());
                        }

                        cell.setFontSize(fontSize + 0.5); // Données un peu plus grandes
                        cell.setFontFamily('Poppins');
                        cell.setPaddingTop(3);
                        cell.setPaddingBottom(3);
                        cell.setPaddingLeft(3);
                        cell.setPaddingRight(3);
                        cell.setVerticalAlignment(DocumentApp.VerticalAlignment.TOP);
                        cell.setBackgroundColor(rowColor);
                        cellParagraph.setLineSpacing(0.7);
                    });
                });

                // Lignes vides avec hauteur fixe
                const missingRows = itemsPerPage - chunk.length;
                if (missingRows > 0) {
                    for (let emptyRowIndex = 0; emptyRowIndex < missingRows; emptyRowIndex++) {
                        const emptyTr = table.appendTableRow();
                        emptyTr.setMinimumHeight(28.8); // Changé pour plus compact

                        availableHeaders.forEach(() => {
                            const emptyCell = emptyTr.appendTableCell("");
                            emptyCell.setFontSize(fontSize + 0.5);
                            emptyCell.setFontFamily('Poppins');
                            emptyCell.setPaddingTop(3);
                            emptyCell.setPaddingBottom(3);
                            emptyCell.setPaddingLeft(3);
                            emptyCell.setPaddingRight(3);
                            emptyCell.setBackgroundColor("#ffffff");
                        });
                    }
                }
            });

            // Numérotation avec espacement contrôlé
            if (sequentialPageCounter >= 4) {
                appendSafe(() => {
                    // Espacement contrôlé entre le tableau et le numéro de page
                    const spacer = body.appendParagraph("");
                    spacer.setFontSize(1);
                    spacer.setSpacingBefore(2); // Espacement original maintenu
                    spacer.setSpacingAfter(0);

                    // Créer le paragraphe de numérotation directement
                    const pageNumberParagraph = body.appendParagraph(String(sequentialPageCounter));
                    pageNumberParagraph.setFontSize(8);
                    pageNumberParagraph.setFontFamily('Poppins');

                    // Espacement du numéro de page
                    pageNumberParagraph.setSpacingBefore(0);
                    pageNumberParagraph.setSpacingAfter(0);

                    // Alignement selon page paire/impaire
                    if (sequentialPageCounter % 2 === 0) {
                        // Page paire : aligné à gauche
                        pageNumberParagraph.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
                        pageNumberParagraph.setIndentStart(17);
                        pageNumberParagraph.setIndentEnd(0);
                    } else {
                        // Page impaire : aligné à droite
                        pageNumberParagraph.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
                        pageNumberParagraph.setIndentStart(0);
                        pageNumberParagraph.setIndentEnd(-55);
                    }
                });
            }

            // Logique de saut de page entre chunks
            const isLastChunkOfCategory = (i + itemsPerPage >= rows.length);
            if (!isLastChunkOfCategory) {
                appendSafe(() => body.appendPageBreak());
            }
        }

        // === GESTION DES PAGES BLANCHES ENTRE CATÉGORIES ===
        const isLastCategory = (catIndex === validCategories.length - 1);

        if (!isLastCategory) {
            // Ajouter un saut de page après la fin des données de la catégorie
            appendSafe(() => body.appendPageBreak());
            incrementPageCounter(`fin catégorie ${cat.name}`);
            // Correction du compteur pour éviter la page en trop
            sequentialPageCounter = sequentialPageCounter - 1;

            console.log(`Fin de catégorie "${cat.name}" sur page ${sequentialPageCounter}`);

            // *** LOGIQUE : Si page IMPAIRE, ajouter une page blanche ***
            if (sequentialPageCounter % 2 === 1) { // Page IMPAIRE
                console.log(`Page ${sequentialPageCounter} est impaire, ajout d'une page blanche`);

                appendSafe(() => {
                    const blankPageParagraph = body.appendParagraph(" ");
                    blankPageParagraph.setFontSize(1);
                    blankPageParagraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
                    blankPageParagraph.setSpacingBefore(0);
                    blankPageParagraph.setSpacingAfter(0);
                });

                // Ajouter la numérotation sur la page blanche
                appendSafe(() => {
                    incrementPageCounter(`page blanche après ${cat.name}`);

                    // Créer un paragraphe espaceur pour pousser le numéro vers le bas
                    const spacerParagraph = body.appendParagraph("");
                    spacerParagraph.setFontSize(1);
                    spacerParagraph.setSpacingBefore(777.5); // Espacement réduit pour les pages blanches
                    spacerParagraph.setSpacingAfter(0);

                    // Créer un paragraphe avec le numéro de page
                    const pageNumberParagraph = body.appendParagraph(String(sequentialPageCounter));
                    pageNumberParagraph.setFontSize(8);
                    pageNumberParagraph.setFontFamily('Poppins');

                    // Espacement réduit sur le numéro lui-même
                    pageNumberParagraph.setSpacingBefore(0);
                    pageNumberParagraph.setSpacingAfter(0);

                    // Alignement selon page paire/impaire
                    pageNumberParagraph.setAlignment(sequentialPageCounter % 2 === 0 ?
                        DocumentApp.HorizontalAlignment.LEFT :
                        DocumentApp.HorizontalAlignment.RIGHT);

                    if (sequentialPageCounter % 2 === 0) {
                        // pageNumberParagraph.setIndentStart(17);
                        //pageNumberParagraph.setIndentEnd(0);
                    } else {
                        //pageNumberParagraph.setIndentStart(0);
                        //pageNumberParagraph.setIndentEnd(-55);
                    }
                });

                appendSafe(() => body.appendPageBreak());

                console.log(`Page blanche avec numérotation ajoutée, prochaine catégorie commencera sur page ${sequentialPageCounter + 1} (paire)`);
            } else { // Page PAIRE
                console.log(`Page ${sequentialPageCounter} est paire, pas de page blanche nécessaire`);
                console.log(`Prochaine catégorie commencera directement sur page ${sequentialPageCounter + 1} (impaire)`);
            }
        }
    }

    // === make sure the total page could be perfectly divided by 4 ===
    const currentPageCount = sequentialPageCounter;
    const totalPagesWithEndCover = currentPageCount + 1;
    const remainder = totalPagesWithEndCover % 4;

    if (remainder !== 0) {
        const pagesToAdd = 4 - remainder;

        for (let i = 0; i < pagesToAdd && i < 3; i++) {
            console.log(`Ajout de la page blanche de remplissage ${i + 1}`);

            appendSafe(() => body.appendPageBreak());
            incrementPageCounter(`page blanche de remplissage ${i + 1}`);

            appendSafe(() => {
                const blankPageParagraph = body.appendParagraph(" ");
                blankPageParagraph.setFontSize(1);
                blankPageParagraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
                blankPageParagraph.setSpacingBefore(0);
                blankPageParagraph.setSpacingAfter(0);
            });

            /* Ajouter la numérotation sur la page blanche si nécessaire
            if (sequentialPageCounter >= 4) {
              appendSafe(() => {
                const spacerParagraph = body.appendParagraph("");
                spacerParagraph.setFontSize(1);
                spacerParagraph.setSpacingBefore(777.5);
                spacerParagraph.setSpacingAfter(0);
      
                const pageNumberParagraph = body.appendParagraph(String(sequentialPageCounter));
                pageNumberParagraph.setFontSize(8);
                pageNumberParagraph.setFontFamily('Poppins');
                pageNumberParagraph.setSpacingBefore(0);
                pageNumberParagraph.setSpacingAfter(0);
      
                if (sequentialPageCounter % 2 === 0) {
                  pageNumberParagraph.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
                  pageNumberParagraph.setIndentStart(17);
                  pageNumberParagraph.setIndentEnd(0);
                } else {
                  pageNumberParagraph.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
                  pageNumberParagraph.setIndentStart(0);
                  pageNumberParagraph.setIndentEnd(-55);
                }
              });
            }*/
        }
    }

    // === AJOUT DE LA PAGE DE FIN ===
    const endBlob = fetchImageSafely(imageEnd);
    /*appendSafe(() => body.appendPageBreak());
    incrementPageCounter("couverture de fin");*/

    if (endBlob) {
        appendSafe(() => body.appendPageBreak());
        appendSafe(() => addResizedImage(endBlob, 780, 1100));
        incrementPageCounter("page de fin");
    }

    // === ALGORITHME AVANCÉ DE SUPPRESSION DES PAGES VIDES ===
    try {
        doc.saveAndClose();
        doc = DocumentApp.openById(docId);
        body = doc.getBody();

        const allElements = body.getChildren();
        console.log(`Analyse de ${allElements.length} éléments dans le document`);

        const contentZones = [];
        let currentZone = [];

        for (let i = 0; i < allElements.length; i++) {
            const element = allElements[i];
            const elementType = element.getType();

            if (elementType === DocumentApp.ElementType.PAGE_BREAK) {
                if (currentZone.length > 0) {
                    contentZones.push([...currentZone]);
                    currentZone = [];
                }
            } else {
                currentZone.push({
                    element: element,
                    index: i,
                    type: elementType
                });
            }
        }

        if (currentZone.length > 0) {
            contentZones.push(currentZone);
        }

        console.log(`${contentZones.length} zones de contenu identifiées`);

        const emptyZones = [];
        contentZones.forEach((zone, zoneIndex) => {
            let hasRealContent = false;
            let contentScore = 0;

            zone.forEach(item => {
                const { element, type } = item;

                switch (type) {
                    case DocumentApp.ElementType.PARAGRAPH:
                        const text = element.asParagraph().getText().trim();

                        if (text === "" || text === " ") {
                            contentScore += 1;
                            break;
                        }

                        if (/^\d+$/.test(text)) {
                            contentScore += 2;
                            break;
                        }

                        const categoryTitles = categories.map(cat => cat.name.toUpperCase());
                        if (categoryTitles.includes(text.trim().toUpperCase())) {
                            contentScore += 3;
                            break;
                        }

                        if (text.length > 3) {
                            hasRealContent = true;
                            contentScore += 10;
                        }
                        break;

                    case DocumentApp.ElementType.TABLE:
                        hasRealContent = true;
                        contentScore += 20;
                        break;

                    case DocumentApp.ElementType.INLINE_IMAGE:
                        hasRealContent = true;
                        contentScore += 15;
                        break;

                    default:
                        contentScore += 1;
                }
            });

            const isEmpty = !hasRealContent && contentScore <= 5;

            if (isEmpty) {
                emptyZones.push(zoneIndex);
                console.log(`Zone ${zoneIndex} identifiée comme vide (score: ${contentScore})`);
            } else {
                console.log(`Zone ${zoneIndex} contient du contenu (score: ${contentScore})`);
            }
        });

        let elementsToRemove = [];

        for (let i = emptyZones.length - 1; i >= 0; i--) {
            const emptyZoneIndex = emptyZones[i];
            const zone = contentZones[emptyZoneIndex];

            zone.forEach(item => {
                elementsToRemove.push(item.element);
            });

            if (emptyZoneIndex > 0) {
                let elementCount = 0;

                for (let j = 0; j < emptyZoneIndex; j++) {
                    elementCount += contentZones[j].length;
                    if (j < emptyZoneIndex - 1) {
                        elementCount++;
                    }
                }

                if (elementCount < allElements.length) {
                    const pageBreakElement = allElements[elementCount];
                    if (pageBreakElement && pageBreakElement.getType() === DocumentApp.ElementType.PAGE_BREAK) {
                        elementsToRemove.push(pageBreakElement);
                    }
                }
            }
        }

        let removedCount = 0;
        const uniqueElementsToRemove = [...new Set(elementsToRemove)];

        uniqueElementsToRemove.forEach(element => {
            try {
                element.removeFromParent();
                removedCount++;
            } catch (removeError) {
                console.warn("Impossible de supprimer un élément:", removeError);
            }
        });

        console.log(`${removedCount} éléments supprimés (${emptyZones.length} zones vides éliminées)`);

        const cleanedElements = body.getChildren();
        let consecutiveBreaksToRemove = [];

        for (let i = 0; i < cleanedElements.length - 1; i++) {
            const current = cleanedElements[i];
            const next = cleanedElements[i + 1];

            if (current.getType() === DocumentApp.ElementType.PAGE_BREAK &&
                next.getType() === DocumentApp.ElementType.PAGE_BREAK) {
                consecutiveBreaksToRemove.push(next);
            }
        }

        consecutiveBreaksToRemove.forEach(element => {
            try {
                element.removeFromParent();
            } catch (error) {
                console.warn("Erreur suppression saut de page consécutif:", error);
            }
        });

        if (consecutiveBreaksToRemove.length > 0) {
            console.log(`${consecutiveBreaksToRemove.length} sauts de page consécutifs supprimés`);
        }

        const finalElements = body.getChildren();
        if (finalElements.length > 0) {
            const lastElement = finalElements[finalElements.length - 1];
            if (lastElement.getType() === DocumentApp.ElementType.PAGE_BREAK) {
                try {
                    lastElement.removeFromParent();
                    console.log("Saut de page final supprimé");
                } catch (error) {
                    console.warn("Erreur suppression saut de page final:", error);
                }
            }
        }

    } catch (error) {
        console.warn("Erreur lors de la suppression des pages vides:", error);
    }

    // === SAUVEGARDE FINALE ET GÉNÉRATION PDF ===
    try {
        doc.saveAndClose();

        const pdf = DriveApp.getFileById(docId).getAs("application/pdf");
        const base64Pdf = Utilities.base64Encode(pdf.getBytes());

        console.log("PDF généré avec succès");
        console.log(`Document temporaire créé avec l'ID: ${docId}`);
        console.log("Note: Vous pouvez supprimer manuellement le document temporaire depuis Google Drive si souhaité");

        return base64Pdf;

    } catch (error) {
        console.error("Erreur lors de la génération finale du PDF:", error);
        console.log(`En cas d'erreur, le document temporaire avec l'ID ${docId} peut rester dans Google Drive`);
        throw error;
    }
}