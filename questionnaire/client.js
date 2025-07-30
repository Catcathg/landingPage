// 获取指定问题
function getQuestion(questionId) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var questionSheet = ss.getSheetByName("question");
    var optionSheet = ss.getSheetByName("option");

    var questionData = questionSheet.getDataRange().getValues();
    var optionData = optionSheet.getDataRange().getValues();

    var questionText = "";
    var isRequired = false;
    var options = [];

    // 遍历 question 表，找到匹配的 questionId
    for (var i = 1; i < questionData.length; i++) {
        if (questionData[i][0] == questionId) {
            questionText = questionData[i][1];
            isRequired = questionData[i][2] == "Y";  // 确保 isRequired 只返回 true/false
            break;
        }
    }

    // 遍历 option 表，找到所有与 questionId 相关的选项
    for (var j = 1; j < optionData.length; j++) {
        if (optionData[j][0] == questionId) {
            options.push({
                optionId: optionData[j][1],
                optionText: optionData[j][2],
                optionScore: optionData[j][3]
            });
        }
    }

    return { questionId, questionText, isRequired, options };
}


function getNextQuestionId(questionId, answerId) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var switchSheet = ss.getSheetByName("switch");
    var switchData = switchSheet.getDataRange().getValues();

    for (var i = 1; i < switchData.length; i++) {
        if (switchData[i][0] === questionId && switchData[i][1] === answerId) {
            return switchData[i][2]; // Retourne `nextQuestionId`
        }
    }
    return null; // Retourne `null` si pas trouvé
}

function recordAnswer(userId, questionId, answerId) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var answerSheet = ss.getSheetByName("answer");
    var optionSheet = ss.getSheetByName("option");
    var userSheet = ss.getSheetByName("user");

    // 📌 Trouver le score de l'option choisie
    var optionData = optionSheet.getDataRange().getValues();
    var optionScore = 0;

    for (var i = 1; i < optionData.length; i++) {
        if (optionData[i][0] === questionId && optionData[i][1] === answerId) {
            optionScore = parseInt(optionData[i][3]) || 0; // Assure un nombre
            break;
        }
    }

    console.log("✅ Score récupéré pour", userId, ":", optionScore);

    // 📌 Enregistrer la réponse dans "answer"
    answerSheet.appendRow([userId, questionId, answerId, optionScore]);

    // 📌 Mettre à jour sumScore et lastQuestion dans "user"
    var userData = userSheet.getDataRange().getValues();
    var userRow = -1;
    var sumScore = 0;
    var lastQuestion = ""; // On initialise lastQuestion comme une variable vide

    for (var i = 1; i < userData.length; i++) {
        if (userData[i][0] === userId) { // Trouver la ligne de l'utilisateur
            userRow = i + 1;

            // ✅ Vérifier si sumScore existe avant d'ajouter
            sumScore = parseInt(userData[i][2]) || 0; // Assure un nombre
            sumScore += optionScore; // Ajouter le score de la réponse

            // 📌 Si `lastQuestion` existe déjà, on la garde
            lastQuestion = userData[i][4] || ""; // Si lastQuestion existe, on la conserve

            // 📌 Déterminer la dernière question en fonction de la réponse à la question actuelle
            if (answerId === "Q1_A") {
                lastQuestion = "Q16"; // Si l'utilisateur choisit Q1_A, la prochaine question est Q16
            } else if (answerId === "Q1_B") {
                lastQuestion = "Q21"; // Si l'utilisateur choisit Q1_B, la prochaine question est Q21
            } else {
                // Si aucune option spécifique n'est choisie, garder la logique actuelle
                var nextQuestionId = getNextQuestionId(questionId, answerId);
                if (nextQuestionId) {
                    lastQuestion = nextQuestionId;
                }
            }
            break;
        }
    }

    // 📌 Mise à jour du score et de la dernière question dans la feuille "user"
    if (userRow !== -1) {
        userSheet.getRange(userRow, 3).setValue(sumScore); // Mettre à jour sumScore
        userSheet.getRange(userRow, 5).setValue(lastQuestion); // Mettre à jour lastQuestion
    }

    console.log("✅ Mise à jour effectuée pour l'utilisateur", userId, "avec lastQuestion :", lastQuestion);
}


function getFinalMessage(userId) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var userSheet = ss.getSheetByName("user");
    var answerSheet = ss.getSheetByName("answer");
    var categorySheet = ss.getSheetByName("category");
    var personaSheet = ss.getSheetByName("persona");

    var userData = userSheet.getDataRange().getValues();
    var answerData = answerSheet.getDataRange().getValues();
    var categoryData = categorySheet.getDataRange().getValues();
    var personaData = personaSheet.getDataRange().getValues();

    var sumScore = 0;
    var q1Answer = "";
    var additionalMessages = [];
    var hasPersonaResponse = false; // Indicateur pour vérifier s'il y a une réponse dans personaData

    // 📌 Récupérer sumScore et la réponse à Q1
    for (var i = 1; i < userData.length; i++) {
        if (userData[i][0] === userId) {
            sumScore = parseInt(userData[i][2]); // Colonne C : sumScore
            break;
        }
    }

    // 📌 Récupérer les réponses aux questions spécifiques
    for (var j = 1; j < answerData.length; j++) {
        if (answerData[j][0] === userId) {
            var questionId = answerData[j][1]; // Colonne B : questionId
            var answerId = answerData[j][2];   // Colonne C : answerId

            // 📌 Stocker la réponse à Q1
            if (questionId === "Q1") {
                q1Answer = answerId;
            }

            // 📌 Vérifier les réponses qui nécessitent des messages supplémentaires via la feuille "persona"
            for (var p = 1; p < personaData.length; p++) {
                if (personaData[p][0] === questionId && personaData[p][1] === answerId) {
                    additionalMessages.push(personaData[p][2]); // Colonne C : description
                    hasPersonaResponse = true; // Marquer qu'il y a au moins une réponse personnalisée
                }
            }
        }
    }

    console.log("📌 Vérification du score et des réponses à Q1 :", userId, q1Answer, sumScore);

    // 📌 Obtenir la description de la catégorie depuis Google Sheets
    var categoryDescription = "";
    for (var c = 1; c < categoryData.length; c++) {
        var minScore = parseInt(categoryData[c][0]); // Colonne A : minScore
        var maxScore = parseInt(categoryData[c][1]); // Colonne B : maxScore

        if (sumScore >= minScore && sumScore <= maxScore) {
            categoryDescription = categoryData[c][2]; // Colonne C : description
            break;
        }
    }

    // 📌 Modifier la description de la catégorie pour intégrer les conseils personnalisés
    if (hasPersonaResponse && categoryDescription.includes("[La section de personnalisation]")) {
        // Générer le contenu des conseils personnalisés avec style pour les liens
        var personaContent = `
    <p style="
      display: flex;
      justify-content: center;
      flex-wrap: wrap;
      gap: 10px;
    ">
      ${additionalMessages.map(msg => {
            // Vérifier si le message contient un lien <a>
            if (msg.includes("<a href=")) {
                // Ajouter le style au lien
                msg = msg.replace(/<a href=/g, '<a style="color: black; text-decoration: none;" href=');
            }

            return `
        <button class="button-persona" style="
          font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
          padding: 15px;
          font-size: 1rem;
          margin-top: 10px;
          background: #d88c9a;
          color: black;
          border: none;
          border-radius: 8px;
          cursor: pointer;
          transition: all 0.3s ease-in-out;">
          ${msg}
        </button>
        `;
        }).join("")}
    </p>
  `;

        // Remplacer la balise [La section de personnalisation] par le contenu personnalisé
        categoryDescription = categoryDescription.replace("[La section de personnalisation]", personaContent);
    } else if (!hasPersonaResponse) {
        // Si aucune réponse personnalisée n'a été choisie, masquer la div "advice" avec display:none
        categoryDescription = categoryDescription.replace(/<div class="conseils-container" id="advice">/, '<div class="conseils-container" id="advice" style="display:none;">');
    }


    var messageFinal = `
    <div class="quiz-result">
      <p class="quiz-score">${categoryDescription}</p>
    </div>
  `;

    return messageFinal;
}


function storeUserId(userId) {
    var sheetUser = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("user");
    var sheetAnswer = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("answer");

    if (!sheetUser) {
        console.error("❌ Erreur : La feuille 'user' n'existe pas.");
        return;
    }

    if (!sheetAnswer) {
        console.error("❌ Erreur : La feuille 'answer' n'existe pas.");
        return;
    }

    var dataUser = sheetUser.getDataRange().getValues();
    var dataAnswer = sheetAnswer.getDataRange().getValues();
    var userExists = false;
    var userData = { answerToQ1: "", lastQuestion: "Q21" };

    // Recherche de l'utilisateur dans la feuille 'user'
    for (var i = 1; i < dataUser.length; i++) {
        if (dataUser[i][0] === userId) {
            // Si l'utilisateur existe, on récupère ses données
            userExists = true;
            userData.lastQuestion = dataUser[i][4]; // Récupère la dernière question de la feuille 'user'
            break;
        }
    }

    if (userExists) {
        // Recherche de la réponse à la question Q1 dans la feuille 'answer'
        for (var j = 1; j < dataAnswer.length; j++) {
            if (dataAnswer[j][0] === userId && dataAnswer[j][1] === "Q1") {
                userData.answerToQ1 = dataAnswer[j][2]; // On suppose que 'answerId' est dans la colonne 3 de 'answer'
                break;
            }
        }
        console.log("ℹ️ userId existe déjà :", userId);
    } else {
        // Si l'utilisateur n'existe pas, on l'ajoute avec des valeurs par défaut
        var timestamp = new Date(); // ✅ Ajoute la date et l'heure
        var sumScore = 0; // Valeur par défaut du score
        sheetUser.appendRow([userId, timestamp, sumScore, "", userData.lastQuestion]); // ✅ Ajout des données dans 'user'
        console.log("✅ userId stocké :", userId);
    }

    return userData;  // Retourne les données de l'utilisateur (réponse à Q1 et dernière question)
}



function getAllQuestions(userId) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var questionSheet = ss.getSheetByName("question");
    var optionSheet = ss.getSheetByName("option");

    if (!questionSheet || !optionSheet) {
        console.error("❌ Erreur : Une des feuilles 'question' ou 'option' est introuvable !");
        return [];
    }

    var questionData = questionSheet.getDataRange().getValues();
    var optionData = optionSheet.getDataRange().getValues();

    var questions = [];

    // 📌 Construire les questions avec leurs options
    for (var i = 1; i < questionData.length; i++) {
        var questionId = questionData[i][0];  // Colonne A : ID de la question
        var questionText = questionData[i][1]; // Colonne B : Texte de la question

        // 🔹 Trouver les options correspondantes
        var options = [];
        for (var j = 1; j < optionData.length; j++) {
            if (optionData[j][0] === questionId) { // Colonne A dans "option" = ID de la question
                options.push({
                    optionId: optionData[j][1], // Colonne B : ID de l'option (A, B, C, D)
                    optionText: optionData[j][2] // Colonne C : Texte de l'option
                });
            }
        }

        // 🔹 Ajouter la question avec ses options
        questions.push({
            questionId: questionId,
            questionText: questionText,
            options: options
        });
    }

    return questions; // ✅ Retourne toutes les questions avec leurs options
}


function storeUserEmail(userId, email) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var userSheet = ss.getSheetByName("user");

    if (!userSheet) {
        console.error("❌ Erreur : La feuille 'user' n'existe pas.");
        return;
    }

    var userData = userSheet.getDataRange().getValues();
    var userRow = -1;

    // 📌 Trouver la ligne correspondante à l'userId
    for (var i = 1; i < userData.length; i++) {
        if (userData[i][0] === userId) {
            userRow = i + 1;
            break;
        }
    }

    if (userRow !== -1) {
        userSheet.getRange(userRow, 4).setValue(email); // 📌 Met à jour la colonne Email (D)
        console.log("✅ Email enregistré pour", userId, ":", email);
    } else {
        console.error("❌ Utilisateur non trouvé dans la feuille 'user'.");
    }
}


function getAllSwitches() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var switchSheet = ss.getSheetByName("switch");
    var switchData = switchSheet.getDataRange().getValues();

    var switchMap = {};
    for (var i = 1; i < switchData.length; i++) {
        var key = switchData[i][0] + "_" + switchData[i][1]; // "Q1_Q1_A"
        switchMap[key] = switchData[i][2]; // Stocke `nextQuestionId`
    }

    return switchMap; // Retourne un objet avec toutes les transitions
}

function updateLastQuestion(userId, answerToQ1) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("user");
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
        if (data[i][0] == userId) {
            var lastQuestion = (answerToQ1 == 'Q1_A') ? 'Q16' : 'Q21';
            sheet.getRange(i + 1, 5).setValue(lastQuestion);
            break;
        }
    }
}

// Fonction pour vérifier si la dernière question a été répondue
function checkLastQuestionAnswered(userId, lastQuestion) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("answer");
    var answers = sheet.getDataRange().getValues();

    // Trouvez la ligne correspondante à l'utilisateur
    var userRow = answers.find(row => row[0] === userId);  // Supposons que la première colonne soit l'ID utilisateur
    if (!userRow) {
        return false;  // Si l'utilisateur n'a pas de réponse
    }

    // Vérifiez si la dernière question (Q16 ou Q21) est répondue
    var lastQuestionAnswered = userRow[1] === lastQuestion;  // Supposons que les réponses sont dans la deuxième colonne
    return lastQuestionAnswered;
}


function getLastQuestionFromSheet(userId) {
    // 📌 Accéder à la feuille "user" et récupérer les données
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("user");
    var data = sheet.getDataRange().getValues(); // Récupérer toutes les données de la feuille

    var lastQuestion = ''; // Valeur à retourner

    // 📌 Parcourir les lignes pour trouver la ligne où "Q16" ou "Q21" apparaît
    for (var i = 0; i < data.length; i++) {
        if (data[i][0] === userId) { // Vérifier si l'utilisateur correspond
            var currentQuestion = data[i][4]; // Colonne "question" (colonne E)
            if (currentQuestion === 'Q16' || currentQuestion === 'Q21') {
                lastQuestion = currentQuestion;
                break; // Sortir de la boucle une fois trouvé
            }
        }
    }

    return lastQuestion; // Retourner la dernière question trouvée
}



