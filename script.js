console.log("script.js загружен");

// Проверим, загрузилась ли библиотека docx
console.log("window.docx:", window.docx);

// Словарь тегов
const tagsDict = {
    'Номер исполнительного производства (ИП)': 'fssp:IpNo',
    'Документ': 'fssp:DocName',
    'Кем вынесен': 'fssp:IdCourtName',
    'Орган, выдавший ИД': 'fssp:IdCourtName',
    'Взыскатель': 'fssp:IdCrdrName',
    'Должник': 'fssp:DbtrName',
    'ИНН должника': 'fssp:DbtrInn',
    'Общая сумма задолженности': 'fssp:IpRestDebtsum',
    'Дата возбуждения ИП': 'fssp:IpRiseDate',
    'Дата вынесения документа': 'fssp:IdDocDate',
    'Кем подписан': 'fssp:IpRisepristName',
    'Адрес органа': 'fssp:IdCourtAdr',
    'Адрес взыскателя': 'fssp:IdCourtAdr',
    'Адрес должника': 'fssp:DbtrAdr',
    'КПП должника': 'fssp:DbtrKpp',
    'Получатель': 'fssp:RecipientName',
    'ИНН получателя': 'fssp:RecipientInn',
    'Счёт получателя': 'fssp:RecipientBankCorAcc',
    'БИК': 'fssp:RecipientBic',
    'Перечисляемая сумма': 'fssp:Amount',
    'ОКТМО получателя': 'fssp:RecipientOktmo',
    'КПП получателя': 'fssp:RecipientKpp',
    'ЕКС': 'fssp:RecipientAccNumber',
    'Банк получателя': 'fssp:RecipientBankName',
    'Пример постановочной части': 'fssp:AdjudicationText',
    'Пример постановочной части 2': 'fssp:ResolutionText',
    'Пример постановочной части 3': 'fssp:IdDebtText'
};

document.getElementById("processButton").addEventListener("click", function () {
    console.log("Кнопка 'Обработать файл' нажата");
    const fileInput = document.getElementById("fileInput");
    if (!fileInput.files.length) {
        alert("Пожалуйста, выберите XML файл.");
        console.log("Файл не выбран");
        return;
    }

    const file = fileInput.files[0];
    console.log("Выбран файл:", file.name);
    const reader = new FileReader();

    reader.onload = function (event) {
        console.log("Файл прочитан успешно");
        const xmlContent = event.target.result;
        console.log("Содержимое файла:", xmlContent.slice(0,100) + "...");
        const processedText = processXML(xmlContent);
        console.log("Обработанный текст:", processedText);
        generateWordFile(processedText);
    };

    reader.onerror = function (error) {
        console.error("Ошибка чтения файла:", error);
    };

    // Просто readAsText без указания кодировки
    reader.readAsText(file);
});

function processXML(xmlContent) {
    let newText = "";
    for (const [oneTagKey, oneTagValue] of Object.entries(tagsDict)) {
        try {
            const pattern = new RegExp(`<${oneTagValue}>([\\s\\S]*?)</${oneTagValue}>`, "i");
            const match = xmlContent.match(pattern);
            if (match && match[1] !== undefined) {
                newText += oneTagKey + ':\n';
                newText += match[1] + '\n\n';
            } else {
                newText += `ОШИБКА. Не найден ${oneTagKey}\n\n`;
            }
        } catch (e) {
            console.error("Ошибка при поиске тега:", oneTagKey, e);
            newText += `ОШИБКА. Не найден ${oneTagKey}\n\n`;
        }
    }

    // Удаляем некоторые фразы
    newText = newText.replace('\nПример постановочной части 3:', '').replace('\nПример постановочной части 2:', '');

    return newText;
}

function generateWordFile(processedText) {
    console.log("Генерация Word файла...");
    const { Document, Packer, Paragraph, TextRun } = window.docx;

    const lines = processedText.split("\n");
    const highlightPhrases = Object.keys(tagsDict);

    const paragraphs = [];
    for (let line of lines) {
        line = line.replace(/\r/g, "");

        if (line.trim().length === 0) {
            paragraphs.push(new Paragraph(""));
            continue;
        }

        let isKeyLine = false;
        let keyPhraseFound = null;
        for (const phrase of highlightPhrases) {
            if (line.startsWith(phrase + ":")) {
                isKeyLine = true;
                keyPhraseFound = phrase + ":";
                break;
            }
        }

        if (isKeyLine && keyPhraseFound) {
            const restText = line.substring(keyPhraseFound.length);
            paragraphs.push(
                new Paragraph({
                    children: [
                        new TextRun({
                            text: keyPhraseFound,
                            bold: true,
                            color: "999999"
                        }),
                        new TextRun({
                            text: restText
                        })
                    ]
                })
            );
        } else {
            paragraphs.push(new Paragraph({
                children: [new TextRun(line)]
            }));
        }
    }

    const doc = new Document({
        sections: [{
            properties: {},
            children: paragraphs
        }]
    });

    Packer.toBlob(doc).then((blob) => {
        console.log("DOCX сформирован");
        const downloadLink = document.getElementById("downloadLink");
        downloadLink.style.display = "inline-block";
        downloadLink.href = URL.createObjectURL(blob);
        downloadLink.download = "output.docx";
        downloadLink.textContent = "Скачать Word файл";
    }).catch(error => {
        console.error("Ошибка при формировании DOCX:", error);
    });
}