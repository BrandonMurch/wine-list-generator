function createWineList() {
  const headers = SpreadsheetApp.getActiveSpreadsheet().getSheetValues(
    1,
    1,
    1,
    -1
  )[0];
  const wineSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetValues(
    3,
    1,
    -1,
    -1
  );
  const templateId = "12yfe6AowMOBML7pkagvPJT_5M-APyuiSzcCSJo3I_80";
  const folderId = "1pTbUMcLlU2q-ZaLLouGSGRiYgdHjlN4U";
  const MAX_LINES = 60;
  // Change these numbers if the font size changes.
  const CUVEE_SIZE = 9;
  const COUNTRY_LINES = 30 / CUVEE_SIZE;
  const REGION_LINES = 27 / CUVEE_SIZE;
  const LINES_NEEDED = {
    // A new category should always be placed on a new page.
    category: 0,
    region: REGION_LINES + 2,
    country: COUNTRY_LINES + REGION_LINES + 2,
  };
  let pageLineCounter = 0;
  let readCounter = 0;
  let writeCounter = 0;

  function getHeaderIndex(headerString) {
    return headers.indexOf(headerString);
  }

  function createCuvee(wine) {
    const nameCell = wine[getHeaderIndex("Name")];
    const vintage = nameCell.match(/\d{4}/) || "";
    let name = nameCell.match(/(\')(.+)(\')/)[2];
    const size = nameCell.match(/(\()(.+)(\))/);
    if (size) {
      name += " (" + size[2] + ")";
    }
    return {
      name: vintage + " " + name,
      grapes: wine[getHeaderIndex("Grapes")],
      price: wine[getHeaderIndex("Restaurant Price")],
      macerated: wine[getHeaderIndex("Type")] == "orange",
    };
  }

  function fixCategoryNames(category) {
    if (category == "white" || category == "orange") {
      return "white & macerated";
    } else if (category == "redwine") {
      return "red";
    }
    return category;
  }

  function createStackForTrie(wine) {
    return [
      fixCategoryNames(wine[getHeaderIndex("Type")]),
      wine[getHeaderIndex("Country")],
      wine[getHeaderIndex("Region")].toLowerCase(),
      wine[getHeaderIndex("Producer")],
    ];
  }

  const insertIntoTrie = function recursivelyCheckTrieHasBeenInitialisedAndInsertCuvee(
    map,
    stack,
    cuvee
  ) {
    let nodeName = stack.shift();
    if (stack.length === 0) {
      if (map[nodeName] === undefined) {
        map[nodeName] = [];
      }
      map[nodeName].push(cuvee);
    } else {
      if (map[nodeName] == undefined) {
        map[nodeName] = {};
      }
      insertIntoTrie(map[nodeName], stack, cuvee);
    }
  };

  function loadWineIntoMapIfIncludedInWineList(wine, map) {
    if (wine[getHeaderIndex("Wine List")]) {
      readCounter++;
      const cuvee = createCuvee(wine);
      const stack = createStackForTrie(wine);
      insertIntoTrie(map, stack, cuvee);
    }
  }

  function loadWinesIntoHashMap() {
    const wineMap = {};
    for (let i = 0; i < wineSheet.length; i++) {
      loadWineIntoMapIfIncludedInWineList(wineSheet[i], wineMap);
    }
    return wineMap;
  }

  function appendOnNewPageIfNeeded(part, writing) {
    const { templates, current } = writing;
    if (current[part]) {
      let textToInsert = current[part];
      if (part == "region") {
        textToInsert += " continued";
        appendLineToDocument(templates.region(textToInsert));
      } else {
        const append = getAppendFunction(part);
        append(textToInsert, writing);
      }
    }
  }

  function getStackOrder() {
    return ["category", "country", "region"];
  }

  function appendPageBreak(writing) {
    writing.document.appendPageBreak();
    pageLineCounter = 0;
    const order = getStackOrder();
    order.forEach((part) => {
      appendOnNewPageIfNeeded(part, writing);
    });
  }

  function correctImagePosition(image) {
    image.setLeftOffset(35);
    image.setTopOffset(30);
  }

  function setFontSizeOfParagraph(size, paragraph) {
    const style = {};
    style[DocumentApp.Attribute.FONT_SIZE] = size;
    paragraph.setAttributes(style);
  }

  function appendCountryImage(country, writing) {
    const { document } = writing;
    const image = writing.images[country];
    const paragraph = document.getBody().appendParagraph("");
    paragraph.addPositionedImage(image);
    correctImagePosition(paragraph.getPositionedImages()[0]);
    setFontSizeOfParagraph(29, paragraph);
    pageLineCounter += COUNTRY_LINES;
  }

  function appendCategory(category, { templates, document }) {
    const template = templates.category(category);
    appendLineToDocument(template, document);
  }

  function formatPrice(price) {
    SpreadsheetApp.getActiveSpreadsheet().toast(price);
    return parseInt(price);
  }

  function updateProgress() {
    writeCounter++;
    SpreadsheetApp.getActiveSpreadsheet().toast(
      Math.round((writeCounter / readCounter) * 100) + "% completed"
    );
  }

  function appendProducer(producer, cuvees, templates, table) {
    const producerRow = templates.producer(producer);
    table.appendTableRow(producerRow);
    pageLineCounter++;
    cuvees.forEach((cuvee) => {
      const cuveeRow = templates.cuvee(cuvee);
      table.appendTableRow(cuveeRow);
      pageLineCounter++;
      updateProgress();
    });
  }

  function willProducersFitOnPage(producer) {
    return pageLineCounter > MAX_LINES - (producer.length + 1);
  }

  function getLastChild(document) {
    return document.getBody().getChild(document.getBody().getNumChildren() - 1);
  }

  function appendRegion(region, writing, producers) {
    const { templates, document } = writing;
    const table = templates.table();
    const regionRow = templates.region(region);
    table.appendTableRow(regionRow);
    pageLineCounter++;
    const producerNames = Object.keys(producers);
    producerNames.forEach((producerName) => {
      if (willProducersFitOnPage(producers[producerName])) {
        appendPageBreak(writing);
      }
      appendProducer(producerName, producers[producerName], templates, table);
    });
    appendLineToDocument(table, document);
    setFontSizeOfParagraph(11, getLastChild(document));
    pageLineCounter += REGION_LINES;
  }

  function getAppendFunction(type) {
    const appendFunctions = {
      region: appendRegion,
      country: appendCountryImage,
      category: appendCategory,
    };
    return appendFunctions[type];
  }

  function hasEndOfPageBeenReached(dataType) {
    return pageLineCounter > MAX_LINES - LINES_NEEDED[dataType];
  }

  function getCountryOrder() {
    return [
      "France",
      "Italy",
      "Austria",
      "Germany",
      "South Africa",
      "Australia",
    ];
  }

  function getKeys(type, data) {
    if (type == "country") {
      return getCountryOrder();
    } else {
      return Object.keys(data).sort();
    }
  }

  function appendNext(data, dataTypeStack, writing) {
    const { document, current } = writing;
    const dataType = dataTypeStack.shift();
    const keys = getKeys(dataType, data);
    keys.forEach((key) => {
      if (hasEndOfPageBeenReached(dataType)) {
        appendPageBreak(writing);
      }
      current[dataType] = key;
      const append = getAppendFunction(dataType);
      append(key, writing, data[key]);
      if (dataTypeStack.length > 0 && data[key]) {
        appendNext(data[key], dataTypeStack, writing);
      }
      current[dataType] = undefined;
      if (dataType === "category") {
        document.appendPageBreak();
        pageLineCounter = 0;
      }
    });
    dataTypeStack.unshift(dataType);
  }

  function loadTemplate() {
    const templateBody = DocumentApp.openById(templateId).getBody().copy();
    const table = templateBody.getChild(2).asTable();
    const template = {
      category: (textToInsert) => {
        const template = templateBody
          .getChild(0)
          .copy()
          .replaceText("{{category}}", textToInsert);

        template.replaceText(
          "{{category_maceration}}",
          textToInsert == "white & macerated" ? " ▴" : ""
        );

        return template;
      },
      region: (textToInsert) => {
        return table.getRow(0).copy().replaceText("{{region}}", textToInsert);
      },
      producer: (textToInsert) => {
        return table.getRow(1).copy().replaceText("{{producer}}", textToInsert);
      },
      cuvee: ({ name, grapes, price, macerated }) => {
        const template = table.getRow(2).copy();
        template.replaceText("{{cuvee}}", name);
        template.replaceText("{{grapes}}", grapes);
        template.replaceText("{{price}}", formatPrice(price));
        template.replaceText("{{cuvee_maceration}}", macerated ? " ▴" : "");
        return template;
      },
      table: () => {
        return table.copy().clear();
      },
    };
    return template;
  }

  function setTopAndBottomMargins(document) {
    const style = {};
    style[DocumentApp.Attribute.MARGIN_BOTTOM] = 45;
    style[DocumentApp.Attribute.MARGIN_TOP] = 45;
    document.getBody().setAttributes(style);
  }

  function createNewWineListFile() {
    const folder = DriveApp.getFolderById(folderId);
    const wineList = DocumentApp.create(
      "Wine List " + new Date().toDateString()
    );
    folder.addFile(DriveApp.getFileById(wineList.getId()));
    setTopAndBottomMargins(wineList);
    return wineList;
  }

  function getCountryImages() {
    const countries = getCountryOrder();
    const images = {};
    countries.forEach((country) => {
      images[country] = DriveApp.getFilesByName(country + ".png")
        .next()
        .getBlob();
    });
    return images;
  }

  function writeWinesToTemplate(wines) {
    const writing = {
      templates: loadTemplate(),
      document: createNewWineListFile(),
      current: {},
      images: getCountryImages(),
    };
    const stackOrder = getStackOrder();
    appendNext(wines, stackOrder, writing);
    writing.document.getBody().appendPageBreak();
    writing.document.saveAndClose();
  }

  function appendLineToDocument(line, document) {
    if (line.getType() == "PARAGRAPH") {
      pageLineCounter += 1;
      document.appendParagraph(line);
    } else if (line.getType() == "TABLE") {
      pageLineCounter++;
      document.appendTable(line);
    }
  }

  function getOutOfStockWines() {
    const promptResponse = SpreadsheetApp.getUi().prompt(
      "Please enter the url of the out of stock items"
    );
    const outOfStockId = SpreadsheetApp.openByUrl(
      promptResponse.getResponseText()
    ).getId();
    return Sheets.Spreadsheets.Values.get(outOfStockId, "A:Z");
  }

  const outOfStockWines = getOutOfStockWines();
  const wines = loadWinesIntoHashMap();
  writeWinesToTemplate(wines);
}
