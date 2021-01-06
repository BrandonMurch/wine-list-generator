const headers = Sheets.Spreadsheets.Values.get(
  "1xWWHhLTvCj-OA1MBp7dDFetp4vCY7oup4KP68IVmmUo",
  "Input!A1:AT1"
);
const wineSheet = Sheets.Spreadsheets.Values.get(
  "1xWWHhLTvCj-OA1MBp7dDFetp4vCY7oup4KP68IVmmUo",
  "Input!A3:AT1000"
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
  return headers.values[0].indexOf(headerString);
}

function createCuvee(wine) {
  const name = wine[getHeaderIndex("Name")].match(/(\')(.+)(\')/)[2];
  return {
    name: name,
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
  for (let i = 0; i < wineSheet.values.length; i++) {
    loadWineIntoMapIfIncludedInWineList(wineSheet.values[i], wineMap);
  }
  return wineMap;
}

function appendOnNewPageIfNeeded(part, writing) {
  const { templates, current } = writing;
  if (current[part]) {
    let textToInsert = current[part];
    if (part == "region") {
      textToInsert += " continued";
      appendLineToDocument(
        templates.region.copy().replaceText("{{region}}", textToInsert)
      );
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

function formatPrice(price) {
  return parseInt(price.replace("$", ""));
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
  // if (pageLineCounter > MAX_LINES - (COUNTRY_LINES + REGION_LINES + 2)) {
  //   appendPageBreak(writing)
  // }
  const image = DriveApp.getFilesByName(country + ".png")
    .next()
    .getBlob();
  const paragraph = document.getBody().appendParagraph("");
  paragraph.addPositionedImage(image);
  correctImagePosition(paragraph.getPositionedImages()[0]);
  setFontSizeOfParagraph(29, paragraph);
  pageLineCounter += COUNTRY_LINES;
}

function appendCategory(category, { templates, document }) {
  const template = templates.category
    .copy()
    .replaceText("{{category}}", category);

  template.replaceText(
    "{{category_maceration}}",
    category == "white & macerated" ? " ▴" : ""
  );
  appendLineToDocument(template, document);
}

function appendProducer(producer, cuvees, templates, table) {
  const producerRow = templates.producer
    .copy()
    .replaceText("{{producer}}", producer);
  table.appendTableRow(producerRow);
  cuvees.forEach((cuvee) => {
    const cuveeRow = templates.cuvee.copy();
    cuveeRow.replaceText("{{cuvee}}", cuvee.name);
    cuveeRow.replaceText("{{grapes}}", cuvee.grapes);
    cuveeRow.replaceText("{{price}}", formatPrice(cuvee.price));
    cuveeRow.replaceText("{{cuvee_maceration}}", cuvee.macerated ? " ▴" : "");
    table.appendTableRow(cuveeRow);
    updateProgress();
  });
}

function updateProgress() {
  writeCounter++;
  Logger.log(Math.round((writeCounter / readCounter) * 100) + "% completed");
}

function willProducersFitOnPage(producer) {
  return pageLineCounter > MAX_LINES - (producer.length + 1);
}

function getLastChild(document) {
  return document.getBody().getChild(document.getBody().getNumChildren() - 1);
}

function appendRegion(region, writing, producers) {
  const { templates, document } = writing;
  // if (pageLineCounter > (MAX_LINES - (REGION_LINES + 2))) {
  //   appendPageBreak(writing)
  // }
  const table = templates.table.copy();
  const regionRow = templates.region.copy().replaceText("{{region}}", region);
  table.appendTableRow(regionRow);
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

function appendNext(data, dataTypeStack, writing) {
  const { document, current } = writing;
  const keys = Object.keys(data);
  const dataType = dataTypeStack.shift();
  keys.forEach((key) => {
    if (hasEndOfPageBeenReached(dataType)) {
      appendPageBreak(writing);
    }
    current[dataType] = key;
    const append = getAppendFunction(dataType);
    append(key, writing, data[key]);
    if (dataTypeStack.length > 0) {
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
    category: templateBody.getChild(0),
    region: table.getRow(0),
    producer: table.getRow(1),
    cuvee: table.getRow(2),
    table: table.copy().clear(),
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
  const wineList = DocumentApp.create("Wine List " + new Date().toDateString());
  folder.addFile(DriveApp.getFileById(wineList.getId()));
  setTopAndBottomMargins(wineList);
  return wineList;
}

function writeWinesToTemplate(wines) {
  const writing = {
    templates: loadTemplate(),
    document: createNewWineListFile(),
    current: {},
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
    pageLineCounter += line.getNumChildren();
    document.appendTable(line);
  }
}

function createWineList() {
  const wines = loadWinesIntoHashMap();
  writeWinesToTemplate(wines);
}
