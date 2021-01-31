// eslint-disable-next-line no-unused-vars
function createWineList() {
  // Global Configurations
  const HEADERS = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetValues(1, 1, 1, -1)[0];
  const TEMPLATE_ID = '12yfe6AowMOBML7pkagvPJT_5M-APyuiSzcCSJo3I_80';
  const FOLDER_ID = '1pTbUMcLlU2q-ZaLLouGSGRiYgdHjlN4U';
  const MAX_LINES = 54;
  const CUVEE_SIZE = 9;
  const COUNTRY_LINES = 21 / CUVEE_SIZE;
  const REGION_LINES = 28 / CUVEE_SIZE;
  const LINES_NEEDED = {
    // A new category should always be placed on a new page.
    category: 0,
    region: REGION_LINES + 2,
    country: COUNTRY_LINES + REGION_LINES + 2,
  };
  let pageLineCounter = 0;
  let readCounter = 0;
  let writeCounter = 0;
  let percentage = 0;
  let getAppendFunction;

  function toast(message) {
    SpreadsheetApp.getActiveSpreadsheet().toast(message);
  }

  function alert(message) {
    SpreadsheetApp.getUi().alert(message);
  }

  function getOutOfStockWines() {
    toast('Loading out of stock wines...');
    function getTrueInventory(row, headers) {
      return parseInt(row[headers.indexOf('Cellar')], 10) + parseInt(row[headers.indexOf('Online Store')], 10);
    }

    function getOutOfStockList(sheet, headers) {
      const outOfStockList = [];
      sheet.forEach((row) => {
        if (getTrueInventory(row, headers) <= 0) {
          outOfStockList.push(row[headers.indexOf('Title')].toString());
        }
      });
      return outOfStockList;
    }

    const sheetUrl = SpreadsheetApp.getUi()
      .prompt('Please enter the URL of the inventory sheet.')
      .getResponseText();
    if (!sheetUrl) {
      alert('URL was not entered. All wines will be considered in stock');
      return [];
    }

    let outOfStockSheet;

    try {
      outOfStockSheet = SpreadsheetApp.openByUrl(sheetUrl);
    } catch (error) {
      throw new Error('URL not found');
    }
    const sheetHeaders = outOfStockSheet.getSheetValues(1, 1, 1, -1)[0];
    const sheetValues = outOfStockSheet.getSheetValues(2, 1, -1, -1);
    return getOutOfStockList(sheetValues, sheetHeaders);
  }

  function getHeaderIndex(headerString) {
    return HEADERS.indexOf(headerString);
  }

  function getFormattedName(wine) {
    const nameCell = wine[getHeaderIndex('Name')];
    const vintage = nameCell.match(/\d{4}|NV/) || '';
    // RegExp: matches all words between double quotations
    let name = nameCell.match(/(?<=").+(?=")/);
    // RegExp: matches all words between parenthesis
    const size = nameCell.match(/(?<=\().+(?=\))/);
    if (size) {
      name += ` (${size})`;
    }

    return `${vintage} ${name}`;
  }

  function getFormattedGrapes(wine) {
    const grapes = wine[getHeaderIndex('Grapes')];
    if (grapes.includes('/')) {
      // RegExp: matches "/word" or "/word word" which is followed
      // either by a "," or the end of line
      return grapes.replace(/\/[\w ]+(?=(,|$))/g, '');
    }
    return grapes;
  }

  function createCuvee(wine) {
    return {
      name: getFormattedName(wine),
      grapes: getFormattedGrapes(wine),
      price: wine[getHeaderIndex('Restaurant Price')],
      macerated: wine[getHeaderIndex('Type')] === 'orange',
    };
  }

  function fixCategoryNames(category) {
    if (category === 'white' || category === 'orange') {
      return 'white & macerated';
    } if (category === 'redwine') {
      return 'red';
    }
    return category;
  }

  function createStackForTrie(wine) {
    return [
      fixCategoryNames(wine[getHeaderIndex('Type')]),
      wine[getHeaderIndex('Country')],
      wine[getHeaderIndex('Region')].toLowerCase(),
      wine[getHeaderIndex('Producer')],
    ];
  }

  const insertIntoTrie = function recursivelyCheckTrieHasBeenInitialisedAndInsertCuvee(
    map, stack, cuvee,
  ) {
    const nodeName = stack.shift();
    if (stack.length === 0) {
      if (map[nodeName] === undefined) {
        map[nodeName] = [];
      }
      map[nodeName].push(cuvee);
    } else {
      if (map[nodeName] === undefined) {
        map[nodeName] = {};
      }
      insertIntoTrie(map[nodeName], stack, cuvee);
    }
  };

  function shouldWineGoOnList(wine, outOfStockWines) {
    const typeIndex = getHeaderIndex('Type');
    if (wine[typeIndex] === 'not-wine'
    || wine[typeIndex] === 'fortified'
    || wine[getHeaderIndex('Hide From Wine List')]
    || wine[typeIndex].length === 0) {
      return false;
    }

    // The website still stores names in between single quotations, when this changes
    // to double quotations, this line can be removed.
    const modifiedName = wine[getHeaderIndex('Name')].split('"').join("'");
    return !outOfStockWines.includes(modifiedName);
  }

  function loadWineIntoMap(wine, map) {
    readCounter += 1;
    const cuvee = createCuvee(wine);
    const stack = createStackForTrie(wine);
    insertIntoTrie(map, stack, cuvee);
  }

  function loadWinesIntoHashMap(outOfStockWines) {
    toast('Reading wines from spreadsheet...');
    const wineMap = {};
    const wineSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetValues(3, 1, -1, -1);
    for (let i = 0; i < wineSheet.length; i++) {
      if (shouldWineGoOnList(wineSheet[i], outOfStockWines)) {
        loadWineIntoMap(wineSheet[i], wineMap);
      }
    }
    return wineMap;
  }

  function appendOnNewPageIfNeeded(part, writing) {
    const { current } = writing;
    if (current[part]) {
      const textToInsert = current[part];
      if (part !== 'region') {
        const append = getAppendFunction(part);
        append(textToInsert, writing);
      }
    }
  }

  function getStackOrder() {
    return ['category', 'country', 'region'];
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
    const paragraph = document.getBody().appendParagraph('');
    paragraph.addPositionedImage(image);
    correctImagePosition(paragraph.getPositionedImages()[0]);
    setFontSizeOfParagraph(18, paragraph);
    pageLineCounter += COUNTRY_LINES;
  }

  function appendLineToDocument(line, document) {
    if (line.getType() === DocumentApp.ElementType.PARAGRAPH) {
      document.appendParagraph(line);
    } else if (line.getType() === DocumentApp.ElementType.TABLE) {
      document.appendTable(line);
    }
  }

  function appendCategory(category, { templates, document }) {
    const template = templates.category(category);
    appendLineToDocument(template, document);
  }

  function formatPrice(price) {
    return parseInt(price, 10);
  }

  function getProgressPercentage() {
    return (writeCounter / readCounter) * 100;
  }

  function updateProgress() {
    writeCounter += 1;
    if (getProgressPercentage() >= percentage + 10) {
      SpreadsheetApp.getActiveSpreadsheet().toast(`Writing to template: ${Math.round(getProgressPercentage())}% completed`);
      percentage = getProgressPercentage();
    }
  }

  function sortByName(a, b) {
    const x = a.name.toLowerCase();
    const y = b.name.toLowerCase();
    if (x < y) {
      return -1;
    } if (y < x) {
      return 1;
    }
    return 0;
  }

  function appendProducer(producer, cuvees, templates, table) {
    const producerRow = templates.producer(producer);
    table.appendTableRow(producerRow);
    cuvees.sort(sortByName);
    cuvees.forEach((cuvee) => {
      const cuveeRow = templates.cuvee(cuvee);
      table.appendTableRow(cuveeRow);
      updateProgress();
    });
  }

  function willProducerExtendToNextPage(producer) {
    return pageLineCounter > MAX_LINES - (producer.length + 1);
  }

  function getLastChild(document) {
    return document.getBody().getChild(document.getBody().getNumChildren() - 1);
  }

  function appendRegion(region, writing, producers) {
    const { templates, document } = writing;
    const producerNames = Object.keys(producers).sort();
    if (willProducerExtendToNextPage(producers[producerNames[0]])) {
      appendPageBreak(writing);
    }
    let table = templates.table();
    const regionRow = templates.region(region);
    table.appendTableRow(regionRow);
    pageLineCounter += REGION_LINES;

    producerNames.forEach((producerName, index) => {
      if (index !== 0 && willProducerExtendToNextPage(producers[producerName])) {
        appendLineToDocument(table, document);
        appendPageBreak(writing);
        table = templates.table();
        const continuedRegion = templates.region(`${region} cont.`);
        pageLineCounter += REGION_LINES;
        table.appendTableRow(continuedRegion);
      }
      pageLineCounter += producers[producerName].length + 1;
      appendProducer(producerName, producers[producerName], templates, table);
    });
    appendLineToDocument(table, document);
    setFontSizeOfParagraph(11, getLastChild(document));
  }

  getAppendFunction = (type) => {
    const appendFunctions = {
      region: appendRegion,
      country: appendCountryImage,
      category: appendCategory,
    };
    return appendFunctions[type];
  };

  function hasEndOfPageBeenReached(dataType) {
    return pageLineCounter > (MAX_LINES - LINES_NEEDED[dataType]);
  }

  function getCountryOrder() {
    return ['France', 'Italy', 'Austria', 'Germany', 'Australia', 'South Africa'];
  }

  function getKeys(type, data) {
    if (type === 'country') {
      return getCountryOrder();
    }
    return Object.keys(data).sort();
  }

  function appendNext(data, dataTypeStack, writing) {
    const { document, current } = writing;
    const dataType = dataTypeStack.shift();
    const keys = getKeys(dataType, data);
    keys.forEach((key) => {
      if (data[key]) {
        if (hasEndOfPageBeenReached(dataType)) {
          appendPageBreak(writing);
        }
        current[dataType] = key;
        const append = getAppendFunction(dataType);

        append(key, writing, data[key]);
      }
      if (dataTypeStack.length > 0 && data[key]) {
        appendNext(data[key], dataTypeStack, writing);
      }
      current[dataType] = undefined;
      if (dataType === 'category') {
        document.appendPageBreak();
        pageLineCounter = 0;
      }
    });
    dataTypeStack.unshift(dataType);
  }

  function loadTemplate() {
    const templateBody = DocumentApp.openById(TEMPLATE_ID).getBody().copy();
    const table = templateBody.getChild(2).asTable();
    const templates = {
      category: (textToInsert) => {
        const template = templateBody.getChild(0).copy().asParagraph().replaceText('{{category}}', textToInsert);

        template.asParagraph().replaceText('{{category_maceration}}', textToInsert === 'white & macerated' ? ' ▴' : '');

        return template;
      },
      region: (textToInsert) => table.getRow(0).copy().replaceText('{{region}}', textToInsert),
      producer: (textToInsert) => table.getRow(1).copy().replaceText('{{producer}}', textToInsert),
      cuvee: ({
        name, grapes, price, macerated,
      }) => {
        const template = table.getRow(2).copy();
        template.replaceText('{{cuvee}}', name);
        template.replaceText('{{grapes}}', grapes);
        template.replaceText('{{price}}', formatPrice(price).toString());
        template.replaceText('{{cuvee_maceration}}', macerated ? ' ▴' : '');
        return template;
      },
      table: () => table.copy().clear(),
    };
    return templates;
  }

  function setTopAndBottomMargins(document) {
    const style = {};
    style[DocumentApp.Attribute.MARGIN_BOTTOM] = 45;
    style[DocumentApp.Attribute.MARGIN_TOP] = 45;
    document.getBody().setAttributes(style);
  }

  function createNewWineListFile() {
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const wineList = DocumentApp.create(`Wine List ${new Date().toDateString()}`);
    folder.addFile(DriveApp.getFileById(wineList.getId()));
    setTopAndBottomMargins(wineList);
    return wineList;
  }

  function getCountryImages() {
    const countries = getCountryOrder();
    const images = {};
    countries.forEach((country) => {
      images[country] = DriveApp.getFilesByName(`${country}.png`)
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

  alert('Please stay on this sheet until the script has completed.');
  const outOfStockWines = getOutOfStockWines();
  const wines = loadWinesIntoHashMap(outOfStockWines);
  writeWinesToTemplate(wines);
  toast('100% completed');
}
