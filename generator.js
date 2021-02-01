// eslint-disable-next-line no-unused-vars
function createWineList() {
  // Global Configurations
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
  let readCounter = 0;

  function getStackOrder() {
    return ['category', 'country', 'region'];
  }

  function toast(message) {
    SpreadsheetApp.getActiveSpreadsheet().toast(message);
  }

  function alert(message) {
    SpreadsheetApp.getUi().alert(message);
  }

  function ProgressTracker(total, getMessage) {
    let progress = 0;
    let lastAnnouncedPercentage = 0;

    function getProgressPercentage() {
      return Math.round((progress / total) * 100);
    }

    function display() {
      const percentage = getProgressPercentage();
      if (percentage >= lastAnnouncedPercentage + 10) {
        toast(getMessage(percentage));
        lastAnnouncedPercentage = percentage;
      }
    }
    return function update(amount) {
      if (amount === undefined) {
        progress += 1;
      } else {
        progress += amount;
      }
      display();
    };
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
  function loadWines(outOfStockWines) {
    const headers = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetValues(1, 1, 1, -1)[0];

    function getHeaderIndex(headerString) {
      return headers.indexOf(headerString);
    }

    function createCuvee(wine) {
      function getFormattedName() {
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

      function getFormattedGrapes() {
        const grapes = wine[getHeaderIndex('Grapes')];
        if (grapes.includes('/')) {
          // RegExp: matches "/word" or "/word word" which is followed
          // either by a "," or the end of line
          return grapes.replace(/\/[\w ]+(?=(,|$))/g, '');
        }
        return grapes;
      }

      readCounter += 1;

      return {
        name: getFormattedName(wine),
        grapes: getFormattedGrapes(wine),
        price: wine[getHeaderIndex('Restaurant Price')],
        macerated: wine[getHeaderIndex('Type')] === 'orange',
      };
    }

    function createStackForTrie(wine) {
      function fixCategoryNames(category) {
        if (category === 'white' || category === 'orange') {
          return 'white & macerated';
        } if (category === 'redwine') {
          return 'red';
        }
        return category;
      }
      return [
        fixCategoryNames(wine[getHeaderIndex('Type')]),
        wine[getHeaderIndex('Country')],
        wine[getHeaderIndex('Region')].toLowerCase(),
        wine[getHeaderIndex('Producer')],
      ];
    }

    const insertWine = function recursivelyCheckTrieHasBeenInitialisedAndInsertCuvee(
      map, stack, cuvee,
    ) {
      const nodeName = stack.shift();
      // Last insert will be a list of cuvees for a producer
      if (stack.length === 0) {
        if (map[nodeName] === undefined) {
          map[nodeName] = [];
        }
        map[nodeName].push(cuvee);
      } else {
        if (map[nodeName] === undefined) {
          map[nodeName] = {};
        }
        insertWine(map[nodeName], stack, cuvee);
      }
    };

    function shouldWineGoOnList(wine) {
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

    function loadWine(wineMap, wine) {
      if (shouldWineGoOnList(wine, outOfStockWines)) {
        insertWine(
          wineMap,
          createCuvee(wine),
          createStackForTrie(wine),
        );
      }
    }

    function displayLoadProgress(current, total) {
      if ((current / total) % 10 === 0) {
        toast(`Reading wines from spreadsheet: ${current / total}% completed`);
      }
    }

    const wines = {};
    const spreadSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetValues(3, 1, -1, -1);
    for (let i = 0; i < spreadSheet.length; i++) {
      loadWine(wines, spreadSheet[i], outOfStockWines);
      displayLoadProgress(i, spreadSheet.length);
    }
    return wines;
  }
  function writeWines(wines) {
    function correctImagePosition(image) {
      image.setLeftOffset(35);
      image.setTopOffset(40);
    }

    function setFontSizeOfParagraph(size, paragraph) {
      const style = {};
      style[DocumentApp.Attribute.FONT_SIZE] = size;
      paragraph.setAttributes(style);
    }

    function willProducerExtendToNextPage(producer, lineCounter) {
      return lineCounter > MAX_LINES - (producer.length + 1);
    }

    const append = (type, {
      templates, document, images, current, ...writing
    }, name, data) => {
      const progress = new ProgressTracker(
        readCounter,
        (percentage) => `Writing to template: ${percentage}% completed`,
      );
      function appendLineToDocument(line) {
        if (line.getType() === DocumentApp.ElementType.PARAGRAPH) {
          document.appendParagraph(line);
        } else if (line.getType() === DocumentApp.ElementType.TABLE) {
          document.appendTable(line);
        }
      }
      function appendPageBreak() {
        document.appendPageBreak();
        writing.lineCounter = 0;
        const order = getStackOrder();
        order.forEach((part) => {
          if (current[part] && part !== 'region') {
            append(part, writing, current[part]);
          }
        });
      }

      function appendCuvee(cuvees, table) {
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

        cuvees
          .sort(sortByName)
          .forEach((cuvee) => {
            const cuveeRow = templates.cuvee(cuvee);
            table.appendTableRow(cuveeRow);
            progress.update();
          });
      }

      function appendProducer(cuvees, producerName, passedInTable, createNewTable) {
        let table = passedInTable;
        if (table.getNumChildren > 1
        && willProducerExtendToNextPage(cuvees, writing.lineCounter)
        ) {
          table = createNewTable();
        }
        writing.lineCounter += cuvees.length + 1;
        const producerRow = templates.producer(producerName);
        table.appendTableRow(producerRow);
        appendCuvee(cuvees, table);
      }

      const appendFunctions = {
        region: (region, producers) => {
          let table = templates.table();
          function createNewTable() {
            appendLineToDocument(table);
            appendPageBreak(writing);
            table = templates.table();
            table.appendTableRow(templates.region(`${region} cont.`));
            writing.lineCounter += REGION_LINES;
          }
          const producerNames = Object.keys(producers).sort();
          if (willProducerExtendToNextPage(producers[producerNames[0]], writing.lineCounter)) {
            appendPageBreak(writing);
          }

          const regionRow = templates.region(region);
          table.appendTableRow(regionRow);
          writing.lineCounter += REGION_LINES;

          producerNames.forEach((producerName) => {
            appendProducer(producers[producerName], producerName, table, createNewTable);
          });
          appendLineToDocument(table);
          setFontSizeOfParagraph(11, document.getLastChild());
        },
        country: (country) => {
          const image = images[country];
          const paragraph = document.getBody().appendParagraph('');
          paragraph.addPositionedImage(image);
          correctImagePosition(paragraph.getPositionedImages()[0]);
          setFontSizeOfParagraph(18, paragraph);
          writing.lineCounter += COUNTRY_LINES;
        },
        category: (category) => {
          appendLineToDocument(
            templates.category(category),
          );
        },
        pageBreak: appendPageBreak,
      };

      return appendFunctions[type](name, data);
    };

    function willEndOfPageBeReached(dataType, lineCounter) {
      return lineCounter > (MAX_LINES - LINES_NEEDED[dataType]);
    }

    function getCountryOrder() {
      return ['France', 'Italy', 'Austria', 'Germany', 'Australia', 'South Africa'];
    }

    function appendNext(data, dataTypeStack, writing) {
      function getKeys(type) {
        if (type === 'country') {
          return getCountryOrder();
        }
        return Object.keys(data).sort();
      }

      if (!data || dataTypeStack.length === 0) {
        return;
      }
      const { document, current } = writing;
      const dataType = dataTypeStack.shift();
      const keys = getKeys(dataType, data);
      keys.forEach((key) => {
        current[dataType] = key;
        if (data[key]) {
          if (willEndOfPageBeReached(dataType, writing.lineCounter)) {
            append('pageBreak', writing);
          }
          append(dataType, writing, key, data[key]);
        }
        appendNext(data[key], dataTypeStack, writing);
        current[dataType] = undefined;
        if (dataType === 'category') {
          document.appendPageBreak();
          writing.lineCounter = 0;
        }
      });
      dataTypeStack.unshift(dataType);
    }

    function loadTemplate() {
      const templateBody = DocumentApp.openById(TEMPLATE_ID).getBody().copy();
      const table = templateBody.getChild(2).asTable();
      return {
        category: (textToInsert) => {
          const template = templateBody
            .getChild(0)
            .copy().asParagraph()
            .replaceText('{{category}}', textToInsert);

          template.asParagraph()
            .replaceText('{{category_maceration}}', textToInsert === 'white & macerated' ? ' ▴' : '');

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
          template.replaceText('{{price}}', price);
          template.replaceText('{{cuvee_maceration}}', macerated ? ' ▴' : '');
          return template;
        },
        table: () => table.copy().clear(),
      };
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
      wineList.getLastChild = () => (
        wineList.getBody().getChild(
          wineList.getBody().getNumChildren() - 1,
        )
      );
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

    return () => {
      const writing = {
        templates: loadTemplate(),
        document: createNewWineListFile(),
        current: {},
        images: getCountryImages(),
        lineCounter: 0,
      };
      const stackOrder = getStackOrder();
      appendNext(wines, stackOrder, writing);
      writing.document.getBody().appendPageBreak();
      writing.document.saveAndClose();
    };
  }

  alert('Please stay on this sheet until the script has completed.');
  const outOfStockWines = getOutOfStockWines();
  const wines = loadWines(outOfStockWines);
  writeWines(wines);
  toast('100% completed');
}
