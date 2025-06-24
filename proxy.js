import puppeteer from "puppeteer-core";
import xlsx from "xlsx";
import fs from "fs";

// ==================== LOGGER ====================
class Logger {
  static log(message, level = "INFO") {
    const timestamp = new Date().toISOString();
    const logLine = `[${timestamp}] [${level}] ${message}\n`;
    fs.appendFileSync("log.txt", logLine);
    console.log(`[${level}] ${message}`);
  }

  static info(message) {
    this.log(message, "INFO");
  }

  static error(message) {
    this.log(message, "ERROR");
  }

  static success(message) {
    this.log(message, "SUCCESS");
  }

  static warning(message) {
    this.log(message, "WARNING");
  }

  static debug(message) {
    this.log(message, "DEBUG");
  }
}

// ==================== CONFIGURATION ====================
const CONFIG = {
  CHROME_PATH: "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
  BASE_URL: "https://congbothongtin.ssc.gov.vn/faces/NewsSearch",
  TIMEOUTS: {
    PAGE_LOAD: 10000,
    ELEMENT_WAIT: 10000,
    TAB_SWITCH: 2000,
    SEARCH_DELAY: 1000,
    NAVIGATION_DELAY: 5000,
  },
  SELECTORS: {
    SEARCH_INPUT: 'input[id$="pt9:it8112::content"]',
    SEARCH_BUTTON: 'div[id="pt9:b1"] a',
    MAX_RECORD: 'span[id$="pt9:it4"]',
    PAGINATION_LINKS: "a.x14f",
    RESULT_LINKS: "a.xgl",
    YEAR_SPAN: 'span[id="pt2:tt1:2:lookupValueId::content"]',
    QUARTER_SPAN: 'span[id="pt2:tt1:3:lookupValueId::content"]',
    MAIN_TABLE: 'div[id="pt2:t2::db"]',
    KQKD_TABLE: 'div[id="pt2:t3::db"]',
    LCTT_TABLE: 'div[id="pt2:t6::db"]',
  },
  TABS: {
    KQKD: "pt2:KQKD::disAcr",
    LCTT_GT: "pt2:LCTT-GT::disAcr",
  },
  EXCEL: {
    INPUT_FILE: "codes.xlsx",
    OUTPUT_FILE: "BCTC.xlsx",
    COLUMN_E_INDEX: 4,
    START_COLUMN: 3,
    START_ROW: 3,
  },
};

function getColumnEValues(filename) {
  const workbook = xlsx.readFile(filename);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const range = xlsx.utils.decode_range(sheet["!ref"]);
  const values = [];
  for (let row = range.s.r + 1; row <= range.e.r; ++row) {
    // +1 to skip header
    const cellAddress = { c: 4, r: row }; // Column E is index 4
    const cellRef = xlsx.utils.encode_cell(cellAddress);
    const cell = sheet[cellRef];
    if (cell && cell.v !== undefined && cell.v !== null && cell.v !== "") {
      values.push(cell.v);
    }
  }
  return values;
}

const input = getColumnEValues("codes.xlsx");

async function typeAndSearch(page, selector, value) {
  try {
    await page.waitForSelector(selector);
    await page.type(selector, value);
  } catch (err) {
    Logger.error(`❌ Error in typeAndSearch: ${err.message}`);
  }
}

async function clickSearchButton(page) {
  try {
    await page.evaluate(() => {
      const searchDiv = document.querySelector('div[id="pt9:b1"]');
      const searchLink = searchDiv?.querySelector("a");
      if (searchLink) {
        searchLink.click();
      }
    });
  } catch (err) {
    Logger.error(`❌ Error in clickSearchButton: ${err.message}`);
  }
}

async function getMaxRecord(page) {
  try {
    const totalCount = await page.evaluate(() => {
      const span = document.querySelector('span[id$="pt9:it4"]');
      if (!span) return 0;
      const text = span.textContent || "";
      const match = text.match(/\d+/);
      return match ? parseInt(match[0], 10) : 0;
    });
    return totalCount;
  } catch (err) {
    Logger.error(`❌ Error in getMaxRecord: ${err.message}`);
    return 0;
  }
}

async function navigateToPage(page, pageNumber) {
  try {
    await page.waitForSelector("a.x14f");
    await page.evaluate((pageNumber) => {
      const anchors = Array.from(document.querySelectorAll("a.x14f"));
      const target = anchors.find(
        (a) => a.textContent.trim() === String(pageNumber)
      );
      if (target) {
        target.click();
        return true;
      }
      return false;
    }, pageNumber);
  } catch (err) {
    Logger.error(`❌ Error in navigateToPage: ${err.message}`);
  }
}

async function extractData(page, year, quarter) {
  try {
    const extractedData = await page.evaluate(
      ({ year, quarter }) => {
        const yearQueried = document
          .querySelector('span[id="pt2:tt1:2:lookupValueId::content"]')
          ?.textContent?.trim();
        const quarterQueried = document
          .querySelector('span[id="pt2:tt1:3:lookupValueId::content"]')
          ?.textContent?.trim();
        if (
          yearQueried === String(year) &&
          quarterQueried === String(quarter)
        ) {
          const tableDiv = document.querySelector('div[id="pt2:t2::db"]');
          if (!tableDiv) {
            return [];
          }
          const tbody = tableDiv.querySelector("tbody");
          if (!tbody) {
            return [];
          }
          const rows = Array.from(tbody.querySelectorAll("tr")).slice(1);
          return rows.map((row) => {
            const cells = row.querySelectorAll("td");
            const td5Span =
              cells[3]?.querySelector("span")?.textContent?.trim() || "-";
            const td6Span =
              cells[4]?.querySelector("span")?.textContent?.trim() || "-";
            return {
              start: td5Span === "0" ? "-" : td5Span,
              end: td6Span === "0" ? "-" : td6Span,
            };
          });
        }
        return [];
      },
      { year, quarter }
    );
    return extractedData;
  } catch (err) {
    Logger.error(`❌ Error in extractData: ${err.message}`);
    return [];
  }
}

async function extractTabData(page, tabId) {
  try {
    // Use a different approach to click the tab - find by text content instead of ID
    await page.evaluate((tabId) => {
      let tabElement;
      if (tabId === "pt2:KQKD::disAcr") {
        // Find KQKD tab by looking for elements with "KQKD" text
        const allElements = document.querySelectorAll("*");
        for (const element of allElements) {
          if (
            element.textContent &&
            element.textContent.includes("KQKD") &&
            element.tagName === "A" &&
            element.getAttribute("role") === "tab"
          ) {
            tabElement = element;
            break;
          }
        }
      } else if (tabId === "pt2:LCTT-GT::disAcr") {
        // Find LCTT-GT tab by looking for elements with "LCTT" text
        tabElement = document.querySelector(`a[id="${tabId}"]`);
      }

      if (tabElement) {
        tabElement.click();
      }
    }, tabId);

    await new Promise((resolve) => setTimeout(resolve, 2000)); // Increased wait time

    // Wait for tab content to be visible
    if (tabId === "pt2:LCTT-GT::disAcr") {
      await page.waitForSelector('div[id="pt2:LCTT-GT::body"]', {
        timeout: 5000,
      });
    } else if (tabId === "pt2:KQKD::disAcr") {
      await page.waitForSelector('div[id="pt2:KQKD"]', { timeout: 5000 });
    }

    return await page.evaluate((tabId) => {
      // For KQKD tab, look for table inside the KQKD div
      let tableDiv;
      if (tabId === "pt2:KQKD::disAcr") {
        const kqkdDiv = document.querySelector('div[id="pt2:KQKD"]');
        tableDiv = kqkdDiv?.querySelector('div[id="pt2:t3::db"]');

        // For KQKD, handle the table structure
        if (tableDiv) {
          const tbody = tableDiv.querySelector("tbody");
          if (tbody) {
            const rows = Array.from(tbody.querySelectorAll("tr"));
            return rows.map((row) => {
              const cells = row.querySelectorAll("td");
              const col3 =
                cells[3]?.querySelector("span")?.textContent?.trim() || "-";
              const col4 =
                cells[4]?.querySelector("span")?.textContent?.trim() || "-";
              return [col3 === "0" ? "-" : col3, col4 === "0" ? "-" : col4];
            });
          }
        }
        return [];
      } else if (tabId === "pt2:LCTT-GT::disAcr") {
        // For LCTT-GT tab, look for table inside the LCTT-GT div
        const lcttBodyDiv = document.querySelector(
          'div[id="pt2:LCTT-GT::body"]'
        );
        const lcttDiv = lcttBodyDiv?.querySelector('div[id="pt2:LCTT-GT"]');
        tableDiv = lcttDiv?.querySelector('div[id="pt2:t6::db"]');

        // For LCTT-GT, the table structure is different - data is directly in tbody
        if (tableDiv) {
          const tbody = tableDiv.querySelector("tbody");
          if (tbody) {
            const rows = Array.from(tbody.querySelectorAll("tr"));
            return rows.map((row) => {
              const cells = row.querySelectorAll("td");
              // For LCTT-GT, we need columns 3 and 4 (index 3 and 4)
              const col3 =
                cells[3]?.querySelector("span")?.textContent?.trim() || "-";
              const col4 =
                cells[4]?.querySelector("span")?.textContent?.trim() || "-";
              return [col3 === "0" ? "-" : col3, col4 === "0" ? "-" : col4];
            });
          }
        }
        return [];
      } else {
        // For other tabs, use the original selector
        tableDiv = document.querySelector('div[id="pt2:t2::db"]');
      }
      if (!tableDiv) return [];
      const tbody = tableDiv.querySelector("tbody");
      if (!tbody) return [];
      let rows;
      // For other tabs, skip the first row (header)
      rows = Array.from(tbody.querySelectorAll("tr")).slice(1);
      Logger.info(
        `🚀 ~ returnawaitpage.evaluate ~ rows: ${JSON.stringify(rows)}`
      );
      return rows.map((row) => {
        const cells = row.querySelectorAll("td");
        const col3 =
          cells[3]?.querySelector("span")?.textContent?.trim() || "-";
        const col4 =
          cells[4]?.querySelector("span")?.textContent?.trim() || "-";
        return [col3 === "0" ? "-" : col3, col4 === "0" ? "-" : col4];
      });
    }, tabId);
  } catch (err) {
    Logger.error(`❌ Error in extractTabData for tab ${tabId}: ${err.message}`);
    return [];
  }
}

async function processCompany(page, code, year, quarter) {
  Logger.info(`PROCESSING CODE: ${code}`);
  try {
    await page.goto("https://congbothongtin.ssc.gov.vn/faces/NewsSearch", {
      waitUntil: "domcontentloaded",
    });
  } catch (err) {
    Logger.error(`❌ Error in page.goto for code ${code}: ${err.message}`);
    return { foundData: false, extractedValues: [] };
  }
  let maxRecord = 10;
  let currentPage = 1;
  try {
    await page.waitForSelector("a.xgl", { timeout: 10000 });
  } catch (err) {
    Logger.error(
      `❌ Error waiting for selector a.xgl for code ${code}: ${err.message}`
    );
    return { foundData: false, extractedValues: [] };
  }
  let foundData = false;
  let extractedValues = [];
  for (let i = 0; i < maxRecord || currentPage <= 4; i++) {
    try {
      if (i === 0 || (i + 1) % 15 === 0) {
        if (i !== 0) {
          const hasNextPage = await page.evaluate((pageNumber) => {
            const anchors = Array.from(document.querySelectorAll("a.x14f"));
            return anchors.some(
              (a) => a.textContent.trim() === String(pageNumber)
            );
          }, i / 15 + 1);
          if (!hasNextPage) {
            Logger.info("No more pages available, moving to next code.");
            break;
          }
        }
        i / 15 >= 1 && (await navigateToPage(page, i / 15 + 1));
        await typeAndSearch(page, 'input[id$="pt9:it8112::content"]', code);
        await new Promise((resolve) => setTimeout(resolve, 1000));
        await clickSearchButton(page);
        await new Promise((resolve) => setTimeout(resolve, 1000));
        maxRecord = await getMaxRecord(page);
      }
      Logger.info(`🚀 ~ processCompany ~ maxRecord: ${maxRecord}`);
      await new Promise((resolve) => setTimeout(resolve, 1000));
      await page.waitForSelector("a.xgl", { timeout: 10000 });
      const links = await page.$$("a.xgl");
      const link = links[i % 15];
      if (!link) continue;
      await link.click();
      await new Promise((resolve) => setTimeout(resolve, 5000));
      const extractedData = await extractData(page, year, quarter);
      if (extractedData?.length > 0) {
        extractedValues = extractedData.map((item) => [
          item.start ?? "",
          item.end ?? "",
        ]);
        foundData = true;
        Logger.info(`✅ Page ${i + 1} has data, moving to next code.`);
        break;
      } else {
        Logger.warning(`⚠️ Page ${i + 1} has no data, skipping`);
      }
      await page.goBack({ waitUntil: "domcontentloaded" });
      await typeAndSearch(page, 'input[id$="pt9:it8112::content"]', code);
      await clickSearchButton(page);
      await page.waitForSelector("a.xgl");
    } catch (err) {
      Logger.error(
        `❌ Error in processCompany loop for code ${code}: ${err.message}`
      );
      continue;
    }
  }
  // Write to file immediately if foundData
  if (foundData) {
    const filename = "BCTC.xlsx";
    const workbookFile = xlsx.readFile(filename);
    const sheet1 = workbookFile.Sheets[workbookFile.SheetNames[0]];
    const sheet2 = workbookFile.Sheets[workbookFile.SheetNames[1]];
    const sheet3 = workbookFile.Sheets[workbookFile.SheetNames[2]];

    if (!sheet1) {
      Logger.error(`❌ Sheet 1 not found in ${filename}`);
      return { foundData, extractedValues };
    }
    let range = xlsx.utils.decode_range(sheet1["!ref"]);
    let col = 3;
    while (true) {
      const cellRef = xlsx.utils.encode_cell({ c: col, r: 1 });
      if (!sheet1[cellRef]) break;
      col++;
    }
    const headerCellStart = xlsx.utils.encode_cell({ c: col, r: 1 });
    const headerCellEnd = xlsx.utils.encode_cell({ c: col + 1, r: 1 });
    sheet1[headerCellStart] = { t: "s", v: code + " cuoi ky" };
    sheet1[headerCellEnd] = { t: "s", v: code + " dau ky" };
    for (let i = 0; i < extractedValues.length; i++) {
      for (let j = 0; j < extractedValues[i].length; j++) {
        const cellRef = xlsx.utils.encode_cell({ c: col + j, r: i + 3 });
        sheet1[cellRef] = { t: "s", v: extractedValues[i][j] };
      }
    }
    if (col + extractedValues[0].length - 1 > range.e.c) {
      range.e.c = col + extractedValues[0].length - 1;
      sheet1["!ref"] = xlsx.utils.encode_range(range);
    }

    if (sheet2) {
      const kqkdData = await extractTabData(page, "pt2:KQKD::disAcr");
      let range2 = xlsx.utils.decode_range(sheet2["!ref"]);
      let col2 = 3;
      while (true) {
        const cellRef = xlsx.utils.encode_cell({ c: col2, r: 1 });
        if (!sheet2[cellRef]) break;
        col2++;
      }
      const headerCellStart2 = xlsx.utils.encode_cell({ c: col2, r: 1 });
      const headerCellEnd2 = xlsx.utils.encode_cell({ c: col2 + 1, r: 1 });
      sheet2[headerCellStart2] = { t: "s", v: code + " cuoi ky" };
      sheet2[headerCellEnd2] = { t: "s", v: code + " dau ky" };
      for (let i = 0; i < kqkdData.length; i++) {
        for (let j = 0; j < kqkdData[i].length; j++) {
          const cellRef = xlsx.utils.encode_cell({ c: col2 + j, r: i + 3 });
          sheet2[cellRef] = { t: "s", v: kqkdData[i][j] };
        }
      }
      if (col2 + 1 > range2.e.c) {
        range2.e.c = col2 + 1;
        sheet2["!ref"] = xlsx.utils.encode_range(range2);
      }
    }

    if (sheet3) {
      const lcttData = await extractTabData(page, "pt2:LCTT-GT::disAcr");
      let range3 = xlsx.utils.decode_range(sheet3["!ref"]);
      let col3 = 3;
      while (true) {
        const cellRef = xlsx.utils.encode_cell({ c: col3, r: 1 });
        if (!sheet3[cellRef]) break;
        col3++;
      }
      const headerCellStart3 = xlsx.utils.encode_cell({ c: col3, r: 1 });
      const headerCellEnd3 = xlsx.utils.encode_cell({ c: col3 + 1, r: 1 });
      sheet3[headerCellStart3] = { t: "s", v: code + " cuoi ky" };
      sheet3[headerCellEnd3] = { t: "s", v: code + " dau ky" };
      for (let i = 0; i < lcttData.length; i++) {
        for (let j = 0; j < lcttData[i].length; j++) {
          const cellRef = xlsx.utils.encode_cell({ c: col3 + j, r: i + 3 });
          sheet3[cellRef] = { t: "s", v: lcttData[i][j] };
        }
      }
      if (col3 + 1 > range3.e.c) {
        range3.e.c = col3 + 1;
        sheet3["!ref"] = xlsx.utils.encode_range(range3);
      }
    }
    try {
      xlsx.writeFile(workbookFile, filename);
      Logger.info(
        `✅ Đã ghi file ${filename} cho mã ${code} (sheet: ${workbookFile.SheetNames[0]}, ${workbookFile.SheetNames[1]}, ${workbookFile.SheetNames[2]})`
      );
    } catch (err) {
      Logger.error(
        `❌ Error writing file ${filename} for code ${code}: ${err.message}`
      );
    }
  }
  return { foundData, extractedValues };
}

async function main(inputList, year, quarter) {
  const browser = await puppeteer.launch({
    executablePath: CONFIG.CHROME_PATH,
    defaultViewport: null,
  });

  const page = await browser.newPage();

  let checked = 0;
  let found = 0;
  let notFound = 0;

  for (const inputItem of inputList) {
    const code = inputItem;
    checked++;
    const { foundData } = await processCompany(page, code, year, quarter);
    if (foundData) found++;
    else notFound++;
    Logger.info(
      `PROGRESS: Checked ${checked}/${inputList.length}. Found: ${found}. Not found: ${notFound}.`
    );
  }

  Logger.info(
    `SUMMARY: Checked ${checked} codes. Found data: ${found}. Not found: ${notFound}. Total: ${inputList.length}`
  );
  await browser.close();
}

main(input, 2024, 4);
// main(input.slice(1), 2024, 4);
// main(input.slice(0, 10), 2024, 4);
