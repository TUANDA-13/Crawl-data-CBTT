import puppeteer from "puppeteer-core";
import xlsx from "xlsx";
import fs from "fs";
// import axios from "axios";
// import fs from "fs";

const part = 1;

// const getCompanyCodesFromDrive = async (fileId ="1w_9FFeca7bWuuZbXAGSttshK11FU2zBp") => {
// const downloadUrl = `https://drive.google.com/uc?export=download&id=${fileId}`;
//   try {
//     const response = await axios.get(downloadUrl, {
//       responseType: "arraybuffer",
//     });

//     // Ghi tạm file ra để đọc (hoặc đọc từ buffer)
//     fs.writeFileSync("temp.xlsx", response.data);

//     const workbook = xlsx.read(response.data, { type: "buffer" });
//     const sheetName = workbook.SheetNames[0];
//     const sheet = workbook.Sheets[sheetName];

//     const rows = xlsx.utils.sheet_to_json(sheet);

//     // Giả sử cột có tên là "Mã"
//     const codes = rows.map((row) => row["Mã chứng khoán"]).filter(Boolean);

//     console.log("✅ Mã công ty:", codes);
//     return codes;
//   } catch (error) {
//     console.error("❌ Lỗi khi tải file từ Google Drive:", error.message);
//   }
// };

// await getCompanyCodesFromDrive();

// Read temp.xlsx and get column E values
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

const input = getColumnEValues("temp.xlsx");

const workbook = xlsx.utils.book_new();

// Helper: Wait for selector and type value
async function typeAndSearch(page, selector, value) {
  await page.waitForSelector(selector);
  await page.type(selector, value);
}

// Helper: Click search button
async function clickSearchButton(page) {
  await page.evaluate(() => {
    const searchDiv = document.querySelector('div[id="pt9:b1"]');
    const searchLink = searchDiv?.querySelector("a");
    if (searchLink) {
      searchLink.click();
    }
  });
}

// Helper: Get max record count
async function getMaxRecord(page) {
  const totalCount = await page.evaluate(() => {
    const span = document.querySelector('span[id$="pt9:it4"]');
    if (!span) return 0;

    const text = span.textContent || "";
    const match = text.match(/\d+/); // tìm số đầu tiên trong chuỗi

    return match ? parseInt(match[0], 10) : 0;
  });

  return totalCount;
}

// Helper: Navigate to a specific page
async function navigateToPage(page, pageNumber) {
  await page.waitForSelector("a.x14f"); // đảm bảo các thẻ a đã xuất hiện

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
}

// Helper: Extract data from detail page
async function extractData(page, year, quarter) {
  const extractedData = await page.evaluate(
    ({ year, quarter }) => {
      const yearQueried = document
        .querySelector('span[id="pt2:tt1:2:lookupValueId::content"]')
        ?.textContent?.trim();

      const quarterQueried = document
        .querySelector('span[id="pt2:tt1:3:lookupValueId::content"]')
        ?.textContent?.trim();

      if (yearQueried === String(year) && quarterQueried === String(quarter)) {
        const tableDiv = document.querySelector('div[id="pt2:t2::db"]');
        if (!tableDiv) {
          return [];
        }
        const tbody = tableDiv.querySelector("tbody");
        if (!tbody) {
          return [];
        }
        const rows = Array.from(tbody.querySelectorAll("tr")).slice(1);
        if (rows.length === 0) {
        }
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
}

// Logging system: write log to file and console
function logMessage(message) {
  const timestamp = new Date().toISOString();
  const logLine = `[${timestamp}] ${message}\n`;
  fs.appendFileSync("log.txt", logLine);
  console.log(message);
}

// Main logic for one company
async function processCompany(page, code, year, quarter) {
  logMessage(`PROCESSING CODE: ${code}`);
  await page.goto("https://congbothongtin.ssc.gov.vn/faces/NewsSearch", {
    waitUntil: "domcontentloaded",
  });
  let maxRecord = 10;
  let currentPage = 1;

  await page.waitForSelector("a.xgl", { timeout: 10000 });
  let foundData = false;
  let extractedValues = [];
  for (let i = 0; i < maxRecord || currentPage <= 4; i++) {
    if (i === 0 || (i + 1) % 15 === 0) {
      if (i !== 0) {
        const hasNextPage = await page.evaluate((pageNumber) => {
          const anchors = Array.from(document.querySelectorAll("a.x14f"));
          return anchors.some(
            (a) => a.textContent.trim() === String(pageNumber)
          );
        }, i / 15 + 1);
        if (!hasNextPage) {
          logMessage("No more pages available, moving to next code.");
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

    logMessage(`🚀 ~ processCompany ~ maxRecord: ${maxRecord}`);
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
      logMessage(`✅ Page ${i + 1} has data, moving to next code.`);
      break;
    } else {
      logMessage(`⚠️ Page ${i + 1} has no data, skipping`);
    }

    await page.goBack({ waitUntil: "domcontentloaded" });
    await typeAndSearch(page, 'input[id$="pt9:it8112::content"]', code);
    await clickSearchButton(page);
    await page.waitForSelector("a.xgl");
  }

  // Write to file immediately if foundData
  if (foundData) {
    const filename = "BCTC.xlsx";
    const workbookFile = xlsx.readFile(filename);
    const sheetName = workbookFile.SheetNames[0];
    const sheet = workbookFile.Sheets[sheetName];
    if (!sheet) {
      logMessage(`❌ Sheet '${sheetName}' not found in ${filename}`);
      return { foundData, extractedValues };
    }
    const range = xlsx.utils.decode_range(sheet["!ref"]);
    // Find next available column
    let col = 3;
    while (true) {
      const cellRef = xlsx.utils.encode_cell({ c: col, r: 1 }); // check header row
      if (!sheet[cellRef]) break;
      col++;
    }
    // Write header
    const headerCellStart = xlsx.utils.encode_cell({ c: col, r: 1 });
    const headerCellEnd = xlsx.utils.encode_cell({ c: col + 1, r: 1 });
    sheet[headerCellStart] = { t: "s", v: code + " cuoi ky" };
    sheet[headerCellEnd] = { t: "s", v: code + " dau ky" };
    for (let i = 0; i < extractedValues.length; i++) {
      for (let j = 0; j < extractedValues[i].length; j++) {
        const cellRef = xlsx.utils.encode_cell({ c: col + j, r: i + 3 });
        sheet[cellRef] = { t: "s", v: extractedValues[i][j] };
      }
    }
    if (col + extractedValues[0].length - 1 > range.e.c) {
      range.e.c = col + extractedValues[0].length - 1;
      sheet["!ref"] = xlsx.utils.encode_range(range);
    }
    xlsx.writeFile(workbookFile, filename);
    logMessage(
      `✅ Đã ghi file ${filename} cho mã ${code} (sheet: ${sheetName})`
    );
  }
  return { foundData, extractedValues };
}

// Main orchestrator
async function main(inputList, year, quarter) {
  const browser = await puppeteer.launch({
    executablePath:
      "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe", // Điều chỉnh nếu dùng OS khác
    headless: false,
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
    logMessage(
      `PROGRESS: Checked ${checked}/${inputList.length}. Found: ${found}. Not found: ${notFound}.`
    );
  }

  logMessage(
    `SUMMARY: Checked ${checked} codes. Found data: ${found}. Not found: ${notFound}. Total: ${inputList.length}`
  );
  await browser.close();
}

main(input.slice(1), 2024, 4);
// main(input.slice(0, 10), 2024, 4);
