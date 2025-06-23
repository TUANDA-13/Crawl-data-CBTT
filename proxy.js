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
        if (!tableDiv) return [];

        const tbody = tableDiv.querySelector("tbody");
        if (!tbody) return [];

        const rows = Array.from(tbody.querySelectorAll("tr")).slice(1);

        return rows.map((row) => {
          const cells = row.querySelectorAll("td");
          const td5Span =
            cells[4]?.querySelector("span")?.textContent?.trim() || "-";
          const td6Span =
            cells[5]?.querySelector("span")?.textContent?.trim() || "-";
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
  const pageData = [];
  let foundData = false;
  for (let i = 0; i < maxRecord || currentPage <= 4; i++) {
    if (i === 0 || (i + 1) % 15 === 0) {
      // If not the first page, check if next page is available
      if (i !== 0) {
        // Check if a next page exists
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
      // Re-perform the search after navigating
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
      pageData.push(extractedData);
      foundData = true;
      logMessage(`✅ Page ${i + 1} has data, moving to next code.`);
      break; // Stop searching further pages for this code
    } else {
      logMessage(`⚠️ Page ${i + 1} has no data, skipping`);
    }

    await page.goBack({ waitUntil: "domcontentloaded" });
    await typeAndSearch(page, 'input[id$="pt9:it8112::content"]', code);
    await clickSearchButton(page);
    await page.waitForSelector("a.xgl");
  }

  // ✅ Xử lý và ghi dữ liệu: ghép các cột theo từng trang, bỏ qua dòng trống
  const header = [];
  const sheetRows = [];

  pageData.forEach((_, pageIndex) => {
    header.push(`${code} (2024 - page ${pageIndex + 1})`);
    header.push(`${code} (2023 - page ${pageIndex + 1})`);
  });

  sheetRows.push(header);

  const maxRowCount = Math.max(...pageData.map((d) => d.length), 0);

  for (let i = 0; i < maxRowCount; i++) {
    const row = [];

    pageData.forEach((data) => {
      const item = data[i];
      row.push(item?.end ?? "");
      row.push(item?.start ?? "");
    });

    const isEmpty = row.every((cell) => !cell || cell === "-" || cell === "");
    if (!isEmpty) {
      sheetRows.push(row);
    }
  }

  const worksheet = xlsx.utils.aoa_to_sheet(sheetRows);
  xlsx.utils.book_append_sheet(workbook, worksheet, code);
  return foundData;
}

// Main orchestrator
async function main(inputList, year, quarter) {
  const browser = await puppeteer.launch({
    executablePath:
      "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe", // Điều chỉnh nếu dùng OS khác
    headless: false,
    defaultViewport: null,
  });

  //wait for chrome launch instant
  const page = await browser.newPage();

  let checked = 0;
  let found = 0;
  let notFound = 0;

  for (const inputItem of inputList) {
    const code = inputItem;
    checked++;
    const hasData = await processCompany(page, code, year, quarter);
    if (hasData) {
      found++;
    } else {
      notFound++;
    }
    // Log progress after each code
    logMessage(`PROGRESS: Checked ${checked}/${inputList.length}. Found: ${found}. Not found: ${notFound}.`);
  }

  const now = new Date();
  const timestamp = now.toISOString().replace(/[:.]/g, "-");
  const timeStr = now.toTimeString().split(" ")[0].replace(/:/g, "-");
  const fileName = `ssc_result_${timestamp}_${timeStr}.xlsx`;
  xlsx.writeFile(workbook, fileName);
  logMessage(`✅ Đã ghi file ${fileName}`);

  // Add summary log
  logMessage(`SUMMARY: Checked ${checked} codes. Found data: ${found}. Not found: ${notFound}. Total: ${inputList.length}`);

  await browser.close();
}

main(input, 2024, 4);

