import puppeteer from "puppeteer-core";
import xlsx from "xlsx";
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

const input = [
  "AAA",
  // { code: "TCM" },
];

const workbook = xlsx.utils.book_new();

const handleSearch = async (code, page) => {
  //select input and type code
  await page.waitForSelector('input[id$="pt9:it8112::content"]');
  await page.type('input[id$="pt9:it8112::content"]', code);

  await page.evaluate(() => {
    const links = Array.from(document.querySelectorAll("a"));

    const searchLink = links.find(
      (a) => a.querySelector("span")?.textContent?.trim() === "Tìm kiếm"
    );

    if (searchLink) {
      searchLink.click();
    }
  });

  await new Promise((resolve) => setTimeout(resolve, 3000));
};

const handleGetMaxRecord = async () => {
  const totalCount = await page.evaluate(() => {
    const span = document.querySelector('span[id$="pt9:it4"]');
    if (!span) return 0;

    const text = span.textContent || "";
    const match = text.match(/\d+/); // tìm số đầu tiên trong chuỗi

    return match ? parseInt(match[0], 10) : 0;
  });

  return totalCount;
};

const handleNavigateToPage = async (page, pageNumber, pageContext) => {
  await pageContext.waitForSelector("a.x14f"); // đảm bảo các thẻ a đã xuất hiện

  await pageContext.evaluate((pageNumber) => {
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
};

const main = async (inputList, year, quarter) => {
  const browser = await puppeteer.launch({
    executablePath:
      "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe", // Điều chỉnh nếu dùng OS khác
    headless: false,
    defaultViewport: null,
  });

  //wait for chrome launch instant
  const page = await browser.newPage();
  await handleNavigateToPage(1 + 1, 15, page)
  await new Promise((resolve) => setTimeout(resolve, 10000));

  for (const inputItem of inputList) {
    const code = inputItem;

    //search page
    await page.goto("https://congbothongtin.ssc.gov.vn/faces/NewsSearch", {
      waitUntil: "domcontentloaded",
    });

    await handleSearch(code, page);

    const maxRecord = await handleGetMaxRecord();

    //check a.xgl be show on the screen
    await page.waitForSelector("a.xgl", { timeout: 10000 });

    //get all of a.xgl
    const linkCount = await page.$$eval("a.xgl", (links) => links.length);
    const pageData = [];

    for (let i = 0; i < maxRecord; i++) {
      if (i !== 0 && (i + 1) % 15 === 0) {
        await handleNavigateToPage(i / 15 + 1, 15);
      }

      await new Promise((resolve) => setTimeout(resolve, 1000));

      await page.waitForSelector("a.xgl", { timeout: 10000 });

      const links = await page.$$("a.xgl");
      const link = links[i];

      if (!link) continue;

      await link.click();
      await new Promise((resolve) => setTimeout(resolve, 5000));

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

      if (extractedData?.length > 0) {
        pageData.push(extractedData);
        console.log(`✅ Page ${i + 1} có dữ liệu`);
      } else {
        console.log(`⚠️ Page ${i + 1} không có dữ liệu, bỏ qua`);
      }

      await page.goBack({ waitUntil: "domcontentloaded" });
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
  }

  // const now = new Date();
  // const timestamp = now.toISOString().replace(/[:.]/g, "-"); // thay : và . để hợp lệ
  // const fileName = `ssc_result_${timestamp}.xlsx`;
  // xlsx.writeFile(workbook, fileName);
  // console.log("✅ Đã ghi file " + fileName);

  // await browser.close();
};

main(input, 2025, 1);

// const year = document.querySelector(
//   'span[id="pt2:tt1:2:lookupValueId::content"]'
// );
// const quarter = document.querySelector(
//   'span[id="pt2:tt1:3:lookupValueId::content"]'
// );
