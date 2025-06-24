# Financial Data Scraper

This project automates the extraction of financial data for stock codes from the SSC (State Securities Commission) website and writes the results into an Excel file (`BCTC.xlsx`).

## Features
- Uses Puppeteer to automate Chrome for web scraping
- Extracts financial data from multiple sheets (CDKT, KQKD, LCTT-GT) for each stock code
- Reads stock codes from `codes.xlsx` file (Column E)
- Writes data to multiple sheets in `BCTC.xlsx` with proper column headers
- Comprehensive logging system with different log levels (INFO, ERROR, SUCCESS, WARNING, DEBUG)
- Supports pagination and handles multiple search results
- Configurable year and quarter parameters for data extraction

## Requirements
- **Node.js** (v16 or newer recommended)
- **npm** (Node package manager)
- **Google Chrome** installed (default path: `C:\Program Files\Google\Chrome\Application\chrome.exe`)
- An Excel file named `codes.xlsx` with stock codes in Column E
- An Excel file named `BCTC.xlsx` with three sheets: CDKT, KQKD, and LCTT-GT

## Installation
1. Clone or download this repository to your local machine
2. Open a terminal in the project directory
3. Install dependencies:
   ```bash
   npm install
   ```

## Configuration
- Make sure `codes.xlsx` exists in the project folder with stock codes in Column E
- Ensure `BCTC.xlsx` exists with three sheets: CDKT, KQKD, and LCTT-GT
- If your Chrome is installed in a different location, update the `CHROME_PATH` in the CONFIG object in `proxy.js`:
  ```js
  CHROME_PATH: "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
  ```

## Usage
Run the script with Node.js:
```bash
node proxy.js
```

The script will:
- Read stock codes from `codes.xlsx` (Column E)
- Open Chrome and navigate to the SSC website
- Search for each stock code and extract financial data
- Write data to three sheets in `BCTC.xlsx`:
  - **CDKT** (Cân đối kế toán): Balance sheet data
  - **KQKD** (Kết quả kinh doanh): Income statement data  
  - **LCTT-GT** (Lưu chuyển tiền tệ): Cash flow data
- Log all progress and errors to `log.txt`

### Customizing the Range, Year, and Quarter

To scrape data for a specific range of codes or change the year and quarter, edit the last line in `proxy.js`:

```js
// Process all codes for Q4 2024
main(input, 2024, 4);

// Process codes from index 1 to end for Q4 2024
main(input.slice(1), 2024, 4);

// Process first 10 codes for Q4 2024
main(input.slice(0, 10), 2024, 4);
```

**Parameters:**
- First argument: Array of stock codes to process
- Second argument: Year (e.g., 2024)
- Third argument: Quarter (1-4)

**Examples:**
- To get data for codes from index 50 to 99 for year 2023, quarter 2:
  ```js
  main(input.slice(50, 100), 2023, 2);
  ```

## Output Structure

The script writes data to three sheets in `BCTC.xlsx`:

1. **CDKT Sheet**: Balance sheet data with columns for each stock code
2. **KQKD Sheet**: Income statement data with columns for each stock code  
3. **LCTT-GT Sheet**: Cash flow data with columns for each stock code

Each stock code gets two columns: `{CODE} cuoi ky` (end of period) and `{CODE} dau ky` (beginning of period).

## Logging

The script provides comprehensive logging with different levels:
- **INFO**: General progress information
- **SUCCESS**: Successful data extraction
- **WARNING**: Non-critical issues
- **ERROR**: Errors that need attention
- **DEBUG**: Detailed debugging information

All logs are written to `log.txt` and also displayed in the console.

## Troubleshooting
- **Chrome not found:**
  - Make sure Chrome is installed and the path in `proxy.js` is correct
- **Excel file errors:**
  - Ensure `codes.xlsx` exists with stock codes in Column E
  - Ensure `BCTC.xlsx` exists with three sheets: CDKT, KQKD, LCTT-GT
- **No data written:**
  - Check the log file for errors or missing selectors
  - Verify the stock codes exist on the SSC website
- **Script closes too quickly:**
  - Make sure your input list (Column E in codes.xlsx) is not empty

## Dependencies
- `puppeteer-core`: Chrome automation
- `xlsx`: Excel file reading and writing
- `fs`: File system operations
- `axios`: HTTP requests (if needed for future features)

## License
ISC
