# Financial Data Scraper

This project automates the extraction of financial data for stock codes from the SSC website and writes the results into an Excel file (`BCTC.xlsx`).

## Features
- Uses Puppeteer to automate Chrome for web scraping.
- Extracts data for each stock code and writes it to the next available column in the `CDKT` sheet of `BCTC.xlsx`.
- Logs progress and errors to `log.txt`.

## Requirements
- **Node.js** (v16 or newer recommended)
- **npm** (Node package manager)
- **Google Chrome** installed (default path: `C:\Program Files\Google\Chrome\Application\chrome.exe`)
- An Excel file named `BCTC.xlsx` with a sheet named `CDKT` (template structure as in your sample)

## Installation
1. Clone or download this repository to your local machine.
2. Open a terminal in the project directory.
3. Install dependencies:
   ```bash
   npm install
   ```

## Configuration
- Make sure `BCTC.xlsx` exists in the project folder and has a sheet named `CDKT` with the correct template (rows and columns as in your sample file).
- If your Chrome is installed in a different location, update the `executablePath` in `proxy.js`:
  ```js
  executablePath: "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
  ```
- The script reads stock codes from column E of `BCTC.xlsx` by default. Adjust the filename in `getColumnEValues` if needed.

## Usage
Run the script with Node.js:
```bash
node proxy.js
```

- The script will open Chrome, scrape data for each code, and write results to `BCTC.xlsx`.
- Progress and errors are logged in `log.txt`.

**Note:**
- The script now copies `BCTC.xlsx` to a temporary file named `BCTC_temp.xlsx` and writes all results there.
- You can keep `BCTC.xlsx` open while the script is running; your original file will not be locked or changed during processing.
- All output will be in `BCTC_temp.xlsx`.

### Customizing the Range, Year, and Quarter

To scrape data for a specific range of codes (for example, from index 100 to 200), or to change the year and quarter, edit the last line in `proxy.js`:

```js
main(input.slice(100, 201), 2024, 4); // Gets codes from index 100 to 200 (inclusive)
```
- The first argument (`input.slice(100, 201)`) specifies the range of codes to process.
- The second argument (`2024`) is the year you want to get data for.
- The third argument (`4`) is the quarter you want to get data for.

**Example:**
- To get data for codes from index 50 to 99 for year 2023, quarter 2:
  ```js
  main(input.slice(50, 100), 2023, 2);
  ```

## Troubleshooting
- **Chrome not found:**
  - Make sure Chrome is installed and the path in `proxy.js` is correct.
- **Excel file errors:**
  - Ensure `BCTC.xlsx` exists and has a `CDKT` sheet with the correct structure.
- **No data written:**
  - Check the log file for errors or missing selectors.
- **Script closes too quickly:**
  - Make sure your input list (column E) is not empty.

## License
MIT
