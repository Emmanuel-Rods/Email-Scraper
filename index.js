const cheerio = require("cheerio");
const xlsx = require("xlsx");

const { connect } = require("puppeteer-real-browser");

const fileName = "Emails-locsmith in los angeles.xlsx";
const filePath =
  "C:/Users/itsro/Downloads/locsmith in los angeles.xlsx";

const extractEmails = (html) => {
  const emailRegex = /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b/g;
  const matches = html.match(emailRegex);

  if (!matches) return [];
  const validEmails = matches.filter((email) => {
    const excludedExtensions = [
      ".png",
      ".jpg",
      ".jpeg",
      ".gif",
      ".svg",
      ".webp",
      ".bmp",
    ];
    return !excludedExtensions.some((ext) => email.toLowerCase().endsWith(ext));
  });
  return Array.from(new Set(validEmails));
};

function excelToArrayJson(filePath) {
  try {
    const workbook = xlsx.readFile(filePath);
    const result = [];
    workbook.SheetNames.forEach((sheetName) => {
      const sheet = workbook.Sheets[sheetName];
      const sheetData = xlsx.utils.sheet_to_json(sheet, { defval: null });
      result.push(...sheetData);
    });

    return result;
  } catch (error) {
    console.error("Error reading the Excel file:", error);
    return [];
  }
}

const jsonArray = excelToArrayJson(filePath);

async function scraper() {
  if (jsonArray.length == 0) {
    throw new Error("Excel file empty");
  }
  //loop over the json array
  const rows = [];
  for (const data of jsonArray) {
    const email = data[Object.keys(data)[2]]; //third property
    const website = data[Object.keys(data)[3]]; // fourth propery

    if (website && !email) {
      const [emailArray] = await jsRender(website);
      data[Object.keys(data)[2]] = emailArray;
      rows.push(data); //modified data
      console.log(data);
    } else {
      rows.push(data); // just to compare , remove if neccessary
    }
  }
  //   return rows;

  //save data to excel
  const worksheet = xlsx.utils.json_to_sheet(rows);
  const workbook = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(workbook, worksheet, "Updated Data");
  xlsx.writeFile(workbook, fileName);

  console.log(`Updated Excel file saved at: ${fileName}`);
}

scraper();

async function jsRender(url) {
  const { browser, page } = await connect({
    defaultViewport: null,
    headless: false,

    args: [],

    customConfig: {},

    turnstile: true,

    connectOption: {
      defaultViewport: null,
    },
    disableXvfb: false,
    ignoreAllFlags: false,
    // proxy:{
    //     host:'<proxy-host>',
    //     port:'<proxy-port>',
    //     username:'<proxy-username>',
    //     password:'<proxy-password>'
    // }
  });
  try {
    try {
      await page.goto(url, { timeout: 30000 }); // Added timeout option
    } catch(e) {
      console.error(`${url} can't be reached`);
      return []; // Return empty array instead of stopping execution
    }
  
    const pageHTML = await page.content();
    const emails = extractEmails(pageHTML);
    console.log(emails);
  
    return emails;
  } catch (err) {
    console.error(`Error: ${err}`);
    return []; // Ensure function always returns something
  } finally {
    if (browser) {
      await browser.close(); // Safely close browser in finally block
    }
  }
}
