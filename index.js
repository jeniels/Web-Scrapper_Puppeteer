const excel = require("xlsx");
const puppeteer = require("puppeteer");

const workbook = excel.readFile("input.xlsx");
const worksheet = workbook.Sheets[workbook.SheetNames[0]];
const data = excel.utils.sheet_to_json(worksheet, { header: 1 });

const ISBN_numbers = [];
const book_names = [];
const result = [];

for (let i = 1; i <= 5; i++) {
  ISBN_numbers.push(String(data[i][2]));
  book_names.push(String(data[i][1]));
}

const main = async () => {
  
  const browser = await puppeteer.launch({ headless: false });
  const page = await browser.newPage();
 
  //Iterate for n books
  for (let j = 0; j < ISBN_numbers.length; j++) {

    await process(j, ISBN_numbers, book_names, browser, page, result);
  }

  for (let k = 1; k <= 5; k++) {
    let itemstoadd = [
      result[k - 1].site,
      result[k - 1].found,
      result[k - 1].url,
      result[k - 1].price,
      result[k - 1].author,
      result[k - 1].publisher,
      result[k - 1].instock,
    ];
    data[k].push(...itemstoadd);
  }

  const modifiedWorksheet = excel.utils.json_to_sheet(data);
  workbook.Sheets[workbook.SheetNames[0]] = modifiedWorksheet;
  excel.writeFile(workbook, "output.xlsx");

  await browser.close();
  return;

};
const wait = async (ms) => {
  return new Promise((resolve) => setTimeout(resolve, ms));
};

const process = async (j, ISBN_numbers, book_names, browser, page, result) => {

  const gotoPromise = page.goto("https://www.snapdeal.com");
  const timeoutPromise = wait(2000);
  await Promise.race([gotoPromise, timeoutPromise]);   //Promise Race

  await page.type("#inputValEnter", ISBN_numbers[j]);
  const searchselector =
    "#sdHeader > div.headerBar.reset-padding > div.topBar.top-bar-homepage.top-freeze-reference-point > div > div.col-xs-14.search-box-wrapper > button";
  await page.waitForSelector(searchselector);
  await page.click(searchselector);

  await page.waitForSelector(".col-xs-6.favDp.product-tuple-listing.js-tuple");


  // To find all the matching books with title
  let matched_books = await page.evaluate(
    (book_names, j) => {
      let items = document.querySelectorAll(
        ".col-xs-6.favDp.product-tuple-listing.js-tuple"
      );
      let matched_books = [];

      items.forEach((item) => {
        let title = item.querySelector(".product-title").innerText;
        title = title.replace(/\(.*?\)|:.*$|-.*$/g, "").trim();

        let price = item.querySelector(".lfloat.product-price").innerText;
        if (title.includes(book_names[j])) {
          matched_books.push({
            title: title,
            price: price,
            item: item.querySelector("a").getAttribute("href"),
          });
        }
      });
      return matched_books;
    },
    book_names,
    j
  );

  if (matched_books.length == 0) {
    result.push({
      site: "NA",
      found: "No",
      url: "NA",
      price: null,
      author: "NA",
      publisher: null,
      instock: null,
    });
    console.log("return");
    return;
  }

  //Find book with smallest price
  let bookWithSmallestPrice = matched_books[0];

  for (let i = 1; i < matched_books.length; i++) {
    if (matched_books[i].price < bookWithSmallestPrice.price) {
      bookWithSmallestPrice = matched_books[i];
    }
  }
  

  
  page.goto(`${bookWithSmallestPrice.item}`);

  page.waitForNavigation();


  await page.waitForSelector(".h-content");

  // Fetch all the details of book

  let values = await page.evaluate((bookWithSmallestPrice) => {
    let elements = document.querySelectorAll(".h-content");

    let values = {
      site: "https://www.snapdeal.com",
      found: "No",
      url: "NA",
      price: null,
      author: null,
      publisher: null,
      instock: null,
    };

    values.price = bookWithSmallestPrice.price;
    (values.url = `${bookWithSmallestPrice.item}`), (values.found = "YES");

    for (const element of elements) {
      let text = element.innerText; 
      if (text.includes("Publisher:")) {
        values.publisher = text.replace("Publisher:", "").trim();
      }
      if (text.includes("Author:")) {
        values.author = text.replace("Author:", "").trim();
      }
    }
    return values;
  }, bookWithSmallestPrice);

  if (j == 0) {
    await page.click("#pincode-check");
    await page.type("#pincode-check", "400092");
    await page.click("#pincode-check-bttn");
  }
  if (await page.waitForSelector(".itm-avail")) {
    values.instock = "Yes";
  } else {
    values.instock = "No";
  }

  //Store the details in result

  result.push(values);
};


main();
