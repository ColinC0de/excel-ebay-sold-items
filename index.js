const XLSX = require('xlsx');
const electron = require('electron');
const { clipboard } = electron;
const { Builder, By, Key, until } = require('selenium-webdriver');
const puppeteer = require('puppeteer');
const request = require('request');

async function grabData (cell) {
 
  return new Promise((resolve, reject) => {
      var options = {
          'method': 'POST',
          'url': 'https://ebay-sold-items-api.herokuapp.com/findCompletedItems',
          'headers': {
              'Content-Type': 'application/json'
          },
          body: JSON.stringify({
              "keywords": cell,
              "max_search_results": "60",
          })
      };
      request(options, function (error, response) {
          if (error) reject(error);
          else resolve(JSON.parse(response.body));
      });
  });
}

async function ScanColumnB() {
  // read the excel file
  const workbook = XLSX.readFile('./Dell-ebay-sold-script.xlsx');
  const sheet = workbook.Sheets["Sheet2"];
  const columnB = "B";
  let i = 3
  
  while(i < 904) {
  let cell = sheet[columnB + i];
  if(!cell)  break;
  let cellValue = cell.v;
 
  let data = await grabData(cellValue);
  let numberOfResults = data.results;
  let AveragePrice = data.average_price;

  let dataToAdd = [[numberOfResults, AveragePrice]];
XLSX.utils.sheet_add_aoa(sheet, dataToAdd, {origin: "G"+ i});
XLSX.writeFile(workbook, './Dell-ebay-sold-script.xlsx');

i++;
  }
}

ScanColumnB();



