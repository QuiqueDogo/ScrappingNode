const puppeteer = require('puppeteer');
const xlsx = require("xlsx");
const spreadsheet = xlsx.readFile('./Prueba.xlsx');
const sheets = spreadsheet.SheetNames;
const fisrtSheet = spreadsheet.Sheets[sheets[0]];
var XLSX_CALC = require('xlsx-calc');
var formulajs = require('@formulajs/formulajs');
XLSX_CALC.import_functions(formulajs);


exports.calc = (async () => {
const emisoras = [
    ['TSLA','https://finance.yahoo.com/quote/TSLA?p=TSLA&.tsrc=fin-srch',27],
    ['FB','https://finance.yahoo.com/quote/FB?p=FB&.tsrc=fin-srch',28],
    ['V','https://finance.yahoo.com/quote/V?p=V&.tsrc=fin-srch',29],
    ['BRK B','https://finance.yahoo.com/quote/BERK.AS?p=BERK.AS&.tsrc=fin-srch',30],
    ['WMT','https://finance.yahoo.com/quote/WMT?p=WMT&.tsrc=fin-srch',31],
    ['AMZ','https://finance.yahoo.com/quote/AMZN?p=AMZN&.tsrc=fin-srch',32],
    ['AAPL','https://finance.yahoo.com/quote/AAPL?p=AAPL&.tsrc=fin-srch',33],
    ['UAA','https://finance.yahoo.com/quote/UAA?p=UAA&.tsrc=fin-srch',34],
    ['GOOGL','https://finance.yahoo.com/quote/GOOGL?p=GOOGL&.tsrc=fin-srch',35],
    ['MSFT','https://finance.yahoo.com/quote/MSFT?p=MSFT&.tsrc=fin-srch',36],
    ['BABA','https://finance.yahoo.com/quote/BABA?p=BABA&.tsrc=fin-srch',37],
    ['NFLX','https://finance.yahoo.com/quote/NFLX?p=NFLX&.tsrc=fin-srch',38],
    ['SBUX','https://finance.yahoo.com/quote/SBUX?p=SBUX&.tsrc=fin-srch',39],
    ['MELI','https://finance.yahoo.com/quote/MELI?p=MELI&.tsrc=fin-srch',40],
    ['BKNG','https://finance.yahoo.com/quote/BKNG?p=BKNG&.tsrc=fin-srch',41],
    ['EXPE','https://finance.yahoo.com/quote/EXPE?p=EXPE&.tsrc=fin-srch',42],
    ['TRIP','https://finance.yahoo.com/quote/TRIP?p=TRIP&.tsrc=fin-srch',43],
    ['TRVG','https://finance.yahoo.com/quote/TRVG?p=TRVG&.tsrc=fin-srch',44],
    ['FSLR','https://finance.yahoo.com/quote/FSLR?p=FSLR&.tsrc=fin-srch',45],
    ['SPWR','https://finance.yahoo.com/quote/SPWR?p=SPWR&.tsrc=fin-srch',46],
    ['RUN','https://finance.yahoo.com/quote/RUN?p=RUN&.tsrc=fin-srch',47],
    ['NVDA','https://finance.yahoo.com/quote/NVDA?p=NVDA&.tsrc=fin-srch',48],
    ['INTC','https://finance.yahoo.com/quote/INTC?p=INTC&.tsrc=fin-srch',49],
    ['AMD','https://finance.yahoo.com/quote/AMD?p=AMD&.tsrc=fin-srch,',50]
];
let all = [];
  const browser = await puppeteer.launch();
  const page = await browser.newPage();
  await page.setRequestInterception(true);
  page.on('request', (request) => {
      if (request.resourceType() === 'document') {
          request.continue();
      } else {
          request.abort();
      }
  });
  const time = Date.now();
    for (let i = 0; i < emisoras.length; i++) {
        await page.goto(emisoras[i][1]);
        await page.waitForSelector('[data-test="qsp-price"]');
        const text = await page.$eval('[data-test="qsp-price"]', el => el.innerText);
        all = [...all, [text, emisoras[i][0], emisoras[i][2]]];    
        
    }
    all.forEach(element => {
        let celda = fisrtSheet[`B${element[2]}`];
        let CeldaNombre = fisrtSheet[`A${element[2]}`].v;
        //verificamos  y usamos xlsx para actualizar
        if(element[1] == CeldaNombre){
            celda.w = element[0];
            celda.v = parseFloat(element[0]);
            fisrtSheet[`C${element[2]}`] = {t:'n', F:`F1*B${element[2]}`, f:`F1*B${element[2]}`};
        }
        
    });
    xlsx.writeFileXLSX(spreadsheet, "PruebaActualizada.xlsx");
    console.log(all,  `Libro actualizado en : ${(Date.now() - time)/1000}sec`);
   
  await browser.close();
})();