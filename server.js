const express = require('express');
const app = express();
const cors = require('cors');
var CronJob = require('cron').CronJob;
require('dotenv').config()
var moment = require('moment');
const puppeteer = require('puppeteer');
const path = require('path')
const fs = require('fs/promises')

const cron = require('node-cron');


const { GoogleSpreadsheet } = require('google-spreadsheet');

const doc = new GoogleSpreadsheet(process.env.GOOGLE_DOC);



cron.schedule('0 9,12,18 * * *', async () => {
    console.log('Running the tasks at 9am, 12:30pm and 6:30pm every day');
    // add your task code here
    try{


    }
    catch(err){
        console.log(err)
    }
});

app.use(cors());


app.use(express.json({
    verify: (req, res, buf) => {
      req.rawBody = buf
    }
}));


app.route('/').get(async (req,res)=>{

    const browser = await puppeteer.launch({ args: ['--no-sandbox'] })
    const page = await browser.newPage();

    const client = await page.target().createCDPSession();
    await client.send('Page.setDownloadBehavior', {
      behavior: 'allow',
      downloadPath: path.join(__dirname, 'excelDownload'),
    });

  
  // Navigate to the webpage that contains the file you want to download
    await page.goto('https://auth.promptemr.com/login');



    await page.waitForSelector('#auth0-lock-container-1');
    await page.waitForTimeout(10000)

  const email = await page.$('.auth0-lock-input[name="email"]');
  const password = await page.$('.auth0-lock-input[name="password"]');
  
  // Check if the element was found
  if (email !== null) {
    // Type in the input element
    await email.type(process.env.USERNAME);
  } else {
    console.log('Could not find input element');
  }

    if (password !== null) {
    // Type in the input element
    await password.type(process.env.PASSWORD);
  } else {
    console.log('Could not find input element');
  }

  const loginButton = await page.$('#auth0-lock-container-1 > div > div.auth0-lock-center > form > div > div > button')
  await loginButton.click()

  await page.waitForTimeout(10000)

    await page.goto('https://go.promptemr.com/reports');

    await page.waitForTimeout(10000)

    const datetime = await page.$('#q-app > div > div.q-page-container > div > div.row.items-center.q-pa-sm.no-wrap.q-card > div > div.row.no-wrap.items-center > div > div');

    // Get the bounding box of the element
    const dtloc = await datetime.boundingBox();
  
    // Calculate the coordinates of the top-left corner of the bounding box
    const x = dtloc.x;
    const y = dtloc.y;
  
    console.log(`Element coordinates: (${x}, ${y})`);
  


    // Simulate a click at the specified coordinates
    await page.mouse.click(49.9765625, 208);


    await page.waitForTimeout(10000)
    await page.waitForSelector('body > div.q-menu.q-position-engine.scroll > div > div.column.q-pa-sm.text-p-medium');

    const alltime = await page.$('body > div.q-menu.q-position-engine.scroll > div > div.column.q-pa-sm.text-p-medium > button:nth-child(7)')

    console.log(alltime)

    const atloc = await alltime.boundingBox();

    console.log(atloc.x, atloc.y)
    // await alltime.click()
    await page.mouse.click(atloc.x+1, atloc.y);

    const apply = await page.$('#q-app > div > div.q-page-container > div > div.row.items-center.q-pa-sm.no-wrap.q-card > div > div.row.no-wrap.items-center > button.q-btn.q-btn-item.non-selectable.no-outline.q-ml-sm.q-btn--unelevated.q-btn--rectangle.bg-p-blue100.text-p-blue.q-btn--actionable.q-focusable.q-hoverable.q-btn--no-uppercase.q-btn--wrap')

    await apply.click()


    await page.waitForSelector("#q-app > div > div.q-page-container > div > div.relative-position.row.q-col-gutter-md.q-px-md.q-pb-md.items-stretch > div:nth-child(3) > div > div.q-pa-md.q-gutter-xs")

    const arreport = await page.$("#q-app > div > div.q-page-container > div > div.relative-position.row.q-col-gutter-md.q-px-md.q-pb-md.items-stretch > div:nth-child(3) > div > div.q-pa-md.q-gutter-xs > div:nth-child(2) > div:nth-child(2) > button")

    await page.waitForTimeout(5000)
    await arreport.click()

    await page.waitForSelector("body > div.q-dialog.fullscreen.no-pointer-events.q-dialog--modal > div.q-dialog__inner.flex.no-pointer-events.q-dialog__inner--minimized.q-dialog__inner--standard.fixed-full.flex-center > div > div.q-pa-md.q-gutter-y-md")

   const firstDown = await page.$('body > div.q-dialog.fullscreen.no-pointer-events.q-dialog--modal > div.q-dialog__inner.flex.no-pointer-events.q-dialog__inner--minimized.q-dialog__inner--standard.fixed-full.flex-center > div > div.q-card__actions.q-card__actions--horiz.row.justify-end > button.q-btn.q-btn-item.non-selectable.no-outline.q-btn--unelevated.q-btn--rectangle.bg-p-blue100.text-p-blue500.q-btn--actionable.q-focusable.q-hoverable.q-btn--no-uppercase.q-btn--wrap')
   await page.waitForTimeout(5000)
    await firstDown.click()

    await page.waitForSelector("body > div.q-notifications > div.q-notifications__list.q-notifications__list--bottom.fixed.column.no-wrap.items-center > div > div > div.q-notification__actions.row.items-center.col-auto.q-notification__actions--with-media > button")
  const secondDown = await page.$("body > div.q-notifications > div.q-notifications__list.q-notifications__list--bottom.fixed.column.no-wrap.items-center > div > div > div.q-notification__actions.row.items-center.col-auto.q-notification__actions--with-media > button")

  await secondDown.click()


  await page.waitForTimeout(5000)

  await page.goto('https://go.promptemr.com/patients');

  await page.waitForSelector("#q-app > div > div.q-page-container > main > div > div.col.list.column.no-wrap.full-height.p-listview > div.q-pa-sm.q-card > div:nth-child(1)")
  await page.waitForSelector("#q-app > div > div.q-page-container > main > div > div.col.list.column.no-wrap.full-height.p-listview > div.q-pa-sm.q-card > div:nth-child(1) > button:nth-child(1)")
  await page.waitForTimeout(5000)
const patient = await page.$("#q-app > div > div.q-page-container > main > div > div.col.list.column.no-wrap.full-height.p-listview > div.q-pa-sm.q-card > div:nth-child(1) > button:nth-child(1)")
  await patient.click()

  await page.waitForSelector("body > div.q-menu.q-position-engine.scroll.q-pa-md > div.row.full-width.justify-end > button")

  const exportPatient = await page.$("body > div.q-menu.q-position-engine.scroll.q-pa-md > div.row.full-width.justify-end > button")
  await exportPatient.click()

  await page.waitForSelector("body > div.q-dialog.fullscreen.no-pointer-events.q-dialog--modal > div.q-dialog__inner.flex.no-pointer-events.q-dialog__inner--minimized.q-dialog__inner--standard.fixed-full.flex-center > div > div.q-card__actions.q-card__actions--horiz.row.justify-end > button.q-btn.q-btn-item.non-selectable.no-outline.q-btn--unelevated.q-btn--rectangle.bg-p-blue100.text-p-blue500.q-btn--actionable.q-focusable.q-hoverable.q-btn--no-uppercase.q-btn--wrap > span.q-btn__wrapper.col.row.q-anchor--skip")

  const secondExportPT = await page.$("body > div.q-dialog.fullscreen.no-pointer-events.q-dialog--modal > div.q-dialog__inner.flex.no-pointer-events.q-dialog__inner--minimized.q-dialog__inner--standard.fixed-full.flex-center > div > div.q-card__actions.q-card__actions--horiz.row.justify-end > button.q-btn.q-btn-item.non-selectable.no-outline.q-btn--unelevated.q-btn--rectangle.bg-p-blue100.text-p-blue500.q-btn--actionable.q-focusable.q-hoverable.q-btn--no-uppercase.q-btn--wrap")
  await page.waitForTimeout(5000)
    console.log(secondExportPT)
    

    const atloc2 = await secondExportPT.boundingBox();

    console.log(atloc2.x, atloc2.y)
    // await alltime.click()
    // await page.mouse.click(atloc2.x+1, atloc2.y);
  await secondExportPT.click()

  await page.waitForSelector("body > div.q-notifications > div.q-notifications__list.q-notifications__list--bottom.fixed.column.no-wrap.items-center > div > div > div.q-notification__actions.row.items-center.col-auto.q-notification__actions--with-media > button")

  const PTdown = await page.$("body > div.q-notifications > div.q-notifications__list.q-notifications__list--bottom.fixed.column.no-wrap.items-center > div > div > div.q-notification__actions.row.items-center.col-auto.q-notification__actions--with-media > button")
  await PTdown.click()


  await page.waitForTimeout(5000)

  browser.close()



    const folderPath = './excelDownload'; // replace with your folder path
    const fileIndexAR = 0;
    const fileIndexPat = 1; 

  try {
    // read the folder content
    const files = await fs.readdir(folderPath);

    // check if the index is within bounds
    if (fileIndexAR >= files.length) {
      console.error('Index out of bounds');
      return;
    }

    if (fileIndexPat >= files.length) {
        console.error('Index out of bounds');
        return;
      }

    // get the file name based on the index
    const AR = files[fileIndexAR];
    const patientDemo = files[fileIndexPat]
    console.log(AR)
    console.log(patientDemo)

    // construct the full path to the file
    const ARPath = path.join(folderPath, AR);

    const PTPath = path.join(folderPath, patientDemo);

    // read the file contents
    // const data = await fs.readFile(filePath, 'utf-8');

    const ARworkbook = xlsx.readFile(ARPath);
    const ARsheetName = ARworkbook.SheetNames[5];
    const ARworksheet = ARworkbook.Sheets[ARsheetName];
    const ARdata = xlsx.utils.sheet_to_json(ARworksheet);

    console.log(ARdata);


    const PTworkbook = xlsx.readFile(PTPath);
    const PTsheetName = PTworkbook.SheetNames[0];
    const PTworksheet = PTworkbook.Sheets[PTsheetName];
    const PTdata = xlsx.utils.sheet_to_json(PTworksheet);

    console.log(PTdata);


    const date = new Date();
    const options = { year: 'numeric', month: '2-digit', day: '2-digit' };
    const formattedDate = date.toLocaleDateString('en-US', options).replace(/\//g, '-');

    const combined = []
    for(let i =0; i < ARdata.length; i++){
        for(let k=0; k < PTdata.length; k++){
            if(ARdata[i]['Patient Account Number'] === PTdata[k]['Account Number']){
                if(PTdata[k]['MobilePhone']){
                    combined.push({"DateAdded": formattedDate,...ARdata[i], "phoneNumber": PTdata[k]['MobilePhone']})
                }else{
                    combined.push({"DateAdded": formattedDate, ...ARdata[i], "phoneNumber": PTdata[k]['HomePhone']})
                }

            }

        }
    }



    await doc.useServiceAccountAuth({
        client_email: process.env.GOOGLE_EMAIL,
        private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, "\n"),
    });


    await doc.loadInfo();

    const firstSheet = await doc.sheetsByIndex[2]

    const numRows = await firstSheet.rowCount;
    const range = `A2:Z${numRows}`;
    await firstSheet.clear(range);

    await firstSheet.addRows(combined.slice(1))


    await firstSheet.saveUpdatedCells();




  } catch (err) {
    console.error(err);
    res.status(400).send(err.message)
  }



})




// define the first route




// app.use('/api/webhooks', webhooks);

// app.use(express.static('client/build'));

// if(process.env.NODE_ENV === 'production'){
//     const path = require('path');
//     app.get('/*',(req,res)=>{
//         res.sendFile(path.resolve(__dirname,'../client','build','index.html'))
//     })
// }








const port = process.env.PORT || 3001

app.listen(port, () =>{
    console.log('SERVER RUNNING', port)
})