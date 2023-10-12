const fspromise = require('fs/promises');
const fs = require('fs')
const puppeteer = require('puppeteer');
const path = require('path');
const xlsx = require('xlsx');
const moment = require('moment')
const { GoogleSpreadsheet } = require('google-spreadsheet');
require('dotenv').config()


const googleEmail = process.env.GOOGLE_EMAIL
const googlePK = process.env.GOOGLE_PRIVATE_KEY




function clearFolder (folderPath) {

    fs.readdir(folderPath, (err, files) => {
      if (err) {
        console.error('Error reading folder:', err);
        return;
      }
  
      files.forEach(file => {
        const filePath = path.join(folderPath, file);
  
        fs.stat(filePath, (err, stats) => {
          if (err) {
            console.error('Error retrieving file stats:', err);
            return;
          }
  
          if (stats.isFile()) {
            fs.unlink(filePath, err => {
              if (err) {
                console.error('Error deleting file:', err);
                return;
              }
  
              console.log('Deleted file:', filePath);
            });
          }
        });
      });
    });
}

const PCO = async () => {

    const folderPath = 'utils/kpiDownload'
    const fileIndexPCO = 1;

    try{



      const doc = new GoogleSpreadsheet(process.env.GOOGLE_KPI_DOC);

      await doc.useServiceAccountAuth({
          client_email: googleEmail,
          private_key: googlePK.replace(/\\n/g, "\n"),
      });


      await doc.loadInfo();

      const firstSheet = await doc.sheetsByIndex[0]


      const providers = await firstSheet.getRows()

    //   console.log(providers)






        const files = await fspromise.readdir(folderPath);

        console.log(files)

        // check if the index is within bounds
        if (fileIndexPCO >= files.length) {
          console.error('Index out of bounds', "1");
          return;
        }
    
      //Clear file function
    
        // get the file name based on the index
        const detailedCharges = files[fileIndexPCO];
    
        // construct the full path to the file
        const chargesPath = path.join(folderPath, detailedCharges);

        const workbook = xlsx.readFile(chargesPath);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = xlsx.utils.sheet_to_json(worksheet);


        console.log(data, "Data")


        const inputDate = new Date(convertNumericDateToFormattedDate(data[0]['Date Of Service'])[0]);
        const upcomingSaturday = getUpcomingSaturday(inputDate);
        const previousSunday = getPreviousSunday(inputDate)
    
        const formattedOutput = upcomingSaturday.toLocaleDateString('en-US', {
          month: '2-digit',
          day: '2-digit',
          year: 'numeric'
        });
    
        const sundayFormattedOutput = previousSunday.toLocaleDateString('en-US', {
          month: '2-digit',
          day: '2-digit',
          year: 'numeric'
        });
    
        const monthName = upcomingSaturday.toLocaleString('en-US', { month: 'long' });
        const year = upcomingSaturday.getFullYear();
      
        const monthYear = `${monthName} ${year}`;
    
        console.log(monthYear, sundayFormattedOutput, formattedOutput);
        const visits = []
        const noCode = []
        const newData = []

        for(let i = 0; i < data.length; i++ ){
            if ((!visits.includes(data[i]['Claim Number'])) & (data[i]['CPT Code'] === 97530)){
             visits.push({claim: data[i]['Claim Number'], code: data[i]['CPT Code'], loc:data[i]['Visit Facility'], therapist: data[i]['Treating Therapist']});
            }
            if(!noCode.includes(data[i]['Claim Number'])){
                noCode.push(data[i]['Claim Number'])
                newData.push({claim :data[i]['Claim Number'], loc:data[i]['Visit Facility'], therapist: data[i]['Treating Therapist']})
            }
        }

        console.log(visits)

        // const providers = []
        // const cleanPro = []
        // for(let k = 0; k < data.length; k++ ){
        //     if (!cleanPro.includes(data[k]['Treating Therapist'])) {
        //     cleanPro.push(data[k]['Treating Therapist'])
        //      providers.push({therapist: data[k]['Treating Therapist'], location: data[k]['Visit Facility']});
        //     }
        // }

        // // console.log(providers)

        const numbers = []

        for(let j = 0; j < providers.length; j++){
            
          var counter = 0;
          var counter2 = 0;

          for(let f = 0; f < newData.length; f++){
            if(providers[j]["ProviderNameFL"] === newData[f].therapist){
              counter++
            }
          }

          for(let d = 0; d < visits.length; d++){
              if(providers[j]["ProviderNameFL"] === visits[d].therapist){
                counter2++
              }
          }
            

          numbers.push({month: monthYear , week: `${sundayFormattedOutput} - ${formattedOutput}`, therapist: providers[j]['ProviderName'], pc: counter2, totalVisits: counter, pocRate: parseFloat(((counter2/counter)*100).toFixed(2)) })

            
        }


      console.log(numbers)


      return numbers

    }catch(err){
        
    }

}

const arrivalRate = async () => {

    const folderPath = 'utils/kpiDownload'
    const fileIndex = 4;

    try{
        
      
      const doc = new GoogleSpreadsheet(process.env.GOOGLE_KPI_DOC);

      await doc.useServiceAccountAuth({
        client_email: googleEmail,
        private_key: googlePK.replace(/\\n/g, "\n"),
      });


      await doc.loadInfo();

      const firstSheet = await doc.sheetsByIndex[0]


      const providers = await firstSheet.getRows()

      console.log(providers[0]["ProviderNameFL"])



        const files = await fspromise.readdir(folderPath);

        // check if the index is within bounds
        if (fileIndex >= files.length) {
          console.error('Index out of bounds');
          return;
        }
    
      //Clear file function
    
        // get the file name based on the index
        const detailedCharges = files[fileIndex];
    
        // construct the full path to the file
        const chargesPath = path.join(folderPath, detailedCharges);

        const workbook = xlsx.readFile(chargesPath);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = xlsx.utils.sheet_to_json(worksheet);



        console.log(data)

        const inputDate = new Date(convertNumericDateToFormattedDate(data[0]['Date'])[0]);
        const upcomingSaturday = getUpcomingSaturday(inputDate);
        const previousSunday = getPreviousSunday(inputDate)
    
        const formattedOutput = upcomingSaturday.toLocaleDateString('en-US', {
          month: '2-digit',
          day: '2-digit',
          year: 'numeric'
        });
    
        const sundayFormattedOutput = previousSunday.toLocaleDateString('en-US', {
          month: '2-digit',
          day: '2-digit',
          year: 'numeric'
        });
    
        const monthName = upcomingSaturday.toLocaleString('en-US', { month: 'long' });
        const year = upcomingSaturday.getFullYear();
      
        const monthYear = `${monthName} ${year}`;
    
        console.log(monthYear, sundayFormattedOutput, formattedOutput);



        const numbers = []

        for(let j = 0; j < providers.length; j++){
            
            console.log(providers[j]["ProviderNameFL"])
            var counter = 0;
            var counter2 = 0;

            for(let f = 0; f < data.length; f++){
                if(providers[j]["ProviderNameFL"] === data[f]['Provider']){
                    counter++
                }
            }

            for(let d = 0; d < data.length; d++){
              if((providers[j]["ProviderNameFL"] === data[d]['Provider']) && (data[d]['Prompt Claim Number'])){
                counter2++
              }
            }
            

            numbers.push({month: monthYear , week: `${sundayFormattedOutput} - ${formattedOutput}` ,therapist: providers[j]["ProviderName"], arrivedVisits: counter2, totalScheduled: counter, arrivalRate: parseFloat((100 - (counter2/counter)*100).toFixed(2))})

            
        }


        console.log(numbers)


      return numbers


    }catch(err){
        
    }

}


function convertNumericDateToFormattedDate(numericDate) {
  const referenceDate = new Date('12-30-1899');
  const actualDate = new Date(referenceDate.getTime() + numericDate * 24 * 60 * 60 * 1000);

  const year = actualDate.getFullYear();
  const month = String(actualDate.getMonth() + 1).padStart(2, '0');
  const day = String(actualDate.getDate()).padStart(2, '0');

  const formattedDate = `${month}/${day}/${year}`;
  
  return [formattedDate, month];
}

function getNextMonday(date) {
  const day = date.getDay(); // Get the day of the week (0 - Sunday, 1 - Monday, ..., 6 - Saturday)
  const daysUntilNextMonday = (8 - day) % 7; // Calculate the number of days until the next Monday
  const nextMonday = new Date(date.getFullYear(), date.getMonth(), date.getDate() + daysUntilNextMonday);
  return nextMonday;
}

function getPreviousSunday(date) {
  const dayOfWeek = date.getDay();
  const daysToSubtract = dayOfWeek === 0 ? 7 : dayOfWeek;
  const previousSunday = new Date(date);
  previousSunday.setDate(date.getDate() - daysToSubtract);
  return previousSunday;
}

function getUpcomingSaturday(date) {
  const dayOfWeek = date.getDay();
  const daysToAdd = dayOfWeek === 6 ? 7 : 6 - dayOfWeek;
  const upcomingSaturday = new Date(date);
  upcomingSaturday.setDate(date.getDate() + daysToAdd);
  return upcomingSaturday;
}

const unitsPerWeek = async ()=>{
  try{

    const doc = new GoogleSpreadsheet(process.env.GOOGLE_KPI_DOC);

      await doc.useServiceAccountAuth({
        client_email: googleEmail,
        private_key: googlePK.replace(/\\n/g, "\n"),
    });


    await doc.loadInfo();

    const firstSheet = await doc.sheetsByIndex[0]


    const providers = await firstSheet.getRows()

    console.log(providers[0]["ProviderName"])





    const folderPath = 'utils/kpiDownload'
    const fileIndex = 0;

    const files = await fspromise.readdir(folderPath);

    // check if the index is within bounds
    if (fileIndex >= files.length) {
      console.error('Index out of bounds');
      return;
    }

  //Clear file function

    // get the file name based on the index
    const detailedCharges = files[fileIndex];

    // construct the full path to the file
    const chargesPath = path.join(folderPath, detailedCharges);

    const workbook = xlsx.readFile(chargesPath);
    const sheetName = workbook.SheetNames[1];
    const worksheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(worksheet);

    // console.log(data)



    // Get the Monday
    const  sheetDos= workbook.SheetNames[0];
    const worksheetDos = workbook.Sheets[sheetDos];
    const date = xlsx.utils.sheet_to_json(worksheetDos);

    // console.log(convertNumericDateToFormattedDate(date[0].DOS))

    const inputDate = new Date(convertNumericDateToFormattedDate(date[0].DOS)[0]);
    const upcomingSaturday = getUpcomingSaturday(inputDate);
    const previousSunday = getPreviousSunday(inputDate)

    const formattedOutput = upcomingSaturday.toLocaleDateString('en-US', {
      month: '2-digit',
      day: '2-digit',
      year: 'numeric'
    });

    const sundayFormattedOutput = previousSunday.toLocaleDateString('en-US', {
      month: '2-digit',
      day: '2-digit',
      year: 'numeric'
    });

    const monthName = upcomingSaturday.toLocaleString('en-US', { month: 'long' });
    const year = upcomingSaturday.getFullYear();
  
    const monthYear = `${monthName} ${year}`;

    console.log(formattedOutput);


    const claims = []
    for(let k = 0; k < providers.length; k++ ){
      for(let i = 0; i < data.length ; i++){
        if(providers[k].ProviderName === data[i]['Provider']){
          claims.push({month: monthYear , week: `${sundayFormattedOutput} - ${formattedOutput}`, therapist: data[i]['Provider'], billed: data[i]['Units']});
        }
      }
    }

    console.log(claims)

  return claims

}catch(err){
    console.log(err)
}

}


const providerNps = async ()=>{

    try{

      const doc = new GoogleSpreadsheet(process.env.GOOGLE_KPI_DOC);

      await doc.useServiceAccountAuth({
        client_email: googleEmail,
        private_key: googlePK.replace(/\\n/g, "\n"),
      });


      await doc.loadInfo();

      const firstSheet = await doc.sheetsByIndex[0]


      const providers = await firstSheet.getRows()

      // console.log(providers[0]["ProviderName"])



        const folderPath = 'utils/kpiDownload'
        const fileIndex = 2;

        const files = await fspromise.readdir(folderPath);

        // check if the index is within bounds
        if (fileIndex >= files.length) {
          console.error('Index out of bounds');
          return;
        }
    
      //Clear file function
    
        // get the file name based on the index
        const detailedCharges = files[fileIndex];
    
        // construct the full path to the file
        const chargesPath = path.join(folderPath, detailedCharges);

        const workbook = xlsx.readFile(chargesPath);
        const sheetName = workbook.SheetNames[1];
        const worksheet = workbook.Sheets[sheetName];
        const data = xlsx.utils.sheet_to_json(worksheet);

        console.log(data)


        const providersNps = []
        // for(let i=0; i < providers.length; i++ ){
        //   // console.log(providers[i]['ProviderNameFL'])
        //   for(let k = 0; k < data.length; k++ ){

        //     if(providers[i]['ProviderNameFL'] === data[k]['Provider Name']){
        //       providersNps.push({therapist: providers[i]['ProviderName'], npsRate: data[k]['Provider NPS']});
        //     }else if(!data.includes(providers[i]['ProviderNameFL'])){
        //       providersNps.push({therapist: providers[i]['ProviderName'], npsRate: 0});
        //     }
        //   }
        // }




      for (let i = 0; i < providers.length; i++) {
          let found = false; // Flag to track if a match is found
          for (let k = 0; k < data.length; k++) {
              if (providers[i]['ProviderNameFL'] === data[k]['Provider Name']) {
                  providersNps.push({
                      therapist: providers[i]['ProviderName'],
                      npsRate: data[k]['Provider NPS']
                  });
                  found = true; // Mark that a match is found
                  break; // No need to continue searching in data
              }
          }
      
          // If no match is found in data, add with npsRate: 0
          if (!found) {
              providersNps.push({
                  therapist: providers[i]['ProviderName'],
                  npsRate: 0
              });
          }
      }

      
      
        console.log(providersNps)




        const sheetNameClinic = workbook.SheetNames[0];
        const worksheetClinic = workbook.Sheets[sheetNameClinic];
        const dataClinic = xlsx.utils.sheet_to_json(worksheetClinic);

        // console.log(dataClinic)






      return providersNps





    }catch(err){

    }




    
}







const combine = async () =>{

    const doc = new GoogleSpreadsheet(process.env.GOOGLE_KPI_DOC);

    await doc.useServiceAccountAuth({
      client_email: googleEmail,
      private_key: googlePK.replace(/\\n/g, "\n"),
    });


    await doc.loadInfo();

    const firstSheet = await doc.sheetsByIndex[0]


    const providers = await firstSheet.getRows()
    



    try{

        const poc = await PCO()
        console.log(poc)
        const arRate = await arrivalRate()
        const unitsRate = await unitsPerWeek()
        const proNps = await providerNps()


        const previousMonth = moment().subtract(1, 'month');

        // Format the previous month as a string (e.g., "August 2023")
        const previousMonthString = previousMonth.format('MMMM YYYY');



        const com = []

        for(let i=0;  i < providers.length; i++){
          for(let d = 0; d < poc.length; d++){
            for(let j =0; j< arRate.length; j++){
                for(let k =0; k< unitsRate.length; k++){
                    for(let g =0; g < proNps.length; g++){
                        if((providers[i]["ProviderName"] === poc[d].therapist) && (providers[i]["ProviderName"] === arRate[j].therapist) && (providers[i]["ProviderName"] === unitsRate[k].therapist)&& (providers[i]["ProviderName"] === proNps[g].therapist)){
                          com.push({Month: unitsRate[k].month , Week: unitsRate[k].week , ProviderName: providers[i]["ProviderName"], Unit_Pro: unitsRate[k].billed, PatientCareOpt : poc[d].pocRate, ArrivalRate: arRate[j].arrivalRate, Nps: proNps[g].npsRate })
                        }
                    }
                }
            }
          }
        }




        console.log(com)




        await doc.useServiceAccountAuth({
            client_email: googleEmail,
            private_key: googlePK.replace(/\\n/g, "\n"),
        });
    
    
        await doc.loadInfo();
    
        const firstSheet = await doc.sheetsByIndex[6]
    
        const numRows = await firstSheet.rowCount;
        const range = `A2:F${numRows}`;
        // await firstSheet.clear(range);
    
        await firstSheet.addRows(com)
    
    
        await firstSheet.saveUpdatedCells();
    












    }catch(err){
        console.log(err)
    }

}



const getDataFromPrompt = async () =>{


  const ColsFolder = path.join(__dirname, './kpiDownload');
  await clearFolder(ColsFolder);




  const browser = await puppeteer.launch({ headless: false });
  const page = await browser.newPage();

  const client = await page.target().createCDPSession();
  await client.send('Page.setDownloadBehavior', {
    behavior: 'allow',
    downloadPath: path.join(__dirname, 'kpiDownload'),
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

  // await page.waitForTimeout(30000)

  // const lastMonthButton = await page.$('body > div.q-menu.q-position-engine.scroll > div > div.column.q-pa-sm.text-p-medium > div:nth-child(6) > button:nth-child(4) > span.q-btn__wrapper.col.row.q-anchor--skip > span') 
  const lastMonthButton = await page.$('body > div.q-menu.q-position-engine.scroll > div > div.column.q-pa-sm.text-p-medium > div:nth-child(6) > button:nth-child(3) > span.q-btn__wrapper.col.row.q-anchor--skip > span')


  await lastMonthButton.click()

  const applyFilter = await page.$('#q-app > div > div.q-page-container > div > div.row.items-center.q-pa-sm.no-wrap.q-card > div > div.row.no-wrap.items-center > button.q-btn.q-btn-item.non-selectable.no-outline.q-ml-sm.q-btn--unelevated.q-btn--rectangle.bg-p-blue100.text-p-blue.q-btn--actionable.q-focusable.q-hoverable.q-btn--no-uppercase.q-btn--wrap > span.q-btn__wrapper.col.row.q-anchor--skip > span')
  
  await applyFilter.click()

  await page.waitForTimeout(5000)


  const opsReport = await page.$('#q-app > div > div.q-page-container > div > div.relative-position.row.q-col-gutter-md.q-px-md.q-pb-md.items-stretch > div:nth-child(1) > div > div.q-pa-md.q-gutter-xs > div:nth-child(1) > div:nth-child(2) > button > span.q-btn__wrapper.col.row.q-anchor--skip > span')
  await opsReport.click()
  await page.waitForSelector("body > div.q-notifications > div.q-notifications__list.q-notifications__list--bottom.fixed.column.no-wrap.items-center > div > div > div.q-notification__actions.row.items-center.col-auto.q-notification__actions--with-media > button",{ timeout: 0 })

  const opsReportDown = await page.$("body > div.q-notifications > div.q-notifications__list.q-notifications__list--bottom.fixed.column.no-wrap.items-center > div > div > div.q-notification__actions.row.items-center.col-auto.q-notification__actions--with-media > button")
  await opsReportDown.click()

  await page.waitForTimeout(10000)

  const capReport = await page.$('#q-app > div > div.q-page-container > div > div.relative-position.row.q-col-gutter-md.q-px-md.q-pb-md.items-stretch > div:nth-child(2) > div > div.q-pa-md.q-gutter-xs > div:nth-child(1) > div:nth-child(2) > button > span.q-btn__wrapper.col.row.q-anchor--skip > span')
  await capReport.click()


  await page.waitForSelector("body > div.q-notifications > div.q-notifications__list.q-notifications__list--bottom.fixed.column.no-wrap.items-center > div > div > div.q-notification__actions.row.items-center.col-auto.q-notification__actions--with-media > button",{ timeout: 0 })

  const capReportDown = await page.$("body > div.q-notifications > div.q-notifications__list.q-notifications__list--bottom.fixed.column.no-wrap.items-center > div > div > div.q-notification__actions.row.items-center.col-auto.q-notification__actions--with-media > button")
  await capReportDown.click()
  await page.waitForTimeout(10000)

  const detReport = await page.$('#q-app > div > div.q-page-container > div > div.relative-position.row.q-col-gutter-md.q-px-md.q-pb-md.items-stretch > div:nth-child(2) > div > div.q-pa-md.q-gutter-xs > div:nth-child(3) > div:nth-child(2) > button > span.q-btn__wrapper.col.row.q-anchor--skip > span')
  await detReport.click()

  await page.waitForTimeout(10000)

  await page.waitForSelector("body > div.q-notifications > div.q-notifications__list.q-notifications__list--bottom.fixed.column.no-wrap.items-center > div > div > div.q-notification__actions.row.items-center.col-auto.q-notification__actions--with-media > button",{ timeout: 0 })

  const detReportDown = await page.$("body > div.q-notifications > div.q-notifications__list.q-notifications__list--bottom.fixed.column.no-wrap.items-center > div > div > div.q-notification__actions.row.items-center.col-auto.q-notification__actions--with-media > button")
  await detReportDown.click()

  await page.waitForTimeout(5000)

  await page.goto('https://go.promptemr.com/dashboard');

  await page.waitForSelector("#NPSSurveys-downloadButton > span.q-btn__wrapper.col.row.q-anchor--skip > span")

  await page.waitForTimeout(10000)

  const npsReport = await page.$("#NPSSurveys-downloadButton > span.q-btn__wrapper.col.row.q-anchor--skip > span")
  await npsReport.click()

  await page.waitForSelector("body > div.q-notifications > div.q-notifications__list.q-notifications__list--bottom.fixed.column.no-wrap.items-center > div > div > div.q-notification__actions.row.items-center.col-auto.q-notification__actions--with-media > button")

  const npsReportDown = await page.$("body > div.q-notifications > div.q-notifications__list.q-notifications__list--bottom.fixed.column.no-wrap.items-center > div > div > div.q-notification__actions.row.items-center.col-auto.q-notification__actions--with-media > button")
  await npsReportDown.click()

  await page.waitForTimeout(10000)

  const googleReviewReport = await page.$("#OnlineReviews-downloadButton > span.q-btn__wrapper.col.row.q-anchor--skip > span")
  await googleReviewReport.click()


  await page.waitForSelector("body > div.q-notifications > div.q-notifications__list.q-notifications__list--bottom.fixed.column.no-wrap.items-center > div > div > div.q-notification__actions.row.items-center.col-auto.q-notification__actions--with-media > button")

  const googleReviewReportDown = await page.$("body > div.q-notifications > div.q-notifications__list.q-notifications__list--bottom.fixed.column.no-wrap.items-center > div > div > div.q-notification__actions.row.items-center.col-auto.q-notification__actions--with-media > button")
  await googleReviewReportDown.click()

  await page.waitForTimeout(10000)

  browser.close()


  await combine()

}


module.exports = { getDataFromPrompt };