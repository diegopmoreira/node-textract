"use strict"

const xlsx = require('node-xlsx').default;
const JSZip = require('jszip');
const Docxtemplater = require('docxtemplater');
const fs = require('fs');
const path = require('path');
const workSheetsFromFile = xlsx.parse(`${__dirname}/batch.xlsx`);
let array_table;
var content = fs
  .readFileSync(path.resolve(__dirname, 'standard.docx'), 'binary');

var zip = new JSZip(content);
var doc = new Docxtemplater();

array_table = workSheetsFromFile[1].data;
doc.loadZip(zip);


array_table.forEach((row, index) => {


  //set the templateVariables
  doc.setData({
    address_url: row[0],
    current_title: row[1],
    new_title: row[2],     
    current_description: row[3],
    new_description: row[4],
    current_keywords: row[5],
    new_keywords: row[6]
  });

  

  try {
    // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
    doc.render()
  }
  catch (error) {
    var e = {
      message: error.message,
      name: error.name,
      stack: error.stack,
      properties: error.properties,
    }
    console.log(JSON.stringify({ error: e }));
    // The error thrown here contains additional information when logged with JSON.stringify (it contains a property object).
    throw error;
  }

  var buf = doc.getZip()
    .generate({ type: 'nodebuffer' });

  // buf is a nodejs buffer, you can either write it to a file or do anything else with it.
  fs.writeFileSync(path.resolve(__dirname, `output${index}.docx`), buf);
});