const express = require('express');
const fs = require('fs');
const PizZip = require('pizzip');
const cors = require('cors');
const Docxtemplater = require("docxtemplater");
const {v4: uuid} = require("uuid");
const {contentDisposition} = require("express/lib/utils");
const Excel = require('exceljs');
const app = express();
const PORT = 5000;

app.use(cors())
app.use(express.json());

const temp = {};
const docFields = [
  'email', 'father_name', 'first_name', 'last_name', 'id_code', 'region', 'settlement', 'address', 'index',
  'phone_number', 'passport_number', 'passport_institute', 'passport_date', 'passport_series'
]
const dataFields = ['specialization', 'full_name', 'learning_mode','id_code', 'passport_series', 'passport_number',
  'passport_institute', 'passport_date', 'index', 'settlement'];

app.get('/documents/download', (req, res) => {
  const { id } = req.query;

  if (!(id in temp)) {
    throw res.status(400).json('Неправильний ідентифікатор завантаження');
  }
  const doc = temp[id];

  delete temp[id];

  res.setHeader('Content-Disposition', contentDisposition(doc.filename));
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');

  res.status(200).send(doc.buffer);
});

app.get('/data', (req, res) => {
  const {surname} = req.query;
  const data = [];
  const wb = new Excel.Workbook();
  wb.xlsx.readFile('./Збір_даних.xlsx').then(() => {
    const ws = wb.worksheets[0];
    let i;

    for (i = 2; i < ws.actualRowCount; i++){

      const row = ws.findRow(i);
      console.log(i + ' : ' + row.values[2] + ' : ' + surname);
      if (row.values[2].includes(surname)){
        const obj = {};

        for (const [index, fieldName] of dataFields.entries()){
          obj[fieldName] = row.values[index + 1];
        }
        obj.address = row.values[11] + ', ' + row.values[12] + ', ' + row.values[13];
        obj.phone_number = row.values[14];
        obj.email = row.values[15];
        console.log('Here' + obj.passport_date);
        if (obj.passport_date) obj.passport_date = obj.passport_date.slice(0, obj.passport_date.indexOf('T')).replaceAll('-','');
        const s = obj.passport_date;
        obj.passport_date = s.slice(6) + '.' + s.slice(4, 6) + '.' + s.slice(0, 4);

        const names = obj.full_name.split(' ');
        obj.last_name = names[0];
        obj.first_name = names[1];
        obj.father_name = names[2];
        delete obj.full_name;

        if (obj.settlement){
          const address = obj.settlement.split(',');
          obj.region = address[0];
          obj.settlement = address[address.length - 1];
        }

        if (obj.phone_number.includes('(')){
          const index = obj.phone_number.indexOf('(');
          obj.country_code = obj.phone_number.slice(0, index);
          obj.phone_number = obj.phone_number.slice(index + 1, index + 14).replace(')', '').replaceAll('-', '').replace(';', '');
        }
        else{
          const str = obj.phone_number;
          obj.country_code = str.slice(0, obj.phone_number.length-11);
          obj.phone_number = str.slice(obj.phone_number.length - 11).replace(';', '');
        }

        data.push(obj);
      }
    }
    console.log(data);
    res.status(200).send(data);
  });


});

app.post('/documents/download', async (req, res) => {
  const { data } = req.body;

  if (typeof data !== "object") {
    throw "Wrong data";
  }

  for (const name of docFields) {
    if (!data[name]) data[name] = '';
    if (!data['parent_' + name] || data['parent_' + name] === '+380') data['parent_' + name] = '';
  }

  data['index'] = '';
  data['parent_index'] = '';

  data['big'] = data['last_name'].toUpperCase();
  data['parent_big'] = data['parent_last_name'].toUpperCase();

  if (data['id_code'] === "") data['id_code'] = data['passport_number']
  if (data['parent_id_code'] === "") data['parent_id_code'] = data['parent_passport_number']
  if (data['region'] !== "" && data['settlement'] !== "") data['region'] += ',';
  if (data['settlement'] !== "" && data['address'] !== "") data['settlement'] += ',';
  if (data['address'] !== "" && data['index'] !== "") data['address'] += ',';
  if (data['passport_number'] !== "" && data['passport_institute'] !== "") data['passport_number'] += ',';
  if (data['passport_institute'] !== "" && data['passport_date'] !== "") data['passport_institute'] += ',';
  if (data['parent_region'] !== "" && data['parent_settlement'] !== "") data['parent_region'] += ',';
  if (data['parent_settlement'] !== "" && data['parent_address'] !== "") data['parent_settlement'] += ',';
  if (data['parent_address'] !== "" && data['parent_index'] !== "") data['parent_address'] += ',';
  if (data['parent_passport_number'] !== "" && data['parent_passport_institute'] !== "") data['parent_passport_number'] += ',';
  if (data['parent_passport_institute'] !== "" && data['parent_passport_date'] !== "") data['parent_passport_institute'] += ',';
  if (data['parent_first_name'] === "") data['noParent'] = true;

  const fileName1 = `${data.specialization}_Контракт_${data.learning_mode}.docx`;
  const buffer1 = generateDoc(`./templates_education/${fileName1}`, data);
  const id1 = uuid();
  temp[id1] = { buffer: buffer1, fileName: fileName1 };

  if (data['noParent']){

    for (const name of docFields) {
      if (data['parent_' + name] === '') data['parent_' + name] = data[name];
    }
    data['big'] = data['last_name'].toUpperCase();
    data['parent_big'] = data['parent_last_name'].toUpperCase();
  }
  const fileName2 = `${data.specialization}_Контракт_${data.learning_mode}_${data.payment_period}.docx`;
  const buffer2 = generateDoc(`./templates_payment/${fileName2}`, data);
  const id2 = uuid();
  temp[id2] = { buffer: buffer2, fileName: fileName2};

  res.status(200).json({id1, id2});
});

app.listen(PORT, () => {
  console.log('Stared on PORT: ' + PORT);
})

function generateDoc(path, data){
  const content = fs.readFileSync(path, 'binary');
  const zip = new PizZip(content);
  const document = new Docxtemplater(zip);

  document.setData(data);
  document.render();

  const buffer = document.getZip().generate({
    type: 'nodebuffer'
  });

  return buffer;
}