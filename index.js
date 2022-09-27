const express = require('express');
const fs = require('fs');
const PizZip = require('pizzip');
const cors = require('cors');
const Docxtemplater = require("docxtemplater");
const {v4: uuid} = require("uuid");
const {contentDisposition} = require("express/lib/utils");
const app = express();
const PORT = 5000;

app.use(cors({
  origin: '*',
}))
app.use(express.json());

const temp = {};
const docFields = [
  'email', 'father_name', 'first_name', 'last_name', 'id_code', 'region', 'settlement', 'address', 'index',
  'phone_number', 'passport_number', 'passport_institute', 'passport_date', 'passport_series'
]

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

app.post('/documents/download', async (req, res) => {
  const { data } = req.body;

  if (typeof data !== "object") {
    throw "Wrong data";
  }

  for (const name of docFields) {
    if (!data[name]) data[name] = '';
    if (!data['parent_' + name] || data['parent_' + name] === '+380') data['parent_' + name] = '';
  }
  if (!data['program']) data['program'] = 'ОНП';

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

  const fileName1 = `${data.payment_type}_${data.learning_mode}_${data.specialization}_${data.program}.docx`;
  const buffer1 = generateDoc(`./templates_education/${fileName1}`, data);
  const id1 = uuid();
  temp[id1] = { buffer: buffer1, fileName: fileName1 };

  if (data.payment_type === 'Контракт'){
    if (data['noParent']){

      for (const name of docFields) {
        if (data['parent_' + name] === '') data['parent_' + name] = data[name];
      }
      data['big'] = data['last_name'].toUpperCase();
      data['parent_big'] = data['parent_last_name'].toUpperCase();
    }

    let fileName2 = `${data.specialization}_Контракт_${data.learning_mode}_${data.payment_period}.docx`;

    if (data.specialization === '121'){
      if (data.program.includes('ІПІ')){
        fileName2 = `121_ІПІ_Контракт_${data.learning_mode}_${data.payment_period}.docx`;
      }
      else{
        fileName2 = `121_ОТ_Контракт_${data.learning_mode}_${data.payment_period}.docx`;
      }
    }

    const buffer2 = generateDoc(`./templates_payment/${fileName2}`, data);
    const id2 = uuid();
    temp[id2] = { buffer: buffer2, fileName: fileName2};
    res.status(200).json({id1, id2});
  }
  else{
    res.status(200).json({id1});
  }

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
