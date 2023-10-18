const mammoth = require('mammoth');
const fs = require('fs');

const inputFile = 'cover_letter_template.docx';
const outputFile = 'output.doc';

const companyName = 'Palms Elementary';
const positionName = 'Second Grade Teacher';

mammoth.embedStyleMap
mammoth.extractRawText({ path: inputFile })
  .then(function (result) {
    let text = result.value;
    text = text.replace(/COMPANY_NAME/g, companyName);
    text = text.replace(/POSITION_NAME/g, positionName);

    fs.writeFile(outputFile, text, function (err) {
      if (err) {
        console.error('Error writing to the output file:', err);
      } else {
        console.log('Text replaced and saved to', outputFile);
      }
    });
  })
  .catch(function (err) {
    console.error('Error converting .doc:', err);
  });
