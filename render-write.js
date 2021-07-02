const fs = require('fs').promises;
const path = require('path');
const Docxtemplater = require("docxtemplater");
const PizZip = require("pizzip");
const { formatPrice } = require("./utils");


let MONTH_NAME;
if (process.argv[2].includes('--month=')) {
    MONTH_NAME = process.argv[2].split('--month=')[1];
} else {
    throw new Error('Please, provide a month name for file generation');
}

// write to docx
// ============================
async function renderWrite(documents) {
    await writeAddressDocuments(documents);
    await writeTotalDocument(documents);
}

async function writeAddressDocuments(documents) {
    for (const document of documents) {
        await writeDocFile('template.docx', document, `${document.address} - ${MONTH_NAME}`);
        console.log(`Report for address "${document.address}" written 👌`);
    }
}
async function writeTotalDocument(documents) {
    await writeDocFile(
    'template-total.docx',
    {
        documents,
        address_total: formatPrice(documents.reduce((acc, document) => acc + document.services_total_price, 0))
    },
    `Total report - ${MONTH_NAME}`
    );
    console.log(`Total report written 👌`);
}

async function writeDocFile(template, variables, fileName) {
    //Load the docx file as a binary
    const content = await fs.readFile(path.resolve(__dirname, template), 'binary');
    const zip = new PizZip(content);
    let doc;

    try {
        doc = new Docxtemplater(zip);
    } catch (error) {
        // Catch compilation errors (errors caused by the compilation of the template : misplaced tags)
        console.error(error);
    }

    //set the templateVariables
    doc.setData(variables);
    //

    try {
        // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
        doc.render();
    } catch (error) {
        // Catch rendering errors (errors relating to the rendering of the template : angularParser throws an error)
        console.error(error);
    }
    const buf = doc.getZip().generate({type: 'nodebuffer'});

    await fs.writeFile(path.resolve(__dirname, `output/${fileName}.docx`), buf);
}

module.exports = { renderWrite };