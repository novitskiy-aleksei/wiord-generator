const fs = require('fs').promises;
const path = require('path');
const Docxtemplater = require("docxtemplater");
const PizZip = require("pizzip");


const FILENAME_DATE = 'Март';

// write to docx
// ============================
async function renderWrite(documents) {
    for (const document of documents) {
        //Load the docx file as a binary
        const content = await fs.readFile(path.resolve(__dirname, 'template.docx'), 'binary');
        const zip = new PizZip(content);
        let doc;
        try {
            doc = new Docxtemplater(zip);
        } catch (error) {
            // Catch compilation errors (errors caused by the compilation of the template : misplaced tags)
            console.error(error);
        }

        //set the templateVariables
        doc.setData(document);
        //

        try {
            // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
            doc.render();
        } catch (error) {
            // Catch rendering errors (errors relating to the rendering of the template : angularParser throws an error)
            console.error(error);
        }
        const buf = doc.getZip().generate({type: 'nodebuffer'});

        const file_name = `output/${document.address} - ${FILENAME_DATE}.docx`;

        await fs.writeFile(path.resolve(__dirname, file_name), buf);
        console.log(`Report for address "${document.address}" written 👌`);
    }
}

module.exports = { renderWrite };