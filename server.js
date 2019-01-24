const express = require('express');
const path = require('path');
const app = express();
const docx = require('docx');
const fs = require('fs');
const { Document, Paragraph, Packer } = docx;
// simulating data from database
const { price, bookingTerms} = require('./paragraphs');
/*

Aproximetly 1 pixel is 9525 emu's

*/
function fromPixeltoEmu(pixels) {
    return 9525 * pixels;
}

app.get('/word', async (req, res) => {
    const styles = fs.readFileSync('./styles.xml', 'utf-8');
    const doc = new Document(undefined, {
        top: 0,
        right: 0,
        bottom: 0,
        left: 0
    });
    var title1 = new docx.Paragraph("First Page").heading1().center().pageBreak();
    var title2 = new docx.Paragraph("Second Page").heading1().center().pageBreak();
    var title3 = new docx.Paragraph("Third Page").heading1().center().pageBreak();

    doc.createImage(fs.readFileSync(path.join(__dirname, 'images/120-sibenik.jpg')), 690, 300, {
        floating: {
            horizontalPosition: {
                offset: fromPixeltoEmu(50),
            },
            verticalPosition: {
                offset: fromPixeltoEmu(50),
            },
        },
    });
    doc.createImage(fs.readFileSync(path.join(__dirname, 'images/zadar.jpg')), 690, 300, {
        floating: {
            horizontalPosition: {
                offset: fromPixeltoEmu(50),
            },
            verticalPosition: {
                offset: fromPixeltoEmu(610),
            },
        },
    });
    doc.addParagraph(title1);
    doc.createImage(fs.readFileSync(path.join(__dirname, 'images/120-sibenik.jpg')), 200, 133, {
        floating: {
            horizontalPosition: {
                offset: fromPixeltoEmu(50),
            },
            verticalPosition: {
                offset: fromPixeltoEmu(50),
            },
        },
    });
    doc.createImage(fs.readFileSync(path.join(__dirname, 'images/zadar.jpg')), 200, 133, {
        floating: {
            horizontalPosition: {
                offset: fromPixeltoEmu(300),
            },
            verticalPosition: {
                offset: fromPixeltoEmu(50),
            },
        },
    });
    doc.createImage(fs.readFileSync(path.join(__dirname, 'images/slikapula.jpg')), 200, 133, {
        floating: {
            horizontalPosition: {
                offset: fromPixeltoEmu(550),
            },
            verticalPosition: {
                offset: fromPixeltoEmu(50),
            },
        },
    });
    doc.addParagraph(title2);
    doc.addParagraph(title3);

    bookingTerms.map((data) => {
        var text = new docx.Paragraph(data.PaymentPolicyText)
        doc.addParagraph(text);
    });
    // packaging the word document
    const packer = new Packer();
    const b64string = await packer.toBase64String(doc);
    res.setHeader('Content-Disposition', 'attachment; filename=Ispis.docx');
    res.send(Buffer.from(b64string, 'base64'));

});

app.listen(3000);
console.log('Listening on port 3000');

