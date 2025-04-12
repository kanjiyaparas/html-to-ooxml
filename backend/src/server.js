
const express = require('express');
const cors = require('cors');
const bodyParser = require('body-parser');
const { JSDOM } = require('jsdom');
const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');

const app = express();
app.use(cors());
app.use(bodyParser.json());

const allowedOrigins = ['https://html--ooxml.vercel.app'];

app.use(cors({
  origin: function (origin, callback) {
    // Allow requests with no origin (like Postman)
    if (!origin) return callback(null, true);
    if (allowedOrigins.includes(origin)) {
      return callback(null, true);
    } else {
      return callback(new Error('CORS policy violation: Not allowed by CORS'));
    }
  }
}));


function escapeXml(unsafe) {
    return unsafe
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}

// ✅ Map named CSS colors to hex
const namedColors = {
    red: 'FF0000',
    blue: '0000FF',
    green: '008000',
    black: '000000',
    white: 'FFFFFF',
    gray: '808080',
    yellow: 'FFFF00',
    purple: '800080',
    orange: 'FFA500',
    pink: 'FFC0CB',
    brown: 'A52A2A',
    cyan: '00FFFF',
    magenta: 'FF00FF',
};

function convertToOOXML(html) {
    const dom = new JSDOM(html);
    const doc = dom.window.document;
    const paragraphs = doc.querySelectorAll('p');

    // Recursive function to convert each node with inherited styles
    const convertNode = (node, inheritedStyles = {}) => {
        if (node.nodeType === 3) {
            // Text node
            const text = escapeXml(node.textContent.trim());
            if (!text) return '';

            let styleXml = '';
            if (inheritedStyles.bold) styleXml += '<w:b/>';
            if (inheritedStyles.italic) styleXml += '<w:i/>';
            if (inheritedStyles.color) styleXml += `<w:color w:val="${inheritedStyles.color}"/>`;
            if (inheritedStyles.size) styleXml += `<w:sz w:val="${inheritedStyles.size}"/>`;

            return `<w:r>${styleXml ? `<w:rPr>${styleXml}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
        }

        if (node.nodeType === 1) {
            const tag = node.tagName.toLowerCase();
            const styleAttr = node.getAttribute('style') || '';

            // Inherit styles from parent
            const newStyles = { ...inheritedStyles };

            // Check for bold
            if (tag === 'strong' || tag === 'b' || /font-weight\s*:\s*bold/.test(styleAttr)) {
                newStyles.bold = true;
            }

            // Check for italic
            if (tag === 'em' || tag === 'i' || /font-style\s*:\s*italic/.test(styleAttr)) {
                newStyles.italic = true;
            }

            // Enhanced Color: capture both hex and named colors
            const colorMatch = styleAttr.match(/color\s*:\s*([^;]+)/);
            if (colorMatch) {
                let colorVal = colorMatch[1].trim();
                if (colorVal.startsWith('#')) {
                    newStyles.color = colorVal.replace('#', '');
                } else {
                    // Convert named color to hex using our mapping:
                    let lower = colorVal.toLowerCase();
                    if (namedColors[lower]) {
                        newStyles.color = namedColors[lower];
                    } else {
                        // If not found, keep the value (you might consider defaulting it)
                        newStyles.color = colorVal;
                    }
                }
            }

            // Check for font size
            const sizeMatch = styleAttr.match(/font-size\s*:\s*(\d+)px/);
            if (sizeMatch) {
                newStyles.size = parseInt(sizeMatch[1]) * 2; // Word uses half-points
            }

            let xml = '';
            node.childNodes.forEach(child => {
                xml += convertNode(child, newStyles);
            });

            return xml;
        }

        return '';
    };

    // Convert each paragraph using our recursive function
    const convertParagraph = (p) => {
        let xml = '<w:p>';
        xml += convertNode(p);
        xml += '</w:p>';
        return xml;
    };

    let bodyContent = '';
    paragraphs.forEach(p => {
        bodyContent += convertParagraph(p);
    });

    return `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>${bodyContent}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body>
</w:document>`;
}

// 🔧 Keep your original route intact with enhancements
// app.post('/convert', (req, res) => {
//     const ooxml = convertToOOXML(req.body.html);
//     console.log(ooxml)
//     const templatePath = path.join(__dirname, 'templet.docx');
//     const templateBinary = fs.readFileSync(templatePath, 'binary');
//     const zip = new PizZip(templateBinary);
//     zip.file('word/document.xml', ooxml);

//     const docxBuffer = zip.generate({ type: 'nodebuffer' });

//     res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
//     res.setHeader('Content-Disposition', 'attachment; filename="output.docx"');
//     res.send(docxBuffer);
// });

// 🔍 Preview route (returns OOXML XML only)
app.post('/convert-preview', (req, res) => {
    const ooxml = convertToOOXML(req.body.html);
    return res.json({ ooxml });
});

// 📄 Download route (returns .docx)
app.post('/convert-docx', (req, res) => {
    const ooxml = convertToOOXML(req.body.html);
    const templatePath = path.join(__dirname, 'templet.docx');
    const templateBinary = fs.readFileSync(templatePath, 'binary');
    const zip = new PizZip(templateBinary);
    zip.file('word/document.xml', ooxml);

    const docxBuffer = zip.generate({ type: 'nodebuffer' });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename="output.docx"');
    res.send(docxBuffer);
});




app.listen(3001, () => {
    console.log('Server is running on http://localhost:3001');
});

