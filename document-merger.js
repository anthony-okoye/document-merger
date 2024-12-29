const fs = require('fs');
const path = require('path');
const merger = require('@groupdocs/groupdocs.merger');

async function mergeDocuments(jsonFilePath, templateDocxPath, outputDocxPath, outputPdfPath) {
    try {
        // Read JSON data from the file
        const jsonData = JSON.parse(fs.readFileSync(jsonFilePath, 'utf8'));

        // Read the DOCX template file
        if (!fs.existsSync(templateDocxPath)) {
            throw new Error(`Template DOCX file not found: ${templateDocxPath}`);
        }

        const mergerInstance = new merger.Merger();

        // Replace placeholders in the DOCX template with JSON data
        let docxContent = fs.readFileSync(templateDocxPath, 'utf8');
        Object.entries(jsonData).forEach(([key, value]) => {
            const placeholder = `{{${key}}}`;
            docxContent = docxContent.replaceAll(placeholder, value);
        });

        const tempDocxPath = 'temp-merged-document.docx';
        fs.writeFileSync(tempDocxPath, docxContent);

        // Add the processed DOCX to the merger instance
        mergerInstance.add(tempDocxPath);

        // Save the merged document as DOCX
        await mergerInstance.save(outputDocxPath);
        console.log(`Merged DOCX saved at: ${outputDocxPath}`);

        // Convert the merged DOCX to PDF
        mergerInstance.add(outputDocxPath); // Ensure the merged DOCX is used
        await mergerInstance.save(outputPdfPath, 'pdf');
        console.log(`Converted PDF saved at: ${outputPdfPath}`);

        // Cleanup temporary file
        fs.unlinkSync(tempDocxPath);
    } catch (error) {
        console.error('Error during merging or conversion:', error);
    }
}

// Files 
const jsonFilePath = path.join(__dirname, 'data.json'); // Path to JSON file
const templateDocxPath = path.join(__dirname, 'template.docx'); // Path to DOCX template
const outputDocxPath = path.join(__dirname, 'merged-document.docx'); // Output DOCX file
const outputPdfPath = path.join(__dirname, 'merged-document.pdf'); // Output PDF file

mergeDocuments(jsonFilePath, templateDocxPath, outputDocxPath, outputPdfPath);
