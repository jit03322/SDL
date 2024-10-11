const express = require('express');
const multer = require('multer');
const pdfParse = require('pdf-parse');
const fs = require('fs');
const { GoogleGenerativeAI } = require('@google/generative-ai');
const { Document, Packer, Paragraph, HeadingLevel } = require('docx');
const XLSX = require('xlsx');

// Initialize the Express app
const app = express();

// Initialize Google Gemini API with your API key
const GOOGLE_API_KEY = "AIzaSyBrqxE7HJD9TGg_7Lh9DKINKCILVcO86qg"; // Replace with your actual API key
const genAI = new GoogleGenerativeAI(GOOGLE_API_KEY);

// Set up multer for handling file uploads (PDFs)
const upload = multer({ dest: 'uploads/' });

app.use(express.urlencoded({ extended: true })); // To parse form data

app.get('/', (req, res) => {
  res.send(`
    <!DOCTYPE html>
    <html lang="en">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>Upload PDF</title>
    </head>
    <body>
      <h2>Upload a PDF</h2>
      <form ref='uploadForm' 
        id='uploadForm' 
        action='/upload-pdf' 
        method='post' 
        encType="multipart/form-data">
          <label for="pdf">Choose a PDF file:</label>
          <input type="file" name="pdf" accept="application/pdf" required />
          <br><br>
          <label for="heading">Enter Heading:</label>
          <input type="text" name="heading" required />
          <br><br>
          <input type='submit' value='Upload PDF!' />
      </form>
    </body>
    </html>
  `);
});

app.post('/upload-pdf', upload.single('pdf'), async (req, res) => {
  const pdfPath = req.file.path; // Access the uploaded PDF file
  const userHeading = req.body.heading; // Get the user-provided heading

  try {
    // Read the uploaded PDF file
    const pdfData = fs.readFileSync(pdfPath);

    // Use pdf-parse to extract text from the PDF
    const pdfText = await pdfParse(pdfData);

    // Pass the extracted text and user heading to the Gemini API for further processing
    const geminiResponse = await sendToGemini(pdfText.text, userHeading);

    // Determine if the response is in paragraph or tabular format
    if (isTabularContent(geminiResponse)) {
      const excelFilePath = await generateExcelFile(geminiResponse);
      res.download(excelFilePath, (err) => {
        if (err) {
          console.error('Error sending the Excel file:', err);
          res.status(500).send('Failed to download Excel file.');
        }
        // Clean up the file after download
        fs.unlinkSync(excelFilePath);
      });
    } else {
      const docFilePath = await generateDocFile(geminiResponse);
      res.download(docFilePath, (err) => {
        if (err) {
          console.error('Error sending the DOC file:', err);
          res.status(500).send('Failed to download DOC file.');
        }
        // Clean up the file after download
        fs.unlinkSync(docFilePath);
      });
    }
  } catch (error) {
    console.error('Error processing PDF:', error);
    res.status(500).send('Failed to process PDF.');
  } finally {
    // Clean up the uploaded file after use
    fs.unlinkSync(pdfPath);
  }
});

// Function to check if the content is in tabular format
function isTabularContent(response) {
  // Example condition: check if the response contains a specific table format indicator
  return response.includes('Table:'); // You can refine this check as needed
}

// Function to generate a DOC file with organized content
async function generateDocFile(content) {
  const doc = new Document({
    sections: [
      {
        properties: {},
        children: [
          new Paragraph({
            text: "Extracted Content",
            heading: HeadingLevel.HEADING_1,
          }),
          new Paragraph({
            text: content,
          }),
        ],
      },
    ],
  });

  const docFilePath = `output/Extracted_Content_${Date.now()}.docx`;
  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(docFilePath, buffer);

  return docFilePath;
}

// Function to generate an Excel file with tabular content
async function generateExcelFile(content) {
  // Assuming content is in CSV format for this example
  const rows = content.split('\n').map(row => row.split(','));

  const ws = XLSX.utils.aoa_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Extracted Content');

  const excelFilePath = `output/Extracted_Content_${Date.now()}.xlsx`;
  XLSX.writeFile(wb, excelFilePath);

  return excelFilePath;
}

// Function to send the extracted text and user heading to the Google Gemini API
async function sendToGemini(pdfText, userHeading) {
  const model = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });

  // Define the prompt that includes the user heading and detailed extraction instructions
  const prompt = `
    Based on the extracted text from the PDF, find the heading "${userHeading}" and return its associated content. If the heading is not present, return a message indicating that.
    Extracted PDF Text:
    ${pdfText}
  `;

  // Send the prompt and the extracted PDF text to the Gemini API
  const result = await model.generateContent([prompt]);

  const responseText = await result.response.text();

  return responseText;
}

// Start the server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
