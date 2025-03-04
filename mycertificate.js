import fs from "fs";
import path from "path";
import { exec } from "child_process";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import xlsx from "xlsx";
import { showBanner, announceStart, announceSuccess } from "./notifier.mjs";

// Define paths
const outputDir = "./certificates"; // Directory where certificates will be stored
const excelFilePath = "./students.xlsx"; // Excel file containing student names
const sheetName = "Sheet1"; // Name of the sheet in Excel
const inputPath =
  "./certificates/Template/Merakicertificate - Introduction_to _Scratch_world.docx"; // Path to the DOCX template

// Show banner
showBanner();

// Announce certificate generation
announceStart();

// Delay for 2 seconds before proceeding (optional)
setTimeout(() => {
  announceSuccess();
}, 2000);

// Ensure the output directory exists
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir, { recursive: true });
}

// Function to create a sample Excel file if it doesn't exist
function createSampleExcel(filePath) {
  const data = [{ Name: "UJJWAL" }]; // Example entry
  const worksheet = xlsx.utils.json_to_sheet(data);
  const workbook = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(workbook, worksheet, sheetName);

  xlsx.writeFile(workbook, filePath);
  console.log(`📄 Created sample Excel file: ${filePath}`);
  console.log(
    "✏️  Please edit 'students.xlsx' and add student names before running the script again."
  );
}

// Function to read student names from Excel
function readStudentsFromExcel(filePath) {
  if (!fs.existsSync(filePath)) {
    console.warn("⚠️ Excel file not found. Creating a new one...");
    createSampleExcel(filePath);
    return [];
  }

  const workbook = xlsx.readFile(filePath);
  const sheet = workbook.Sheets[sheetName];

  if (!sheet) {
    console.error(`❌ Sheet "${sheetName}" not found in ${filePath}`);
    return [];
  }

  const data = xlsx.utils.sheet_to_json(sheet);
  return data.map((row) => row.Name).filter((name) => name);
}

// Function to generate a certificate for a given student
function generateCertificate(
  replacements,
  outputPath,
  pdfOutputPath,
  callback
) {
  try {
    const content = fs.readFileSync(inputPath, "binary");
    const zip = new PizZip(content);
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
    });
    doc.render(replacements);

    fs.writeFileSync(outputPath, doc.getZip().generate({ type: "nodebuffer" }));
    console.log(`✅ DOCX certificate generated: ${outputPath}`);

    exec(
      `libreoffice --headless --convert-to pdf "${outputPath}" --outdir "${outputDir}"`,
      (err) => {
        if (err) {
          console.error(
            `❌ Error converting to PDF for ${replacements.Name}:`,
            err
          );
          callback();
          return;
        }

        fs.renameSync(outputPath.replace(".docx", ".pdf"), pdfOutputPath);
        console.log(`✅ PDF certificate generated: ${pdfOutputPath}`);

        fs.unlinkSync(outputPath);
        console.log(`🗑️ Deleted DOCX file: ${outputPath}`);

        callback();
      }
    );
  } catch (error) {
    console.error(
      `❌ Error generating certificate for ${replacements.Name}:`,
      error
    );
    callback();
  }
}

// Read student names from the Excel file
const students = readStudentsFromExcel(excelFilePath);

if (students.length === 0) {
  console.error(
    "❌ No student names found in 'students.xlsx'. Please add names and re-run the script."
  );
  process.exit(1);
}

// Process students one by one (to prevent concurrent execution issues)
function processStudents(index) {
  if (index >= students.length) return;

  const student = students[index];
  const outputPath = path.join(outputDir, `certificate_${student}.docx`);
  const pdfOutputPath = path.join(outputDir, `${student}.pdf`);
  const replacements = { Name: student };

  generateCertificate(replacements, outputPath, pdfOutputPath, () => {
    processStudents(index + 1);
  });
}

// Start processing
processStudents(0);
