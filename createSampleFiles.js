const XLSX = require('xlsx');
const { PDFDocument, StandardFonts, rgb } = require('pdf-lib');
const fs = require('fs-extra');
const path = require('path');

async function createSampleExcel() {
    // Sample data with names and emails for certificate generation and email delivery
    const data = [
        { name: 'John Doe', email: 'john.doe@example.com', department: 'Engineering', phone: '+1-555-0101' },
        { name: 'Jane Smith', email: 'jane.smith@example.com', department: 'Marketing', phone: '+1-555-0102' },
        { name: 'Mike Johnson', email: 'mike.johnson@example.com', department: 'Sales', phone: '+1-555-0103' },
        { name: 'Sarah Wilson', email: 'sarah.wilson@example.com', department: 'HR', phone: '+1-555-0104' },
        { name: 'David Brown', email: 'david.brown@example.com', department: 'Finance', phone: '+1-555-0105' },
        { name: 'Lisa Davis', email: 'lisa.davis@example.com', department: 'Operations', phone: '+1-555-0106' },
        { name: 'Tom Anderson', email: 'tom.anderson@example.com', department: 'IT', phone: '+1-555-0107' },
        { name: 'Emily Taylor', email: 'emily.taylor@example.com', department: 'Design', phone: '+1-555-0108' },
        { name: 'Alex Rodriguez', email: 'alex.rodriguez@example.com', department: 'Research', phone: '+1-555-0109' },
        { name: 'Maria Garcia', email: 'maria.garcia@example.com', department: 'Quality Assurance', phone: '+1-555-0110' }
    ];

    // Create workbook and worksheet
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet(data);

    // Add worksheet to workbook
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Participants');

    // Write to file
    const excelPath = path.join(__dirname, 'sample-names.xlsx');
    XLSX.writeFile(workbook, excelPath);

    console.log(`Sample Excel file created: ${excelPath}`);
    return excelPath;
}

async function createSamplePDFTemplate() {
    // Create a new PDF document
    const pdfDoc = await PDFDocument.create();

    // Add a page
    const page = pdfDoc.addPage([600, 800]);

    // Get fonts
    const titleFont = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
    const bodyFont = await pdfDoc.embedFont(StandardFonts.Helvetica);

    // Get page dimensions
    const { width, height } = page.getSize();

    // Draw certificate border
    page.drawRectangle({
        x: 50,
        y: 50,
        width: width - 100,
        height: height - 100,
        borderColor: rgb(0, 0, 0),
        borderWidth: 3,
    });

    // Draw inner border
    page.drawRectangle({
        x: 70,
        y: 70,
        width: width - 140,
        height: height - 140,
        borderColor: rgb(0, 0, 0),
        borderWidth: 1,
    });

    // Title
    const title = 'CERTIFICATE OF COMPLETION';
    const titleSize = 28;
    const titleWidth = titleFont.widthOfTextAtSize(title, titleSize);
    page.drawText(title, {
        x: (width - titleWidth) / 2,
        y: height - 150,
        size: titleSize,
        font: titleFont,
        color: rgb(0, 0, 0),
    });

    // Subtitle
    const subtitle = 'This is to certify that';
    const subtitleSize = 16;
    const subtitleWidth = bodyFont.widthOfTextAtSize(subtitle, subtitleSize);
    page.drawText(subtitle, {
        x: (width - subtitleWidth) / 2,
        y: height - 250,
        size: subtitleSize,
        font: bodyFont,
        color: rgb(0, 0, 0),
    });

    // Placeholder for name (this will be replaced)
    const namePlaceholder = '{{NAME WILL BE PLACED HERE}}';
    const nameSize = 24;
    const nameWidth = titleFont.widthOfTextAtSize(namePlaceholder, nameSize);
    page.drawText(namePlaceholder, {
        x: (width - nameWidth) / 2,
        y: height - 320,
        size: nameSize,
        font: titleFont,
        color: rgb(0.2, 0.2, 0.8),
    });

    // Achievement text
    const achievement = 'has successfully completed the training program';
    const achievementSize = 16;
    const achievementWidth = bodyFont.widthOfTextAtSize(achievement, achievementSize);
    page.drawText(achievement, {
        x: (width - achievementWidth) / 2,
        y: height - 380,
        size: achievementSize,
        font: bodyFont,
        color: rgb(0, 0, 0),
    });

    // Course name
    const courseName = 'Advanced Web Development';
    const courseSize = 20;
    const courseWidth = titleFont.widthOfTextAtSize(courseName, courseSize);
    page.drawText(courseName, {
        x: (width - courseWidth) / 2,
        y: height - 420,
        size: courseSize,
        font: titleFont,
        color: rgb(0, 0, 0),
    });

    // Date
    const dateText = `Date: ${new Date().toLocaleDateString()}`;
    const dateSize = 14;
    page.drawText(dateText, {
        x: 100,
        y: 150,
        size: dateSize,
        font: bodyFont,
        color: rgb(0, 0, 0),
    });

    // Signature line
    page.drawLine({
        start: { x: width - 250, y: 150 },
        end: { x: width - 100, y: 150 },
        thickness: 1,
        color: rgb(0, 0, 0),
    });

    const signatureText = 'Authorized Signature';
    const signatureSize = 12;
    page.drawText(signatureText, {
        x: width - 220,
        y: 130,
        size: signatureSize,
        font: bodyFont,
        color: rgb(0, 0, 0),
    });

    // Save the PDF
    const pdfBytes = await pdfDoc.save();
    const pdfPath = path.join(__dirname, 'certificate-template.pdf');
    await fs.writeFile(pdfPath, pdfBytes);

    console.log(`Sample PDF template created: ${pdfPath}`);
    return pdfPath;
}

async function createSampleFiles() {
    try {
        console.log('Creating sample files...');

        const excelPath = await createSampleExcel();
        const pdfPath = await createSamplePDFTemplate();

        console.log('\nSample files created successfully!');
        console.log('Files created:');
        console.log(`- Excel file: ${excelPath}`);
        console.log(`- PDF template: ${pdfPath}`);
        console.log('\nYou can use these files to test the certificate generator.');

    } catch (error) {
        console.error('Error creating sample files:', error);
    }
}

// Run if called directly
if (require.main === module) {
    createSampleFiles();
}

module.exports = { createSampleFiles, createSampleExcel, createSamplePDFTemplate };
