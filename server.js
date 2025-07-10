const express = require('express');
const multer = require('multer');
const cors = require('cors');
const path = require('path');
const fs = require('fs-extra');
const XLSX = require('xlsx');
const PDFProcessor = require('./pdfProcessor');
const PPTXProcessor = require('./pptxProcessor');
const EmailService = require('./emailService');
const CanvaIntegration = require('./canvaIntegration');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.static('public'));

// Ensure directories exist
const ensureDirectories = async () => {
    await fs.ensureDir('./uploads');
    await fs.ensureDir('./templates');
    await fs.ensureDir('./generated');
    await fs.ensureDir('./public');
};

// Configure multer for file uploads
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        if (file.fieldname === 'excel') {
            cb(null, './uploads');
        } else if (file.fieldname === 'template') {
            cb(null, './templates');
        }
    },
    filename: (req, file, cb) => {
        const timestamp = Date.now();
        cb(null, `${timestamp}-${file.originalname}`);
    }
});

const upload = multer({
    storage: storage,
    fileFilter: (req, file, cb) => {
        if (file.fieldname === 'excel') {
            // Accept Excel files
            if (file.mimetype.includes('spreadsheet') ||
                file.originalname.endsWith('.xlsx') ||
                file.originalname.endsWith('.xls')) {
                cb(null, true);
            } else {
                cb(new Error('Only Excel files are allowed for excel field'));
            }
        } else if (file.fieldname === 'template') {
            // Accept PDF and PPTX files
            if (file.mimetype === 'application/pdf' || file.originalname.endsWith('.pdf') ||
                file.mimetype === 'application/vnd.openxmlformats-officedocument.presentationml.presentation' ||
                file.originalname.endsWith('.pptx')) {
                cb(null, true);
            } else {
                cb(new Error('Only PDF and PPTX files are allowed for template field'));
            }
        } else {
            cb(new Error('Unknown field'));
        }
    }
});

// Function to read Excel file and extract names and emails
const extractDataFromExcel = async (filePath) => {
    try {
        const workbook = XLSX.readFile(filePath);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet);

        // Look for name column (case insensitive)
        const nameColumns = ['name', 'Name', 'NAME', 'names', 'Names', 'NAMES'];
        let nameColumn = null;

        for (const col of nameColumns) {
            if (data.length > 0 && data[0].hasOwnProperty(col)) {
                nameColumn = col;
                break;
            }
        }

        if (!nameColumn) {
            throw new Error('No name column found. Please ensure your Excel file has a column named "name", "Name", or "NAME"');
        }

        // Look for email column (case insensitive)
        const emailColumns = ['email', 'Email', 'EMAIL', 'emails', 'Emails', 'EMAILS', 'e-mail', 'E-mail', 'E-MAIL'];
        let emailColumn = null;

        for (const col of emailColumns) {
            if (data.length > 0 && data[0].hasOwnProperty(col)) {
                emailColumn = col;
                break;
            }
        }

        const extractedData = data.map(row => ({
            name: row[nameColumn],
            email: emailColumn ? row[emailColumn] : null,
            rawData: row
        })).filter(item => item.name && item.name.trim());

        return {
            data: extractedData,
            hasEmails: !!emailColumn,
            emailColumn: emailColumn
        };
    } catch (error) {
        throw new Error(`Error reading Excel file: ${error.message}`);
    }
};

// Initialize processors
const pdfProcessor = new PDFProcessor();
const pptxProcessor = new PPTXProcessor();
const emailService = new EmailService();
const canvaIntegration = new CanvaIntegration();

// Configure email service from environment variables
emailService.configureFromEnv();

// Function to generate certificates using appropriate processor
const generateCertificates = async (templatePath, names, outputDir, options = {}) => {
    try {
        const ext = path.extname(templatePath).toLowerCase();

        if (ext === '.pptx') {
            // Use PPTX processor
            return await pptxProcessor.generateCertificates(templatePath, names, outputDir, options);
        } else if (ext === '.pdf') {
            // Use PDF processor
            return await pdfProcessor.generateCertificates(templatePath, names, outputDir, {
                fontSize: 24,
                ...options
            });
        } else {
            throw new Error('Unsupported template format. Use .pdf or .pptx files.');
        }
    } catch (error) {
        console.error('Error in generateCertificates:', error);
        return [];
    }
};

// Routes
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Upload and process files
app.post('/upload', upload.fields([
    { name: 'excel', maxCount: 1 },
    { name: 'template', maxCount: 1 }
]), async (req, res) => {
    try {
        if (!req.files || !req.files.excel || !req.files.template) {
            return res.status(400).json({
                error: 'Both Excel file and PDF template are required'
            });
        }

        const excelFile = req.files.excel[0];
        const templateFile = req.files.template[0];

        // Extract data from Excel
        const excelData = await extractDataFromExcel(excelFile.path);

        if (excelData.data.length === 0) {
            return res.status(400).json({
                error: 'No names found in the Excel file'
            });
        }

        const names = excelData.data.map(item => item.name);

        // Create output directory for this batch
        const timestamp = Date.now();
        const outputDir = path.join('./generated', `batch-${timestamp}`);
        await fs.ensureDir(outputDir);

        // Generate certificates with additional options
        const additionalOptions = {
            date: req.body.date || new Date().toLocaleDateString(),
            course: req.body.course || '',
            instructor: req.body.instructor || '',
            organization: req.body.organization || ''
        };

        const generatedFiles = await generateCertificates(templateFile.path, names, outputDir, additionalOptions);

        res.json({
            success: true,
            message: `Generated ${generatedFiles.length} certificates`,
            names: names,
            generatedFiles: generatedFiles,
            batchId: `batch-${timestamp}`,
            hasEmails: excelData.hasEmails,
            emailColumn: excelData.emailColumn,
            excelData: excelData.data
        });

    } catch (error) {
        console.error('Error processing files:', error);
        res.status(500).json({
            error: error.message || 'Internal server error'
        });
    }
});

// Download generated certificates
app.get('/download/:batchId/:filename', (req, res) => {
    const { batchId, filename } = req.params;
    const filePath = path.join('./generated', batchId, filename);

    if (fs.existsSync(filePath)) {
        res.download(filePath);
    } else {
        res.status(404).json({ error: 'File not found' });
    }
});

// List generated batches
app.get('/batches', async (req, res) => {
    try {
        const generatedDir = './generated';
        const batches = await fs.readdir(generatedDir);
        const batchInfo = [];

        for (const batch of batches) {
            const batchPath = path.join(generatedDir, batch);
            const stats = await fs.stat(batchPath);
            if (stats.isDirectory()) {
                const files = await fs.readdir(batchPath);
                batchInfo.push({
                    id: batch,
                    created: stats.birthtime,
                    fileCount: files.length,
                    files: files
                });
            }
        }

        res.json(batchInfo);
    } catch (error) {
        res.status(500).json({ error: 'Error listing batches' });
    }
});

// Delete a batch
app.delete('/batches/:batchId', async (req, res) => {
    try {
        const { batchId } = req.params;
        const batchPath = path.join('./generated', batchId);

        if (await fs.pathExists(batchPath)) {
            await fs.remove(batchPath);
            res.json({ success: true, message: 'Batch deleted successfully' });
        } else {
            res.status(404).json({ error: 'Batch not found' });
        }
    } catch (error) {
        res.status(500).json({ error: 'Error deleting batch' });
    }
});

// Get batch details
app.get('/batches/:batchId', async (req, res) => {
    try {
        const { batchId } = req.params;
        const batchPath = path.join('./generated', batchId);

        if (await fs.pathExists(batchPath)) {
            const files = await fs.readdir(batchPath);
            const stats = await fs.stat(batchPath);

            res.json({
                id: batchId,
                created: stats.birthtime,
                fileCount: files.length,
                files: files
            });
        } else {
            res.status(404).json({ error: 'Batch not found' });
        }
    } catch (error) {
        res.status(500).json({ error: 'Error getting batch details' });
    }
});

// Check Python environment for PPTX processing
app.get('/check-pptx-support', async (req, res) => {
    try {
        const envCheck = await pptxProcessor.checkPythonEnvironment();
        res.json(envCheck);
    } catch (error) {
        res.status(500).json({
            available: false,
            message: 'Error checking Python environment',
            error: error.message
        });
    }
});

// Get supported placeholders for PPTX templates
app.get('/pptx-placeholders', (req, res) => {
    try {
        const placeholders = pptxProcessor.getSupportedPlaceholders();
        res.json({
            success: true,
            placeholders: placeholders
        });
    } catch (error) {
        res.status(500).json({
            success: false,
            error: error.message
        });
    }
});

// Validate template file
app.post('/validate-template', upload.single('template'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'No template file provided' });
        }

        const ext = path.extname(req.file.originalname).toLowerCase();
        let validation;

        if (ext === '.pptx') {
            validation = await pptxProcessor.validateTemplate(req.file.path);
        } else if (ext === '.pdf') {
            // Basic PDF validation
            validation = {
                valid: true,
                message: 'PDF template accepted'
            };
        } else {
            validation = {
                valid: false,
                message: 'Unsupported file format. Use .pdf or .pptx files.'
            };
        }

        res.json({
            success: validation.valid,
            message: validation.message,
            fileType: ext,
            fileName: req.file.originalname
        });

    } catch (error) {
        res.status(500).json({
            success: false,
            error: error.message
        });
    }
});

// Configure email service
app.post('/configure-email', async (req, res) => {
    try {
        const { service, user, password } = req.body;

        if (!user || !password) {
            return res.status(400).json({
                success: false,
                error: 'Email and password are required'
            });
        }

        // Validate email format
        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        if (!emailRegex.test(user)) {
            return res.status(400).json({
                success: false,
                error: 'Please enter a valid email address'
            });
        }

        const result = emailService.configure({
            service: service || 'gmail',
            user: user,
            password: password
        });

        if (result.success) {
            // Test the connection immediately after configuration
            const testResult = await emailService.testConnection();
            if (!testResult.success) {
                return res.json({
                    success: false,
                    error: `Configuration saved but connection test failed: ${testResult.error}`
                });
            }
        }

        res.json(result);
    } catch (error) {
        res.status(500).json({
            success: false,
            error: error.message
        });
    }
});

// Test email configuration
app.get('/test-email', async (req, res) => {
    try {
        const result = await emailService.testConnection();
        res.json(result);
    } catch (error) {
        res.status(500).json({
            success: false,
            error: error.message
        });
    }
});

// Get email service status
app.get('/email-status', (req, res) => {
    try {
        const status = emailService.getStatus();
        res.json(status);
    } catch (error) {
        res.status(500).json({
            success: false,
            error: error.message
        });
    }
});

// Get email templates
app.get('/email-templates', (req, res) => {
    try {
        const templates = emailService.getEmailTemplates();
        res.json({
            success: true,
            templates: templates
        });
    } catch (error) {
        res.status(500).json({
            success: false,
            error: error.message
        });
    }
});

// Send certificates via email
app.post('/send-certificates', async (req, res) => {
    try {
        const {
            batchId,
            emailConfig,
            recipients
        } = req.body;

        if (!batchId || !recipients || recipients.length === 0) {
            return res.status(400).json({
                error: 'Batch ID and recipients are required'
            });
        }

        // Prepare recipients with certificate paths
        const recipientsWithPaths = recipients.map(recipient => ({
            ...recipient,
            certificatePath: path.join('./generated', batchId, recipient.filename)
        }));

        // Send emails
        const result = await emailService.sendBulkCertificates(recipientsWithPaths, emailConfig);

        res.json({
            success: result.success,
            sent: result.sent,
            failed: result.failed,
            results: result.results,
            errors: result.errors
        });

    } catch (error) {
        console.error('Error sending certificates:', error);
        res.status(500).json({
            error: error.message || 'Internal server error'
        });
    }
});

// Canva Integration Endpoints

// Get Canva templates
app.get('/canva/templates', async (req, res) => {
    try {
        const { category, search, limit, offset } = req.query;

        const result = await canvaIntegration.searchTemplates({
            query: search,
            category: category,
            limit: parseInt(limit) || 20,
            offset: parseInt(offset) || 0
        });

        res.json(result);
    } catch (error) {
        res.status(500).json({
            success: false,
            error: error.message
        });
    }
});

// Get Canva template categories
app.get('/canva/categories', (req, res) => {
    try {
        const categories = canvaIntegration.getCategories();
        res.json({
            success: true,
            categories: categories
        });
    } catch (error) {
        res.status(500).json({
            success: false,
            error: error.message
        });
    }
});

// Get specific Canva template details
app.get('/canva/templates/:templateId', (req, res) => {
    try {
        const { templateId } = req.params;
        const result = canvaIntegration.getTemplateDetails(templateId);
        res.json(result);
    } catch (error) {
        res.status(500).json({
            success: false,
            error: error.message
        });
    }
});

// Download Canva template
app.post('/canva/download/:templateId', async (req, res) => {
    try {
        const { templateId } = req.params;
        const timestamp = Date.now();
        const outputPath = path.join('./templates', `canva-${templateId}-${timestamp}.pptx`);

        const result = await canvaIntegration.downloadTemplate(templateId, outputPath);

        if (result.success) {
            res.json({
                success: true,
                templatePath: outputPath,
                template: result.template
            });
        } else {
            res.status(400).json(result);
        }
    } catch (error) {
        res.status(500).json({
            success: false,
            error: error.message
        });
    }
});

// Check Canva configuration status
app.get('/canva/status', (req, res) => {
    try {
        const status = canvaIntegration.checkConfiguration();
        res.json(status);
    } catch (error) {
        res.status(500).json({
            success: false,
            error: error.message
        });
    }
});

// Initialize and start server
const startServer = async () => {
    await ensureDirectories();
    app.listen(PORT, () => {
        console.log(`Certificate Generator Server running on http://localhost:${PORT}`);
    });
};

startServer().catch(console.error);
