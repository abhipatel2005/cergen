# Certificate Generator

A web-based system to generate personalized certificates from Excel data using PowerPoint (PPTX) or PDF templates.

## 🎉 NEW: PPTX Template Support!

**True placeholder replacement** with PowerPoint templates - the best way to create professional certificates!

## Features

- **🆕 PPTX Template Processing**: TRUE text replacement with {{placeholders}} in PowerPoint templates
- **📧 Email Delivery**: Automatically send certificates to recipients' email addresses with custom messages
- **Excel File Processing**: Upload Excel files (.xlsx, .xls) with "name" and "email" columns
- **PDF Template Processing**: Legacy support with overlay method for PDF templates
- **Multiple Placeholders**: Support for {{name}}, {{date}}, {{course}}, {{instructor}}, {{organization}}
- **Batch Generation**: Generate multiple certificates at once
- **File Management**: Organize generated certificates in folders
- **Download System**: Download individual certificates or view all batches
- **Email Templates**: Pre-built email templates with professional formatting
- **Web Interface**: User-friendly web interface for easy operation

## Installation

### Node.js Dependencies

1. Clone or download this repository
2. Install Node.js dependencies:
   ```bash
   npm install
   ```

### Python Dependencies (for PPTX support)

3. Install Python dependencies:

   ```bash
   pip install -r requirements.txt
   ```

   Or manually:

   ```bash
   pip install python-pptx comtypes
   ```

**Note**: PPTX processing requires Python and works best on Windows (uses COM automation for PDF conversion).

## Usage

1. Start the server:

   ```bash
   npm start
   ```

2. Open your browser and go to: `http://localhost:3000`

3. Upload your files:

   - **Excel File**: Must contain a "name" column. Optionally include an "email" column for automatic email delivery
   - **PPTX Template** (Recommended): PowerPoint template with {{placeholders}} for true text replacement
   - **PDF Template** (Legacy): PDF template where names will be overlaid

4. For PPTX templates, optionally fill in additional information (course, instructor, organization, date)

5. Click "Generate Certificates" to create personalized certificates

6. **NEW!** If email addresses are present, configure email settings and send certificates automatically

7. Download the generated certificates individually or send them via email

## Sample Files

The project includes sample files for testing:

- `sample-names.xlsx`: Sample Excel file with 8 names
- `certificate-template.pdf`: Sample PDF certificate template
- `certificate-template.pptx`: Sample PPTX certificate template (create with Python script)

To create sample files:

```bash
# Create Excel and PDF samples
node createSampleFiles.js

# Create PPTX sample template
python create_sample_pptx.py
```

## 🎯 PPTX Templates (Recommended)

### Why PPTX is Better:

- ✅ **True Text Replacement**: Actually replaces {{placeholders}} instead of overlaying
- ✅ **Font Preservation**: Maintains original fonts, sizes, and styling
- ✅ **Layout Preservation**: Keeps exact positioning and formatting
- ✅ **Multiple Placeholders**: Support for name, date, course, instructor, organization
- ✅ **Easy Creation**: Create templates in familiar PowerPoint interface
- ✅ **Rich Formatting**: Full support for colors, gradients, images, shapes

### Creating PPTX Templates:

1. Open PowerPoint
2. Design your certificate layout
3. Add text placeholders:
   - `{{name}}` - Person's name (original case)
   - `{{NAME}}` - Person's name (uppercase)
   - `{{Name}}` - Person's name (title case)
   - `{{date}}` - Date
   - `{{course}}` - Course name
   - `{{instructor}}` - Instructor name
   - `{{organization}}` - Organization name
4. Save as .pptx file
5. Upload and use!

### Example PPTX Content:

```
CERTIFICATE OF COMPLETION

This certifies that {{name}} has successfully
completed the course {{course}} on {{date}}.

Instructor: {{instructor}}
{{organization}}
```

## 📧 Email Delivery Feature

### Setup Email Configuration:

1. **Create .env file** (copy from .env.example):

   ```bash
   cp .env.example .env
   ```

2. **Configure your email credentials** in .env:

   ```env
   EMAIL_SERVICE=gmail
   EMAIL_USER=your-email@gmail.com
   EMAIL_PASS=your-app-password
   ```

3. **For Gmail users**:
   - Enable 2-Factor Authentication
   - Generate App Password: https://myaccount.google.com/apppasswords
   - Use the App Password (not your regular password)

### Excel File Requirements for Email:

Your Excel file should have both columns:
| name | email | department |
|------|-------|------------|
| John Doe | john.doe@example.com | Engineering |
| Jane Smith | jane.smith@example.com | Marketing |

### Email Features:

- ✅ **Automatic Detection**: System detects email column automatically
- ✅ **Custom Messages**: Write personalized email content with placeholders
- ✅ **Professional Templates**: Pre-built email templates
- ✅ **Bulk Sending**: Send to all recipients with rate limiting
- ✅ **Error Handling**: Detailed reports of successful/failed deliveries
- ✅ **HTML Emails**: Rich formatting with certificate details

### Email Templates Available:

1. **Completion Template**: Congratulatory tone for course completion
2. **Achievement Template**: Celebratory tone for achievements
3. **Professional Template**: Formal business communication

### Supported Email Providers:

- Gmail (recommended)
- Outlook/Hotmail
- Yahoo Mail
- Custom SMTP servers

## File Structure

```
certificate-generator/
├── server.js              # Main server file
├── pdfProcessor.js         # PDF processing module
├── pptxProcessor.js        # PPTX processing module (Node.js)
├── pptx_processor.py       # PPTX processing script (Python)
├── emailService.js         # Email delivery service
├── createSampleFiles.js    # Script to create sample files
├── create_sample_pptx.py   # Script to create PPTX template
├── .env.example           # Email configuration template
├── requirements.txt       # Python dependencies
├── public/
│   └── index.html         # Web interface with email functionality
├── uploads/               # Uploaded files storage
├── templates/             # Template files storage
├── generated/             # Generated certificates storage
└── sample files...
```

## API Endpoints

### Certificate Generation

- `GET /` - Web interface
- `POST /upload` - Upload Excel and template files, generate certificates
- `GET /download/:batchId/:filename` - Download specific certificate
- `GET /batches` - List all generated batches
- `GET /batches/:batchId` - Get batch details
- `DELETE /batches/:batchId` - Delete a batch

### Email Functionality

- `POST /configure-email` - Configure email service credentials
- `GET /test-email` - Test email configuration
- `GET /email-status` - Get email service status
- `GET /email-templates` - Get available email templates
- `POST /send-certificates` - Send certificates via email

### Template Support

- `GET /check-pptx-support` - Check Python environment for PPTX processing
- `GET /pptx-placeholders` - Get supported PPTX placeholders
- `POST /validate-template` - Validate uploaded template file

## Excel File Requirements

Your Excel file must have a column with one of these names (case-insensitive):

- `name`
- `Name`
- `NAME`
- `names`
- `Names`
- `NAMES`

Example Excel structure:
| name | email | department |
|-------------|-------------------|------------|
| John Doe | john@example.com | Engineering|
| Jane Smith | jane@example.com | Marketing |

## PDF Template Notes

The system now includes enhanced placeholder replacement features:

### ✅ **Enhanced Features:**

- **Smart Placeholder Detection**: Automatically covers existing placeholder text with white rectangles
- **Multiple Font Support**: Uses Helvetica Bold and other fonts to match certificate styles
- **Intelligent Positioning**: Analyzes PDF dimensions to determine optimal name placement
- **Template Analysis**: Automatically adjusts font size based on certificate dimensions

### **How It Works:**

1. **Covers Placeholders**: Places white rectangles over likely placeholder areas (like "{{NAME WILL BE PLACED HERE}}")
2. **Font Matching**: Uses bold fonts (Helvetica Bold, Times Roman Bold) for professional appearance
3. **Smart Positioning**: Analyzes certificate layout to place names in optimal positions
4. **Multiple Coverage Areas**: Covers multiple potential placeholder positions to ensure clean replacement

### **Supported Placeholder Patterns:**

- `{{NAME WILL BE PLACED HERE}}`
- `{{name}}` or `{{NAME}}`
- `[NAME]`
- Underscores `_____`
- Any text in the center area of certificates

## Customization

### Changing Text Placement

Edit `pdfProcessor.js` and modify the `processPage` method to change where names are placed on certificates.

### Styling Options

You can customize:

- Font size
- Font family
- Text color
- Text position (x, y coordinates)
- Multiple text fields

### Adding More Data Fields

To process more than just names:

1. Update the Excel processing to extract additional columns
2. Modify the PDF processor to handle multiple placeholders
3. Update the template processing logic

## Dependencies

- **express**: Web server framework
- **multer**: File upload handling
- **xlsx**: Excel file processing
- **pdf-lib**: PDF manipulation
- **cors**: Cross-origin resource sharing
- **fs-extra**: Enhanced file system operations

## Development

For development with auto-restart:

```bash
npm run dev
```

## Troubleshooting

### Common Issues

1. **"No name column found"**: Ensure your Excel file has a column named "name" (case-insensitive)
2. **PDF not generating**: Check that the PDF template is valid and not corrupted
3. **Server not starting**: Make sure port 3000 is available or set a different PORT environment variable

### Error Logs

Check the server console for detailed error messages when processing fails.

## License

ISC License

## Contributing

Feel free to submit issues and enhancement requests!
