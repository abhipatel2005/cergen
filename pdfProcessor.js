const { PDFDocument, rgb, StandardFonts } = require('pdf-lib');
const fs = require('fs-extra');

class PDFProcessor {
    constructor() {
        this.defaultFont = StandardFonts.Helvetica;
        this.defaultFontSize = 24;
        this.defaultColor = rgb(0, 0, 0);

        // Common placeholder patterns and their likely positions
        this.placeholderPatterns = [
            { pattern: '{{NAME WILL BE PLACED HERE}}', replacement: 'name' },
            { pattern: '{{name}}', replacement: 'name' },
            { pattern: '{{NAME}}', replacement: 'name' },
            { pattern: '[NAME]', replacement: 'name' },
            { pattern: '_____', replacement: 'name' }
        ];
    }

    /**
     * Process PDF template and replace placeholders with actual values
     * @param {string} templatePath - Path to the PDF template
     * @param {Object} data - Data object with placeholder values (e.g., {name: "John Doe"})
     * @param {string} outputPath - Path where the processed PDF should be saved
     * @param {Object} options - Additional options for text styling
     */
    async processPDFTemplate(templatePath, data, outputPath, options = {}) {
        try {
            // Read the template PDF
            const templateBytes = await fs.readFile(templatePath);
            const pdfDoc = await PDFDocument.load(templateBytes);

            // Get font
            const font = await pdfDoc.embedFont(options.font || this.defaultFont);

            // Process each page
            const pages = pdfDoc.getPages();

            for (const page of pages) {
                await this.processPage(page, data, font, options);
            }

            // Save the processed PDF
            const pdfBytes = await pdfDoc.save();
            await fs.writeFile(outputPath, pdfBytes);

            return { success: true, outputPath };
        } catch (error) {
            throw new Error(`Error processing PDF template: ${error.message}`);
        }
    }

    /**
     * Process a single page and replace placeholders
     * @param {Object} page - PDF page object
     * @param {Object} data - Data object with placeholder values
     * @param {Object} font - PDF font object
     * @param {Object} options - Styling options
     */
    async processPage(page, data, font, options) {
        const { width, height } = page.getSize();

        if (data.name) {
            // Try to find and replace placeholder text intelligently
            const placeholderInfo = this.findPlaceholderPosition(width, height);

            const fontSize = options.fontSize || placeholderInfo.fontSize || this.defaultFontSize;
            const color = options.color || this.defaultColor;

            // Calculate text dimensions
            const textWidth = font.widthOfTextAtSize(data.name, fontSize);
            const textHeight = font.heightAtSize(fontSize);

            // Use intelligent positioning
            const x = options.x !== undefined ? options.x : placeholderInfo.x - (textWidth / 2);
            const y = options.y !== undefined ? options.y : placeholderInfo.y;

            // Cover the placeholder area with a white rectangle to "erase" it
            if (placeholderInfo.coverArea) {
                page.drawRectangle({
                    x: placeholderInfo.coverArea.x,
                    y: placeholderInfo.coverArea.y,
                    width: placeholderInfo.coverArea.width,
                    height: placeholderInfo.coverArea.height,
                    color: rgb(1, 1, 1), // White background
                });
            }

            // Draw the replacement text
            page.drawText(data.name, {
                x: x,
                y: y,
                size: fontSize,
                font: font,
                color: color,
            });
        }
    }

    /**
     * Find the likely position of placeholder text based on common patterns
     * @param {number} width - Page width
     * @param {number} height - Page height
     * @returns {Object} Position and styling information
     */
    findPlaceholderPosition(width, height) {
        // This is a heuristic approach since we can't extract existing text
        // We'll assume the placeholder is in a common certificate position

        // Common certificate name positions:
        // 1. Center of page (most common)
        // 2. Slightly above center
        // 3. In the upper third of the page

        const centerX = width / 2;
        const centerY = height / 2;

        // Assume placeholder text area (approximate)
        const placeholderWidth = 300; // Approximate width of placeholder text
        const placeholderHeight = 30; // Approximate height

        return {
            x: centerX,
            y: centerY - 20, // Slightly below center for better visual balance
            fontSize: 24,
            coverArea: {
                x: centerX - (placeholderWidth / 2),
                y: centerY - 25,
                width: placeholderWidth,
                height: placeholderHeight
            }
        };
    }

    /**
     * Advanced method to find and replace text placeholders in PDF
     * This is a simplified version - real implementation would need to parse PDF content
     * @param {string} templatePath - Path to PDF template
     * @param {Object} replacements - Object with placeholder -> value mappings
     * @param {string} outputPath - Output file path
     */
    async replaceTextInPDF(templatePath, replacements, outputPath) {
        try {
            // Read the template PDF
            const templateBytes = await fs.readFile(templatePath);
            const pdfDoc = await PDFDocument.load(templateBytes);

            // Get font
            const font = await pdfDoc.embedFont(StandardFonts.Helvetica);

            // Process each page
            const pages = pdfDoc.getPages();

            for (const page of pages) {
                const { width, height } = page.getSize();

                // For each replacement
                Object.entries(replacements).forEach(([placeholder, value]) => {
                    if (placeholder === 'name' && value) {
                        // Calculate position for name (center of page)
                        const fontSize = 24;
                        const textWidth = font.widthOfTextAtSize(value, fontSize);
                        const x = (width - textWidth) / 2;
                        const y = height / 2;

                        // Draw the replacement text
                        page.drawText(value, {
                            x: x,
                            y: y,
                            size: fontSize,
                            font: font,
                            color: rgb(0, 0, 0),
                        });
                    }
                });
            }

            // Save the processed PDF
            const pdfBytes = await pdfDoc.save();
            await fs.writeFile(outputPath, pdfBytes);

            return { success: true, outputPath };
        } catch (error) {
            throw new Error(`Error replacing text in PDF: ${error.message}`);
        }
    }

    /**
     * Enhanced method to process PDF with better placeholder detection
     * @param {string} templatePath - Path to PDF template
     * @param {Object} data - Data object with values
     * @param {string} outputPath - Output path
     * @param {Object} options - Processing options
     */
    async processAdvancedPDFTemplate(templatePath, data, outputPath, options = {}) {
        try {
            // Read the template PDF
            const templateBytes = await fs.readFile(templatePath);
            const pdfDoc = await PDFDocument.load(templateBytes);

            // Try to embed multiple fonts to match existing styles
            const fonts = await this.embedCommonFonts(pdfDoc);

            // Analyze the template to get better positioning
            const templateAnalysis = await this.analyzeTemplate(pdfDoc);

            // Process each page
            const pages = pdfDoc.getPages();

            for (const page of pages) {
                await this.processAdvancedPage(page, data, fonts, { ...options, templateAnalysis });
            }

            // Save the processed PDF
            const pdfBytes = await pdfDoc.save();
            await fs.writeFile(outputPath, pdfBytes);

            return { success: true, outputPath };
        } catch (error) {
            throw new Error(`Error processing PDF template: ${error.message}`);
        }
    }

    /**
     * Analyze the PDF template to understand its structure
     * @param {PDFDocument} pdfDoc - PDF document
     * @returns {Object} Analysis results
     */
    async analyzeTemplate(pdfDoc) {
        const pages = pdfDoc.getPages();
        const firstPage = pages[0];
        const { width, height } = firstPage.getSize();

        // Heuristic analysis based on common certificate layouts
        const analysis = {
            pageWidth: width,
            pageHeight: height,
            likelyNamePositions: [],
            suggestedFontSize: 24,
            suggestedFont: 'helveticaBold'
        };

        // Common certificate name positions based on standard layouts
        if (width > 500 && height > 600) {
            // Standard certificate size
            analysis.likelyNamePositions = [
                { x: width / 2, y: height * 0.45, confidence: 0.9 }, // Center-low
                { x: width / 2, y: height * 0.5, confidence: 0.8 },  // Center
                { x: width / 2, y: height * 0.55, confidence: 0.7 }, // Center-high
            ];
            analysis.suggestedFontSize = Math.min(32, width / 20);
        } else {
            // Smaller or non-standard size
            analysis.likelyNamePositions = [
                { x: width / 2, y: height / 2, confidence: 0.8 },
            ];
            analysis.suggestedFontSize = Math.min(24, width / 25);
        }

        return analysis;
    }

    /**
     * Embed common fonts that might match the template
     * @param {PDFDocument} pdfDoc - PDF document
     * @returns {Object} Object containing embedded fonts
     */
    async embedCommonFonts(pdfDoc) {
        return {
            helvetica: await pdfDoc.embedFont(StandardFonts.Helvetica),
            helveticaBold: await pdfDoc.embedFont(StandardFonts.HelveticaBold),
            timesRoman: await pdfDoc.embedFont(StandardFonts.TimesRoman),
            timesRomanBold: await pdfDoc.embedFont(StandardFonts.TimesRomanBold),
            courier: await pdfDoc.embedFont(StandardFonts.Courier),
            courierBold: await pdfDoc.embedFont(StandardFonts.CourierBold)
        };
    }

    /**
     * Advanced page processing with better font matching
     * @param {Object} page - PDF page
     * @param {Object} data - Data to insert
     * @param {Object} fonts - Available fonts
     * @param {Object} options - Options
     */
    async processAdvancedPage(page, data, fonts, options) {
        const { width, height } = page.getSize();

        if (data.name) {
            // Use template analysis if available
            const analysis = options.templateAnalysis;
            let positions;

            if (analysis && analysis.likelyNamePositions.length > 0) {
                // Use analyzed positions
                positions = analysis.likelyNamePositions.map(pos => ({
                    x: pos.x,
                    y: pos.y,
                    coverArea: {
                        x: pos.x - 250,
                        y: pos.y - 20,
                        width: 500,
                        height: 45
                    }
                }));
            } else {
                // Fallback to heuristic positions
                positions = this.getMultiplePositions(width, height);
            }

            // Cover potential placeholder areas with white rectangles
            for (const position of positions) {
                page.drawRectangle({
                    x: position.coverArea.x,
                    y: position.coverArea.y,
                    width: position.coverArea.width,
                    height: position.coverArea.height,
                    color: rgb(1, 1, 1), // White to cover existing text
                });
            }

            // Choose the best font and size
            const font = this.selectBestFont(fonts, options);
            const fontSize = options.fontSize ||
                (analysis ? analysis.suggestedFontSize : 24);
            const color = options.color || rgb(0, 0, 0);

            // Calculate text dimensions
            const textWidth = font.widthOfTextAtSize(data.name, fontSize);

            // Use the primary position (highest confidence or first)
            const primaryPosition = positions[0];
            const x = primaryPosition.x - (textWidth / 2);
            const y = primaryPosition.y;

            // Draw the replacement text
            page.drawText(data.name, {
                x: x,
                y: y,
                size: fontSize,
                font: font,
                color: color,
            });
        }
    }

    /**
     * Get multiple potential positions for placeholders
     * @param {number} width - Page width
     * @param {number} height - Page height
     * @returns {Array} Array of position objects
     */
    getMultiplePositions(width, height) {
        const centerX = width / 2;
        const centerY = height / 2;

        return [
            // Primary position - center
            {
                x: centerX,
                y: centerY - 20,
                coverArea: {
                    x: centerX - 200,
                    y: centerY - 35,
                    width: 400,
                    height: 40
                }
            },
            // Secondary position - slightly higher
            {
                x: centerX,
                y: centerY + 20,
                coverArea: {
                    x: centerX - 200,
                    y: centerY + 5,
                    width: 400,
                    height: 40
                }
            },
            // Tertiary position - lower third
            {
                x: centerX,
                y: height * 0.4,
                coverArea: {
                    x: centerX - 200,
                    y: height * 0.4 - 15,
                    width: 400,
                    height: 40
                }
            }
        ];
    }

    /**
     * Select the best font based on certificate style
     * @param {Object} fonts - Available fonts
     * @param {Object} options - Options
     * @returns {Object} Selected font
     */
    selectBestFont(fonts, options) {
        if (options.font) return options.font;

        // For certificates, bold fonts often look better
        return fonts.helveticaBold || fonts.timesRomanBold || fonts.helvetica;
    }

    /**
     * Generate multiple certificates from a template
     * @param {string} templatePath - Path to PDF template
     * @param {Array} names - Array of names
     * @param {string} outputDir - Directory to save generated certificates
     * @param {Object} options - Styling options
     */
    async generateCertificates(templatePath, names, outputDir, options = {}) {
        const generatedFiles = [];

        // Ensure output directory exists
        await fs.ensureDir(outputDir);

        for (let i = 0; i < names.length; i++) {
            const name = names[i];
            try {
                // Create safe filename
                const safeFileName = name.replace(/[^a-zA-Z0-9\s]/g, '').replace(/\s+/g, '_');
                const outputPath = `${outputDir}/certificate-${safeFileName}.pdf`;

                // Use the advanced processing method
                await this.processAdvancedPDFTemplate(templatePath, { name }, outputPath, options);

                generatedFiles.push({
                    name: name,
                    filename: `certificate-${safeFileName}.pdf`,
                    path: outputPath
                });
            } catch (error) {
                console.error(`Error generating certificate for ${name}:`, error);
                // Continue with other names even if one fails
            }
        }

        return generatedFiles;
    }

    /**
     * Get PDF information
     * @param {string} pdfPath - Path to PDF file
     */
    async getPDFInfo(pdfPath) {
        try {
            const pdfBytes = await fs.readFile(pdfPath);
            const pdfDoc = await PDFDocument.load(pdfBytes);

            const pages = pdfDoc.getPages();
            const pageCount = pages.length;

            const firstPage = pages[0];
            const { width, height } = firstPage.getSize();

            return {
                pageCount,
                dimensions: { width, height },
                title: pdfDoc.getTitle() || 'Untitled',
                author: pdfDoc.getAuthor() || 'Unknown'
            };
        } catch (error) {
            throw new Error(`Error reading PDF info: ${error.message}`);
        }
    }
}

module.exports = PDFProcessor;
