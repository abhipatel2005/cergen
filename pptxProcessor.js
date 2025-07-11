const { spawn } = require('child_process');
const fs = require('fs-extra');
const path = require('path');

class PPTXProcessor {
    constructor() {
        this.pythonScript = path.join(__dirname, 'pptx_processor.py');
    }

    /**
     * Check if Python and required packages are available
     */
    async checkPythonEnvironment() {
        return new Promise((resolve) => {
            const python = spawn('python', ['--version']);

            python.on('close', (code) => {
                if (code === 0) {
                    // Check if required packages are installed
                    const checkPackages = spawn('python', ['-c', 'import pptx, comtypes; print("OK")']);

                    checkPackages.on('close', (packageCode) => {
                        resolve({
                            available: packageCode === 0,
                            message: packageCode === 0 ? 'Python environment ready' : 'Required Python packages not installed'
                        });
                    });
                } else {
                    resolve({
                        available: false,
                        message: 'Python not found'
                    });
                }
            });
        });
    }

    /**
     * Process PPTX templates and generate certificates
     * @param {string} templatePath - Path to PPTX template
     * @param {Array} data - Array of data objects or names
     * @param {string} outputDir - Output directory
     * @param {Object} options - Additional options
     */
    async generateCertificates(templatePath, data, outputDir, options = {}) {
        try {
            // Ensure output directory exists
            await fs.ensureDir(outputDir);

            // Handle both old format (array of names) and new format (array of objects)
            let processedData;
            if (Array.isArray(data) && typeof data[0] === 'string') {
                // Old format: array of names
                processedData = data.map(name => ({ name }));
            } else {
                // New format: array of objects
                processedData = data;
            }

            // Prepare additional data
            const additionalData = {
                date: options.date || new Date().toLocaleDateString(),
                course: options.course || '',
                instructor: options.instructor || '',
                organization: options.organization || '',
                fieldMappings: options.fieldMappings || {}
            };

            console.log('PPTX Processor - processedData sample:', processedData[0]);
            console.log('PPTX Processor - additionalData:', additionalData);

            // Call Python script
            const result = await this.callPythonProcessor(
                templatePath,
                outputDir,
                processedData,
                additionalData
            );

            return result;

        } catch (error) {
            throw new Error(`Error processing PPTX templates: ${error.message}`);
        }
    }

    /**
     * Call the Python processor script
     * @param {string} templatePath - Template path
     * @param {string} outputDir - Output directory
     * @param {Array} processedData - Processed data array
     * @param {Object} additionalData - Additional data
     */
    async callPythonProcessor(templatePath, outputDir, processedData, additionalData) {
        return new Promise((resolve, reject) => {
            const args = [
                this.pythonScript,
                '--template', templatePath,
                '--output-dir', outputDir,
                '--data', JSON.stringify(processedData),
                '--additional', JSON.stringify(additionalData)
            ];

            const python = spawn('python', args);
            let stdout = '';
            let stderr = '';

            python.stdout.on('data', (data) => {
                stdout += data.toString();
            });

            python.stderr.on('data', (data) => {
                stderr += data.toString();
            });

            python.on('close', (code) => {
                if (code === 0) {
                    try {
                        const result = JSON.parse(stdout);
                        if (result.success) {
                            resolve(result.results);
                        } else {
                            reject(new Error(result.error));
                        }
                    } catch (parseError) {
                        reject(new Error(`Failed to parse Python output: ${parseError.message}`));
                    }
                } else {
                    reject(new Error(`Python script failed with code ${code}: ${stderr}`));
                }
            });

            python.on('error', (error) => {
                reject(new Error(`Failed to start Python process: ${error.message}`));
            });
        });
    }

    /**
     * Create a sample PPTX template
     * @param {string} outputPath - Where to save the template
     */
    async createSampleTemplate(outputPath) {
        // This would require a more complex implementation
        // For now, we'll provide instructions for manual creation
        const instructions = `
To create a PPTX template:

1. Open PowerPoint
2. Create a certificate design
3. Use these placeholders in text boxes:
   - {{name}} or {{NAME}} - for the person's name
   - {{date}} - for the date
   - {{course}} - for course name
   - {{instructor}} - for instructor name
   - {{organization}} - for organization name

4. Save as .pptx file

Example placeholders:
"This certifies that {{name}} has successfully completed {{course}} on {{date}}"
        `;

        await fs.writeFile(outputPath.replace('.pptx', '_instructions.txt'), instructions);

        return {
            success: true,
            message: 'Instructions created. Please create PPTX template manually.',
            instructions: instructions
        };
    }

    /**
     * Validate PPTX template
     * @param {string} templatePath - Path to template
     */
    async validateTemplate(templatePath) {
        try {
            const exists = await fs.pathExists(templatePath);
            if (!exists) {
                return {
                    valid: false,
                    message: 'Template file does not exist'
                };
            }

            const ext = path.extname(templatePath).toLowerCase();
            if (ext !== '.pptx') {
                return {
                    valid: false,
                    message: 'Template must be a .pptx file'
                };
            }

            return {
                valid: true,
                message: 'Template appears valid'
            };

        } catch (error) {
            return {
                valid: false,
                message: `Error validating template: ${error.message}`
            };
        }
    }

    /**
     * Get supported placeholders
     */
    getSupportedPlaceholders() {
        return [
            { placeholder: '{{name}}', description: 'Person\'s name (original case)' },
            { placeholder: '{{NAME}}', description: 'Person\'s name (uppercase)' },
            { placeholder: '{{Name}}', description: 'Person\'s name (title case)' },
            { placeholder: '{{date}}', description: 'Current date' },
            { placeholder: '{{course}}', description: 'Course name' },
            { placeholder: '{{instructor}}', description: 'Instructor name' },
            { placeholder: '{{organization}}', description: 'Organization name' }
        ];
    }
}

module.exports = PPTXProcessor;
