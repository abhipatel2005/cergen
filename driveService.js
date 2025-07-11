const { google } = require('googleapis');
const fs = require('fs-extra');
const path = require('path');

class DriveService {
    constructor() {
        this.drive = null;
        this.isConfigured = false;
        this.folderId = null;
        this.serviceAccount = null;
    }

    /**
     * Configure Google Drive service with service account credentials
     * @param {Object} config - Drive configuration
     */
    configure(config) {
        try {
            let credentials;

            // Handle different input formats
            if (config.serviceAccountJson) {
                // Full JSON object provided
                credentials = config.serviceAccountJson;
            } else if (config.privateKey) {
                // Legacy format - try to parse private key
                let privateKey = config.privateKey;
                if (typeof privateKey === 'string') {
                    try {
                        privateKey = JSON.parse(privateKey);
                        credentials = privateKey;
                    } catch (e) {
                        // If it's not JSON, create credentials object
                        credentials = {
                            type: 'service_account',
                            project_id: 'certificate-generator',
                            private_key: privateKey,
                            client_email: config.serviceAccountEmail,
                            auth_uri: 'https://accounts.google.com/o/oauth2/auth',
                            token_uri: 'https://oauth2.googleapis.com/token',
                            auth_provider_x509_cert_url: 'https://www.googleapis.com/oauth2/v1/certs'
                        };
                    }
                } else {
                    credentials = privateKey;
                }
            } else {
                throw new Error('No service account credentials provided');
            }

            // Validate required fields
            if (!credentials.client_email || !credentials.private_key) {
                throw new Error('Invalid service account JSON. Missing client_email or private_key.');
            }

            // Create JWT auth
            const auth = new google.auth.GoogleAuth({
                credentials: credentials,
                scopes: ['https://www.googleapis.com/auth/drive.readonly']
            });

            // Initialize Drive API
            this.drive = google.drive({ version: 'v3', auth });
            this.folderId = config.folderId;
            this.serviceAccount = config.serviceAccountEmail;
            this.isConfigured = true;

            return { success: true, message: 'Google Drive configured successfully' };
        } catch (error) {
            console.error('Drive configuration error:', error);
            return { success: false, error: error.message };
        }
    }

    /**
     * Test the Drive connection
     */
    async testConnection() {
        if (!this.isConfigured) {
            return { success: false, error: 'Google Drive not configured' };
        }

        try {
            // Test by listing files in the folder
            const response = await this.drive.files.list({
                q: `'${this.folderId}' in parents and trashed=false`,
                fields: 'files(id, name, mimeType, size)',
                pageSize: 10
            });

            return {
                success: true,
                message: 'Connection successful',
                templateCount: response.data.files.length
            };
        } catch (error) {
            console.error('Drive connection test failed:', error);
            return { success: false, error: error.message };
        }
    }

    /**
     * Get templates from Google Drive folder
     * @param {string} category - Template category filter
     */
    async getTemplates(category = '') {
        if (!this.isConfigured) {
            return {
                success: false,
                error: 'Google Drive not configured. Please configure your service account credentials first.',
                templates: []
            };
        }

        try {
            // Build query to get PPTX and PDF files
            let query = `'${this.folderId}' in parents and trashed=false and (mimeType='application/vnd.openxmlformats-officedocument.presentationml.presentation' or mimeType='application/pdf')`;

            // Add category filter if specified
            if (category) {
                query += ` and name contains '${category}'`;
            }

            const response = await this.drive.files.list({
                q: query,
                fields: 'files(id, name, mimeType, size, createdTime, modifiedTime)',
                orderBy: 'modifiedTime desc',
                pageSize: 50
            });

            const templates = response.data.files.map(file => ({
                id: file.id,
                name: file.name,
                mimeType: file.mimeType,
                size: parseInt(file.size) || 0,
                createdTime: file.createdTime,
                modifiedTime: file.modifiedTime,
                category: this.extractCategory(file.name)
            }));

            return {
                success: true,
                templates: templates,
                count: templates.length
            };
        } catch (error) {
            console.error('Error fetching Drive templates:', error);
            return { success: false, error: error.message };
        }
    }

    /**
     * Download a template file from Google Drive
     * @param {string} fileId - Google Drive file ID
     * @param {string} fileName - Original file name
     */
    async downloadTemplate(fileId, fileName) {
        if (!this.isConfigured) {
            throw new Error('Google Drive not configured');
        }

        try {
            const response = await this.drive.files.get({
                fileId: fileId,
                alt: 'media'
            }, { responseType: 'stream' });

            // Create temp directory if it doesn't exist
            const tempDir = path.join(__dirname, 'temp');
            await fs.ensureDir(tempDir);

            // Generate unique filename
            const timestamp = Date.now();
            const extension = path.extname(fileName);
            const tempFileName = `drive_template_${timestamp}${extension}`;
            const tempFilePath = path.join(tempDir, tempFileName);

            // Save file to temp directory
            const writer = fs.createWriteStream(tempFilePath);
            response.data.pipe(writer);

            return new Promise((resolve, reject) => {
                writer.on('finish', () => {
                    resolve({
                        success: true,
                        filePath: tempFilePath,
                        fileName: tempFileName,
                        originalName: fileName
                    });
                });
                writer.on('error', reject);
            });
        } catch (error) {
            console.error('Error downloading Drive template:', error);
            throw new Error(`Failed to download template: ${error.message}`);
        }
    }

    /**
     * Extract category from filename
     * @param {string} fileName - File name
     */
    extractCategory(fileName) {
        const name = fileName.toLowerCase();
        if (name.includes('professional')) return 'professional';
        if (name.includes('academic')) return 'academic';
        if (name.includes('corporate')) return 'corporate';
        if (name.includes('creative')) return 'creative';
        return 'general';
    }

    /**
     * Get Drive folder URL
     */
    getFolderUrl() {
        if (!this.folderId) return null;
        return `https://drive.google.com/drive/folders/${this.folderId}`;
    }

    /**
     * Check if service is configured
     */
    isServiceConfigured() {
        return this.isConfigured;
    }

    /**
     * Get configuration status
     */
    getStatus() {
        return {
            configured: this.isConfigured,
            folderId: this.folderId,
            serviceAccount: this.serviceAccount,
            folderUrl: this.getFolderUrl()
        };
    }
}

module.exports = new DriveService();
