const fs = require('fs-extra');
const path = require('path');

class CanvaIntegration {
    constructor() {
        this.apiKey = process.env.CANVA_API_KEY || '';
        this.baseUrl = 'https://api.canva.com/v1';
        
        // Mock templates for demonstration (in real implementation, these would come from Canva API)
        this.mockTemplates = [
            {
                id: 'cert-professional-001',
                title: 'Professional Certificate',
                description: 'Clean and professional design with elegant borders',
                category: 'Professional',
                thumbnail: this.generateThumbnailSVG('Professional Certificate', '#667eea'),
                downloadUrl: 'https://example.com/templates/professional-cert.pptx',
                placeholders: ['{{name}}', '{{course}}', '{{date}}', '{{organization}}']
            },
            {
                id: 'cert-modern-002',
                title: 'Modern Achievement',
                description: 'Contemporary design with gradient backgrounds',
                category: 'Modern',
                thumbnail: this.generateThumbnailSVG('Modern Achievement', '#764ba2'),
                downloadUrl: 'https://example.com/templates/modern-achievement.pptx',
                placeholders: ['{{name}}', '{{course}}', '{{instructor}}', '{{date}}']
            },
            {
                id: 'cert-classic-003',
                title: 'Classic Diploma',
                description: 'Traditional diploma style with formal layout',
                category: 'Classic',
                thumbnail: this.generateThumbnailSVG('Classic Diploma', '#28a745'),
                downloadUrl: 'https://example.com/templates/classic-diploma.pptx',
                placeholders: ['{{name}}', '{{course}}', '{{organization}}', '{{date}}']
            },
            {
                id: 'cert-creative-004',
                title: 'Creative Certificate',
                description: 'Artistic design with creative elements',
                category: 'Creative',
                thumbnail: this.generateThumbnailSVG('Creative Certificate', '#ff6b6b'),
                downloadUrl: 'https://example.com/templates/creative-cert.pptx',
                placeholders: ['{{name}}', '{{course}}', '{{instructor}}', '{{organization}}']
            },
            {
                id: 'cert-corporate-005',
                title: 'Corporate Training',
                description: 'Professional corporate training certificate',
                category: 'Corporate',
                thumbnail: this.generateThumbnailSVG('Corporate Training', '#4ecdc4'),
                downloadUrl: 'https://example.com/templates/corporate-training.pptx',
                placeholders: ['{{name}}', '{{course}}', '{{instructor}}', '{{date}}', '{{organization}}']
            },
            {
                id: 'cert-academic-006',
                title: 'Academic Excellence',
                description: 'Academic achievement certificate with scholarly design',
                category: 'Academic',
                thumbnail: this.generateThumbnailSVG('Academic Excellence', '#9b59b6'),
                downloadUrl: 'https://example.com/templates/academic-excellence.pptx',
                placeholders: ['{{name}}', '{{course}}', '{{instructor}}', '{{date}}', '{{organization}}']
            }
        ];
    }

    /**
     * Generate SVG thumbnail for mock templates
     * @param {string} title - Template title
     * @param {string} color - Primary color
     * @returns {string} Data URL for SVG
     */
    generateThumbnailSVG(title, color) {
        const svg = `
            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 300 200">
                <defs>
                    <linearGradient id="grad" x1="0%" y1="0%" x2="100%" y2="100%">
                        <stop offset="0%" style="stop-color:${color};stop-opacity:1" />
                        <stop offset="100%" style="stop-color:${color}88;stop-opacity:1" />
                    </linearGradient>
                </defs>
                <rect width="300" height="200" fill="url(#grad)" rx="10"/>
                <rect x="20" y="20" width="260" height="160" fill="none" stroke="white" stroke-width="2" rx="5"/>
                <text x="150" y="80" text-anchor="middle" fill="white" font-size="16" font-weight="bold">CERTIFICATE</text>
                <text x="150" y="100" text-anchor="middle" fill="white" font-size="12">OF COMPLETION</text>
                <text x="150" y="130" text-anchor="middle" fill="white" font-size="10">${title}</text>
                <line x1="50" y1="150" x2="250" y2="150" stroke="white" stroke-width="1"/>
                <text x="150" y="165" text-anchor="middle" fill="white" font-size="8">{{name}}</text>
            </svg>
        `;
        return `data:image/svg+xml,${encodeURIComponent(svg)}`;
    }

    /**
     * Get available certificate templates
     * @param {Object} filters - Filter options
     * @returns {Array} Array of template objects
     */
    async getTemplates(filters = {}) {
        try {
            // In a real implementation, this would make an API call to Canva
            // For now, we'll return mock templates
            
            let templates = [...this.mockTemplates];
            
            // Apply filters
            if (filters.category) {
                templates = templates.filter(template => 
                    template.category.toLowerCase() === filters.category.toLowerCase()
                );
            }
            
            if (filters.search) {
                const searchTerm = filters.search.toLowerCase();
                templates = templates.filter(template =>
                    template.title.toLowerCase().includes(searchTerm) ||
                    template.description.toLowerCase().includes(searchTerm)
                );
            }
            
            return {
                success: true,
                templates: templates,
                total: templates.length
            };
        } catch (error) {
            return {
                success: false,
                error: error.message,
                templates: []
            };
        }
    }

    /**
     * Get template categories
     * @returns {Array} Array of category names
     */
    getCategories() {
        const categories = [...new Set(this.mockTemplates.map(template => template.category))];
        return categories.sort();
    }

    /**
     * Download template by ID
     * @param {string} templateId - Template ID
     * @param {string} outputPath - Where to save the template
     * @returns {Object} Download result
     */
    async downloadTemplate(templateId, outputPath) {
        try {
            const template = this.mockTemplates.find(t => t.id === templateId);
            
            if (!template) {
                throw new Error('Template not found');
            }

            // In a real implementation, this would download from Canva
            // For now, we'll create a mock PPTX file or copy an existing template
            
            // Check if we have a sample template to copy
            const sampleTemplatePath = path.join(__dirname, 'certificate-template.pptx');
            
            if (await fs.pathExists(sampleTemplatePath)) {
                await fs.copy(sampleTemplatePath, outputPath);
            } else {
                // Create a placeholder file
                await fs.writeFile(outputPath, 'Mock PPTX template content');
            }

            return {
                success: true,
                templateId: templateId,
                outputPath: outputPath,
                template: template
            };
        } catch (error) {
            return {
                success: false,
                error: error.message
            };
        }
    }

    /**
     * Get template details by ID
     * @param {string} templateId - Template ID
     * @returns {Object} Template details
     */
    getTemplateDetails(templateId) {
        const template = this.mockTemplates.find(t => t.id === templateId);
        
        if (!template) {
            return {
                success: false,
                error: 'Template not found'
            };
        }

        return {
            success: true,
            template: template
        };
    }

    /**
     * Check if Canva API is configured
     * @returns {Object} Configuration status
     */
    checkConfiguration() {
        return {
            configured: !!this.apiKey,
            hasApiKey: !!this.apiKey,
            mockMode: !this.apiKey, // Using mock templates when no API key
            message: this.apiKey ? 'Canva API configured' : 'Using mock templates (no API key configured)'
        };
    }

    /**
     * Search templates with advanced options
     * @param {Object} searchOptions - Search parameters
     * @returns {Object} Search results
     */
    async searchTemplates(searchOptions = {}) {
        const {
            query = '',
            category = '',
            limit = 20,
            offset = 0
        } = searchOptions;

        try {
            let templates = [...this.mockTemplates];

            // Apply search query
            if (query) {
                const searchTerm = query.toLowerCase();
                templates = templates.filter(template =>
                    template.title.toLowerCase().includes(searchTerm) ||
                    template.description.toLowerCase().includes(searchTerm) ||
                    template.category.toLowerCase().includes(searchTerm)
                );
            }

            // Apply category filter
            if (category) {
                templates = templates.filter(template =>
                    template.category.toLowerCase() === category.toLowerCase()
                );
            }

            // Apply pagination
            const total = templates.length;
            templates = templates.slice(offset, offset + limit);

            return {
                success: true,
                templates: templates,
                total: total,
                limit: limit,
                offset: offset,
                hasMore: offset + limit < total
            };
        } catch (error) {
            return {
                success: false,
                error: error.message,
                templates: []
            };
        }
    }
}

module.exports = CanvaIntegration;
