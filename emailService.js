const nodemailer = require('nodemailer');
const fs = require('fs-extra');
const path = require('path');

class EmailService {
    constructor() {
        this.transporter = null;
        this.isConfigured = false;
        this.defaultConfig = {
            service: 'gmail', // Default to Gmail
            auth: {
                user: process.env.EMAIL_USER || '',
                pass: process.env.EMAIL_PASS || ''
            }
        };
    }

    /**
     * Configure email service with custom settings
     * @param {Object} config - Email configuration
     */
    configure(config) {
        try {
            this.transporter = nodemailer.createTransport({
                service: config.service || 'gmail',
                auth: {
                    user: config.user,
                    pass: config.password
                },
                ...config.options
            });
            this.isConfigured = true;
            return { success: true, message: 'Email service configured successfully' };
        } catch (error) {
            return { success: false, error: error.message };
        }
    }

    /**
     * Configure with environment variables
     */
    configureFromEnv() {
        if (process.env.EMAIL_USER && process.env.EMAIL_PASS) {
            return this.configure({
                service: process.env.EMAIL_SERVICE || 'gmail',
                user: process.env.EMAIL_USER,
                password: process.env.EMAIL_PASS
            });
        }
        return { success: false, error: 'Email credentials not found in environment variables' };
    }

    /**
     * Test email configuration
     */
    async testConnection() {
        if (!this.isConfigured) {
            return { success: false, error: 'Email service not configured' };
        }

        try {
            await this.transporter.verify();
            return { success: true, message: 'Email connection verified successfully' };
        } catch (error) {
            // Provide more helpful error messages
            let errorMessage = error.message;
            if (error.code === 'EAUTH') {
                errorMessage = 'Authentication failed. Please check your email and password. For Gmail, use an App Password instead of your regular password.';
            } else if (error.code === 'ECONNECTION') {
                errorMessage = 'Connection failed. Please check your internet connection and email service settings.';
            }
            return { success: false, error: errorMessage };
        }
    }

    /**
     * Send certificate via email
     * @param {Object} options - Email options
     */
    async sendCertificate(options) {
        if (!this.isConfigured) {
            throw new Error('Email service not configured');
        }

        const {
            to,
            name,
            certificatePath,
            subject,
            customText,
            course,
            organization,
            senderName
        } = options;

        // Generate email content
        const emailContent = this.generateEmailContent({
            name,
            customText,
            course,
            organization,
            senderName
        });

        const mailOptions = {
            from: `${senderName || 'Certificate System'} <${process.env.EMAIL_USER}>`,
            to: to,
            subject: subject || `Your Certificate - ${course || 'Course Completion'}`,
            html: emailContent,
            attachments: [
                {
                    filename: `certificate-${name.replace(/[^a-zA-Z0-9]/g, '_')}.pdf`,
                    path: certificatePath,
                    contentType: 'application/pdf'
                }
            ]
        };

        try {
            const result = await this.transporter.sendMail(mailOptions);
            return {
                success: true,
                messageId: result.messageId,
                recipient: to,
                name: name
            };
        } catch (error) {
            throw new Error(`Failed to send email to ${to}: ${error.message}`);
        }
    }

    /**
     * Generate HTML email content
     * @param {Object} data - Email data
     */
    generateEmailContent(data) {
        const { name, customText, course, organization, senderName } = data;

        const defaultText = `
            <p>Dear ${name},</p>
            <p>Congratulations! We are pleased to inform you that you have successfully completed the requirements for <strong>${course || 'the course'}</strong>.</p>
            <p>Please find your certificate attached to this email.</p>
            <p>We appreciate your dedication and hard work throughout the program.</p>
        `;

        const emailText = customText || defaultText;

        return `
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="utf-8">
                <style>
                    body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; }
                    .container { max-width: 600px; margin: 0 auto; padding: 20px; }
                    .header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 20px; text-align: center; border-radius: 8px 8px 0 0; }
                    .content { background: #f9f9f9; padding: 30px; border-radius: 0 0 8px 8px; }
                    .footer { text-align: center; margin-top: 20px; color: #666; font-size: 12px; }
                    .certificate-info { background: white; padding: 15px; border-radius: 5px; margin: 15px 0; border-left: 4px solid #667eea; }
                </style>
            </head>
            <body>
                <div class="container">
                    <div class="header">
                        <h1>🎓 Certificate Delivery</h1>
                        <p>${organization || 'Certificate Authority'}</p>
                    </div>
                    <div class="content">
                        ${emailText}
                        
                        <div class="certificate-info">
                            <h3>📋 Certificate Details:</h3>
                            <p><strong>Recipient:</strong> ${name}</p>
                            ${course ? `<p><strong>Course:</strong> ${course}</p>` : ''}
                            <p><strong>Date Issued:</strong> ${new Date().toLocaleDateString()}</p>
                            ${organization ? `<p><strong>Issued by:</strong> ${organization}</p>` : ''}
                        </div>
                        
                        <p>Best regards,<br>
                        ${senderName || 'The Certificate Team'}</p>
                    </div>
                    <div class="footer">
                        <p>This is an automated message. Please do not reply to this email.</p>
                        <p>Generated by Certificate Generator System</p>
                    </div>
                </div>
            </body>
            </html>
        `;
    }

    /**
     * Send certificates to multiple recipients
     * @param {Array} recipients - Array of recipient objects
     * @param {Object} emailConfig - Email configuration
     */
    async sendBulkCertificates(recipients, emailConfig) {
        const results = [];
        const errors = [];

        for (const recipient of recipients) {
            try {
                const result = await this.sendCertificate({
                    to: recipient.email,
                    name: recipient.name,
                    certificatePath: recipient.certificatePath,
                    subject: emailConfig.subject,
                    customText: emailConfig.customText,
                    course: emailConfig.course,
                    organization: emailConfig.organization,
                    senderName: emailConfig.senderName
                });
                results.push(result);

                // Add delay between emails to avoid rate limiting
                await this.delay(1000);

            } catch (error) {
                errors.push({
                    name: recipient.name,
                    email: recipient.email,
                    error: error.message
                });
            }
        }

        return {
            success: errors.length === 0,
            sent: results.length,
            failed: errors.length,
            results: results,
            errors: errors
        };
    }

    /**
     * Utility function to add delay
     * @param {number} ms - Milliseconds to delay
     */
    delay(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    /**
     * Get email service status
     */
    getStatus() {
        return {
            configured: this.isConfigured,
            service: this.transporter ? 'Configured' : 'Not configured',
            user: process.env.EMAIL_USER ? 'Set' : 'Not set'
        };
    }

    /**
     * Generate email templates for different scenarios
     */
    getEmailTemplates() {
        return {
            completion: {
                subject: "🎓 Congratulations! Your {{course}} Certificate is Ready",
                text: `
                    <p>Dear {{name}},</p>
                    <p>Congratulations! We are thrilled to inform you that you have successfully completed <strong>{{course}}</strong>.</p>
                    <p>Your dedication and hard work throughout the program have been truly impressive. Please find your official certificate attached to this email.</p>
                    <p>We wish you continued success in your future endeavors!</p>
                `
            },
            achievement: {
                subject: "🏆 Achievement Unlocked - {{course}} Certificate",
                text: `
                    <p>Hello {{name}},</p>
                    <p>What an achievement! You have successfully completed <strong>{{course}}</strong> and earned your certificate.</p>
                    <p>This accomplishment represents your commitment to learning and professional growth. Your certificate is attached and ready for you to share with pride.</p>
                    <p>Keep up the excellent work!</p>
                `
            },
            professional: {
                subject: "Certificate of Completion - {{course}}",
                text: `
                    <p>Dear {{name}},</p>
                    <p>This email confirms that you have successfully completed the requirements for <strong>{{course}}</strong>.</p>
                    <p>Your certificate has been generated and is attached to this email for your records.</p>
                    <p>Thank you for your participation in our program.</p>
                `
            }
        };
    }
}

module.exports = EmailService;
