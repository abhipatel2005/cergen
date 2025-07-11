const AdmZip = require('adm-zip');

async function testPlaceholderExtraction() {
    try {
        console.log('Testing placeholder extraction...');
        const templatePath = './test-template.pptx';
        
        const placeholders = new Set();
        
        console.log('Processing PPTX file:', templatePath);
        const zip = new AdmZip(templatePath);
        const zipEntries = zip.getEntries();
        console.log('Found', zipEntries.length, 'entries in PPTX');
        
        // Look for slide content in PPTX
        zipEntries.forEach(entry => {
            console.log('Entry:', entry.entryName);
            if (entry.entryName.includes('slides/slide') && entry.entryName.endsWith('.xml')) {
                console.log('Processing slide:', entry.entryName);
                const content = entry.getData().toString('utf8');
                
                // Show first 500 characters of content
                console.log('Content preview:', content.substring(0, 500));
                
                // Find all placeholders in format {{placeholder}}
                const matches = content.match(/\{\{([^}]+)\}\}/g);
                console.log('Found matches in', entry.entryName, ':', matches);
                if (matches) {
                    matches.forEach(match => {
                        const placeholder = match.replace(/[{}]/g, '').trim();
                        if (placeholder) {
                            console.log('Adding placeholder:', placeholder);
                            placeholders.add(placeholder);
                        }
                    });
                }
            }
        });
        
        const result = Array.from(placeholders);
        console.log('Final extracted placeholders:', result);
        return result;
    } catch (error) {
        console.error('Error:', error);
        return [];
    }
}

testPlaceholderExtraction();
