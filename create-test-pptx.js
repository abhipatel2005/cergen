const PptxGenJS = require('pptxgenjs');

// Create a new presentation
const pptx = new PptxGenJS();

// Add a slide
const slide = pptx.addSlide();

// Add title with placeholder
slide.addText('Certificate of Completion', {
    x: 1,
    y: 1,
    w: 8,
    h: 1,
    fontSize: 24,
    bold: true,
    align: 'center'
});

// Add content with placeholders
slide.addText('This is to certify that {{name}} has successfully completed the {{course}} course.', {
    x: 1,
    y: 3,
    w: 8,
    h: 1,
    fontSize: 16,
    align: 'center'
});

slide.addText('Completed on: {{date}}', {
    x: 1,
    y: 4.5,
    w: 8,
    h: 0.5,
    fontSize: 14,
    align: 'center'
});

slide.addText('Instructor: {{instructor}}', {
    x: 1,
    y: 5.5,
    w: 8,
    h: 0.5,
    fontSize: 14,
    align: 'center'
});

slide.addText('Organization: {{organization}}', {
    x: 1,
    y: 6.5,
    w: 8,
    h: 0.5,
    fontSize: 14,
    align: 'center'
});

// Save the presentation
pptx.writeFile({ fileName: 'test-template.pptx' })
    .then(() => {
        console.log('Test PPTX template created: test-template.pptx');
        console.log('Placeholders included: {{name}}, {{course}}, {{date}}, {{instructor}}, {{organization}}');
    })
    .catch(err => {
        console.error('Error creating PPTX:', err);
    });
