const XLSX = require('xlsx');

// Create test data
const testData = [
    { name: 'John Smith', email: 'john.smith@example.com' },
    { name: 'Jane Doe', email: 'jane.doe@example.com' },
    { name: 'Mike Johnson', email: 'mike.johnson@example.com' },
    { name: 'Sarah Wilson', email: 'sarah.wilson@example.com' },
    { name: 'David Brown', email: 'david.brown@example.com' }
];

// Create workbook and worksheet
const wb = XLSX.utils.book_new();
const ws = XLSX.utils.json_to_sheet(testData);

// Add worksheet to workbook
XLSX.utils.book_append_sheet(wb, ws, 'Recipients');

// Write to file
XLSX.writeFile(wb, 'test-recipients.xlsx');

console.log('Test Excel file created: test-recipients.xlsx');
