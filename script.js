// script.js
async function generateWord() {
    try {
        // Get input values
        const authorName = document.getElementById('author').value.trim();
        const excelDataText = document.getElementById('excelData').value.trim();
        const imageElement = document.getElementById('previewImage');
        const imageData = imageElement.hidden ? null : imageElement.src;
        
        // Validate inputs
        if (!excelDataText) {
            alert('Please paste Excel data first.');
            return;
        }
        
        // Parse Excel data (assuming tab or comma separated values)
        const rowData = parseExcelData(excelDataText);
        
        // Load the template document
        const templateUrl = 'High Risk Manoeuvre Template.docx';
        const templateArrayBuffer = await fetch(templateUrl).then(res => {
            if (!res.ok) throw new Error('Failed to load template document. Make sure the template file is in the correct location.');
            return res.arrayBuffer();
        });
        
        // Prepare data for template replacement
        const patchData = preparePatchData(rowData, authorName, imageData);
        
        // Create a new document from the template
        const doc = await docx.TemplateHandler.process(templateArrayBuffer, patchData);
        
        // Generate the document
        const blob = await docx.Packer.toBlob(doc);
        
        // Save the document using FileSaver.js
        saveAs(blob, `High_Risk_Report_${formatDate(new Date())}.docx`);
        
    } catch (error) {
        console.error('Error generating document:', error);
        alert('Error generating document: ' + error.message);
    }
}

function parseExcelData(excelText) {
    // Try to determine the delimiter (tab or comma)
    const delimiter = excelText.includes('\t') ? '\t' : ',';
    
    // Split the data into rows
    const rows = excelText.split('\n').filter(row => row.trim());
    
    // Assuming the first row contains headers
    const headers = rows[0].split(delimiter).map(header => header.trim());
    
    // Get the data row (assuming single row paste)
    const dataValues = rows.length > 1 ? 
        rows[1].split(delimiter).map(value => value.trim()) : 
        rows[0].split(delimiter).map(value => value.trim());
    
    // Create an object mapping headers to values
    const data = {};
    headers.forEach((header, index) => {
        if (index < dataValues.length) {
            data[header] = dataValues[index];
        }
    });
    
    return data;
}

function preparePatchData(rowData, author, imageData) {
    // Define template fields based on the document structure
    const templateFields = [
        'description', 'entry_time', 'entry_gate', 'exit_time', 'exit_gate',
        'fee', 'paid_amt', 'serial', 'offender_name', 'offender_contact',
        'offender_vehicle', 'offender_store'
    ];
    
    // Initialize patch data with default values
    const patchData = {
        description: '',
        entry_time: '',
        entry_gate: '',
        exit_time: '',
        exit_gate: '',
        fee: '',
        paid_amt: '',
        serial: 'No',
        author: author || 'Not Specified',
        offender_name: '',
        offender_contact: '',
        offender_vehicle: '',
        offender_store: ''
    };
    
    // Map Excel data to template fields
    // First try direct mapping (exact field names)
    Object.keys(rowData).forEach(key => {
        // Check if the key matches any template field directly
        if (templateFields.includes(key.toLowerCase())) {
            patchData[key.toLowerCase()] = rowData[key];
        }
        // Try to match common variations of field names
        else {
            // Try to map based on partial matches
            if (key.toLowerCase().includes('desc')) patchData.description = rowData[key];
            else if (key.toLowerCase().includes('entry') && key.toLowerCase().includes('time')) patchData.entry_time = rowData[key];
            else if (key.toLowerCase().includes('entry') && key.toLowerCase().includes('gate')) patchData.entry_gate = rowData[key];
            else if (key.toLowerCase().includes('exit') && key.toLowerCase().includes('time')) patchData.exit_time = rowData[key];
            else if (key.toLowerCase().includes('exit') && key.toLowerCase().includes('gate')) patchData.exit_gate = rowData[key];
            else if (key.toLowerCase().includes('fee') && !key.toLowerCase().includes('paid')) patchData.fee = rowData[key];
            else if (key.toLowerCase().includes('paid')) patchData.paid_amt = rowData[key];
            else if (key.toLowerCase().includes('serial') || key.toLowerCase().includes('offender')) patchData.serial = rowData[key];
            else if (key.toLowerCase().includes('name')) patchData.offender_name = rowData[key];
            else if (key.toLowerCase().includes('contact')) patchData.offender_contact = rowData[key];
            else if (key.toLowerCase().includes('vehicle')) patchData.offender_vehicle = rowData[key];
            else if (key.toLowerCase().includes('store')) patchData.offender_store = rowData[key];
        }
    });
    
    // Handle image if available
    if (imageData) {
        patchData.images = {
            type: 'image',
            data: imageData,
            width: 400,
            height: 300
        };
    } else {
        patchData.images = 'None available';
    }
    
    return patchData;
}

function formatDate(date) {
    const day = date.getDate().toString().padStart(2, '0');
    const month = (date.getMonth() + 1).toString().padStart(2, '0');
    const year = date.getFullYear();
    return `${day}-${month}-${year}`;
}

// Add event listeners for the dropzone functionality
document.addEventListener('DOMContentLoaded', function() {
    // All your existing drag and drop code is already in the HTML
    // This ensures the function is available when called from the button
});
