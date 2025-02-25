// script.js - Fixed for DocxTemplater
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
        
        try {
            const arrayBuffer = await loadFile(templateUrl);
            
            if (!arrayBuffer) {
                throw new Error('Failed to load template document');
            }
            
            // Prepare data for template replacement
            const patchData = preparePatchData(rowData, authorName, imageData);
            
            // Create a zip of the docx template
            const zip = new PizZip(arrayBuffer);
            
            // Create a new DocxTemplater instance
            const doc = new window.docxtemplater();
            
            // Configure document
            doc.loadZip(zip);
            
            // Set the data to be injected
            doc.setData(patchData);
            
            // Perform the template substitution
            doc.render();
            
            // Generate output
            const output = doc.getZip().generate({
                type: 'blob',
                mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            });
            
            // Save the document
            saveAs(output, `High_Risk_Report_${formatDate(new Date())}.docx`);
            
        } catch (error) {
            if (error.properties && error.properties.errors) {
                const errorMessages = error.properties.errors.map(error => {
                    return `Error in template: ${error}`;
                }).join("\n");
                
                console.error("Template errors:", errorMessages);
                alert("Template errors: " + errorMessages);
            } else {
                console.error("Error:", error);
                alert("Error: " + error.message);
            }
        }
        
    } catch (error) {
        console.error('Error generating document:', error);
        alert('Error generating document: ' + error.message);
    }
}

// Function to load the template file
function loadFile(url) {
    return new Promise((resolve, reject) => {
        // Use raw XMLHttpRequest for better error handling
        const xhr = new XMLHttpRequest();
        xhr.open('GET', url, true);
        xhr.responseType = 'arraybuffer';
        
        xhr.onload = function() {
            if (xhr.status === 200) {
                resolve(xhr.response);
            } else {
                reject(new Error(`Failed to load ${url}: ${xhr.status} ${xhr.statusText}`));
            }
        };
        
        xhr.onerror = function() {
            reject(new Error(`Network error while loading ${url}`));
        };
        
        xhr.send();
    });
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
        offender_store: '',
        images: 'None available'  // Default value for images
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
    
    // Handle image if available - but note that images require a special module in DocxTemplater
    // This provides a basic text fallback
    if (imageData) {
        // For now, we're just indicating an image was provided
        // To actually embed images, you would need the docxtemplater-image-module
        patchData.images = "Image provided (requires image module)";
    }
    
    // Ensure all values are strings to prevent template errors
    Object.keys(patchData).forEach(key => {
        if (patchData[key] === null || patchData[key] === undefined) {
            patchData[key] = '';
        } else {
            patchData[key] = String(patchData[key]);
        }
    });
    
    return patchData;
}

function formatDate(date) {
    const day = date.getDate().toString().padStart(2, '0');
    const month = (date.getMonth() + 1).toString().padStart(2, '0');
    const year = date.getFullYear();
    return `${day}-${month}-${year}`;
}
