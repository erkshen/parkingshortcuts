// script.js - With Image Module Support
async function generateWord() {
    try {
        // Get input values
        const authorName = document.getElementById('author').value.trim();
        const excelDataText = document.getElementById('excelData').value.trim();
        //const imageElement = document.getElementById('previewImage');
        //const imageData = imageElement.hidden ? null : imageElement.src;
        
        // Validate inputs
        if (!excelDataText) {
            alert('Please paste Excel data first.');
            return;
        }
        
        // Parse Excel data
        const rowData = parseExcelData(excelDataText);
        
        // Load the template document
        const templateUrl = 'High Risk Manoeuvre Template.docx';
        
        const docxType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
        PizZipUtils.getBinaryContent(
            templateUrl,
            function (error, content) {
                if (error) {
                    console.error(error);
                    return;
                }

                const zip = new PizZip(content);
                const doc = new docxtemplater(zip);

                doc.render({
                    description: rowData[2],
                    entry_time: rowData[3],
                    entry_gate: rowData[4],
                    exit_time: rowData[5],
                    exit_gate: rowData[6],
                    fee: rowData[8],
                    paid_amt: '$0',
                    serial: rowData[9] || 'No',
                    author: author || 'Not Specified',
                    offender_name: rowData[10],
                    offender_contact: rowData[11],
                    offender_vehicle: rowData[12],
                    offender_store: rowData[13],
                });
                
                const out = doc.getZip().generate({
                    type: "blob",
                    mimeType: docxType,
                });
                saveAs(out, `${rowData[1]}.docx`);
                
            }
        );
        
    } catch (error) {
        console.error('Error generating document:', error);
        alert('Error generating document: ' + error.message);
    }
}

// Function to convert data URL to array buffer
function dataURLtoArrayBuffer(dataURL) {
    // Skip if no dataURL
    if (!dataURL || !dataURL.startsWith('data:')) {
        return null;
    }
    
    // Remove the data URL prefix
    const base64 = dataURL.split(',')[1];
    const binary = atob(base64);
    const len = binary.length;
    const buffer = new ArrayBuffer(len);
    const view = new Uint8Array(buffer);
    
    for (let i = 0; i < len; i++) {
        view[i] = binary.charCodeAt(i);
    }
    
    return buffer;
}

// Your existing functions (loadFile, parseExcelData, formatDate)
function loadFile(url) {
    return new Promise((resolve, reject) => {
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
    const rows = excelText.split('\t').map(function(item) {
        return item.trim();
    });
    console.log(rows);
    return rows;
}

function preparePatchData(rowData, author, imageData) {
    // Initialize patch data with default values
    const patchData = {
        description: rowData[2] || '',
        entry_time: rowData[3] || '',
        entry_gate: rowData[4] || '',
        exit_time: rowData[5] || '',
        exit_gate: rowData[6] || '',
        fee: rowData[8] || '',
        paid_amt: '$0',
        serial: rowData[9] || 'No',
        author: author || 'Not Specified',
        offender_name: rowData[10] || '',
        offender_contact: rowData[11] || '',
        offender_vehicle: rowData[12] || '',
        offender_store: rowData[13] || '',
    };
    
    
    // Ensure all values are strings to prevent template errors
    Object.keys(patchData).forEach(key => {
        if (patchData[key] === null || patchData[key] === undefined) {
            patchData[key] = '';
        } else if (key !== 'images') { // Don't convert image data to string
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
