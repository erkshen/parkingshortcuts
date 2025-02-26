// script.js - With Image Module Support
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
        
        // Parse Excel data
        const rowData = parseExcelData(excelDataText);
        
        // Load the template document
        const templateUrl = 'High Risk Manoeuvre Template.docx';
        
        try {
            const arrayBuffer = await loadFile(templateUrl);
            
            if (!arrayBuffer) {
                throw new Error('Failed to load template document');
            }

            // Configure the image module (if image is available)
            let imageModule = null;
            if (imageData) {
                imageModule = new ImageModule({
                    centered: false,
                    fileType: "docx",
                    getImage(tagValue) {
                        // In this case tagValue will be a URL tagValue = "https://docxtemplater.com/puffin.png"
                        return new Promise(function (resolve, reject) {
                            PizZipUtils.getBinaryContent(
                                tagValue,
                                function (error, content) {
                                    if (error) {
                                        return reject(error);
                                    }
                                    return resolve(content);
                                }
                            );
                        });
                    },
                    getSize(img, tagValue, tagName) {
                        return new Promise(function (resolve, reject) {
                            const image = new Image();
                            image.src = tagValue;
                            image.onload = function () {
                                resolve([image.width, image.height]);
                            };
                            image.onerror = function (e) {
                                console.log(
                                    "img, tagValue, tagName : ",
                                    img,
                                    tagValue,
                                    tagName
                                );
                                alert(
                                    "An error occured while loading " +
                                        tagValue
                                );
                                reject(e);
                            };
                        });
                    },
                });
            }
            
            // Prepare data for template replacement
            const patchData = preparePatchData(rowData, authorName, imageData);
            
            // Create a zip of the docx template
            const zip = new PizZip(arrayBuffer);
            
            // Create a new DocxTemplater instance with the image module
            const doc = new window.docxtemplater();
            
            // Attach the image module if available
            if (imageModule) {
                doc.attachModule(imageModule);
            }
            
            // Configure document
            doc.loadZip(zip);
            doc.setOptions({
                paragraphLoop: true,
                linebreaks: true
            });
            
            // Set the data to be injected
            doc.setData(patchData);
            
            // Perform the template substitution
            doc.render();

            // render image
            doc.render({
                images: imageData
            });
            
            // Generate output
            const output = doc.getZip().generate({
                type: 'blob',
                mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            });
            
            // Save the document
            saveAs(output, `${rowData[1]}.docx`);
            
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
    // Define template fields based on the document structure
    const templateFields = [
        'description', 'entry_time', 'entry_gate', 'exit_time', 'exit_gate',
        'fee', 'paid_amt', 'serial', 'offender_name', 'offender_contact',
        'offender_vehicle', 'offender_store'
    ];
    
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
    
    // Handle image if available
    if (imageData) {
        patchData.images = imageData; // Pass the actual image data (data URL)
    } else {
        patchData.images = ''; // Empty string if no image
    }
    
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
