// script.js - With Image Module Support
async function generateWord() {
    try {
        // Get excel data
        const excelDataText = document.getElementById('excelData').value.trim();
        
        // Validate inputs
        if (!excelDataText) {
            alert('Please paste Excel data first.');
            return;
        }
        
        // Parse Excel data
        const rowsData = parseExcelData(excelDataText);
        
        // Load the template document
        const templateUrl = 'High Risk Manoeuvre Template.docx';
        
        const docxType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

        rowsData.forEach(function (rowData, index) {
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
                        author: rowData[15],
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
        });
    } catch (error) {
        console.error('Error generating document:', error);
        alert('Error generating document: ' + error.message);
    }
}

function parseExcelData(excelText) {
	// loop through rows to get individual cells in each row
    const rows = excelText.split('\n').map(function(row) {
        return row.split('\t').map(function(cell) {
        	return cell.trim();
        });
    });
    
    console.log(rows);
    return rows;
}
