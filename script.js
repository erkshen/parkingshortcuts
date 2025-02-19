function generateWord() {
    let inputData = document.getElementById("excelData").value.trim();

    if (!inputData) {
        alert("Please paste some data first.");
        return;
    }

    // Split input by tab (Excel uses tabs when copying multiple cells)
    let rowData = inputData.split("\t");

    // Set filename based on first cell (sanitize to remove special characters)
    let fileName = rowData[0].replace(/[^a-zA-Z0-9]/g, "_") || "ExcelRowTable";
    
    // Generate the alphabet column (A-Z)
    const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    
    // Create an HTML structure for the Word document
    let docContent = `<!DOCTYPE html>
        <html xmlns:o="urn:schemas-microsoft-com:office:office"
              xmlns:w="urn:schemas-microsoft-com:office:word"
              xmlns="http://www.w3.org/TR/REC-html40">
        <head><meta charset="utf-8"></head>
        <body>
            <table border="1" style="border-collapse: collapse; width: 100%;">
                <tr>
                    <th style="padding: 8px; border: 1px solid black;">Letter</th>
                    <th style="padding: 8px; border: 1px solid black;">Data</th>
                </tr>`;

    // Populate table with alphabet letters and corresponding Excel data
    for (let i = 0; i < rowData.length; i++) {
        let letter = alphabet[i] || `Extra${i+1}`; // Fallback for >26 columns
        docContent += `
                <tr>
                    <td style="padding: 8px; border: 1px solid black;">${letter}</td>
                    <td style="padding: 8px; border: 1px solid black;">${rowData[i]}</td>
                </tr>`;
    }

    docContent += `
            </table>
        </body>
        </html>`;

    // Create a Blob object with proper MIME type
    let blob = new Blob(["\ufeff" + docContent], { type: "application/msword" });

    // Generate a URL for the blob
    let url = URL.createObjectURL(blob);

    // Create a temporary download link
    let link = document.createElement("a");
    link.href = url;
    link.download = `${fileName}.doc`;

    // Append to document and trigger click
    document.body.appendChild(link);
    link.click();

    // Cleanup: Remove the link and revoke the Blob URL after a short delay
    setTimeout(() => {
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
    }, 100);
}
