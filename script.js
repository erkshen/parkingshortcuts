function generateWord() {
    let inputData = document.getElementById("excelData").value.trim();
    let author = document.getElementById("author").value.trim() || "Unknown Author";
    let imageElement = document.getElementById("previewImage");
    let imageSrc = imageElement.src || null;

    if (!inputData) {
        alert("Please paste some data first.");
        return;
    }

    // Split input by tab (Excel uses tabs when copying multiple cells)
    let rowData = inputData.split("\t");

    // Set filename based on first cell (sanitize to remove special characters)
    let fileName = rowData[0].replace(/[^a-zA-Z0-9]/g, "_") || "ExcelRowTable";

    // Placeholder names for first column
    const placeholders = [
        "Name", "Department", "Location", "Store ID", "Sales Amount",
        "Date", "Product ID", "Product Name", "Brand", "Unit Cost",
        "Quantity Sold", "Sale Price", "Discount", "Category", "Region"
    ];

    // If data has more fields than placeholders, generate additional placeholders
    while (placeholders.length < rowData.length) {
        placeholders.push(`Field ${placeholders.length + 1}`);
    }

    // Separate the last 4 rows
    let mainTableRows = rowData.slice(0, -4);
    let lastFourRows = rowData.slice(-4);

    // Create HTML content for the Word document
    let docContent = `<!DOCTYPE html>
        <html xmlns:o="urn:schemas-microsoft-com:office:office"
              xmlns:w="urn:schemas-microsoft-com:office:word"
              xmlns="http://www.w3.org/TR/REC-html40">
        <head><meta charset="utf-8"></head>
        <body>
            <h3>Author: ${author}</h3>

            <table border="1" style="border-collapse: collapse; width: 100%;">
                <tr>
                    <th style="padding: 8px; border: 1px solid black;">Field</th>
                    <th style="padding: 8px; border: 1px solid black;">Data</th>
                </tr>
                <tr>
                    <td style="padding: 8px; border: 1px solid black;">Author</td>
                    <td style="padding: 8px; border: 1px solid black;">${author}</td>
                </tr>`;

    // Add main table rows
    for (let i = 0; i < mainTableRows.length; i++) {
        docContent += `
                <tr>
                    <td style="padding: 8px; border: 1px solid black;">${placeholders[i]}</td>
                    <td style="padding: 8px; border: 1px solid black;">${mainTableRows[i]}</td>
                </tr>`;
    }

    // Add image row if an image is uploaded
    if (imageSrc) {
        docContent += `
                <tr>
                    <td style="padding: 8px; border: 1px solid black;">Uploaded Image</td>
                    <td style="padding: 8px; border: 1px solid black;">
                        <img src="${imageSrc}" style="max-width: 300px; max-height: 150px;">
                    </td>
                </tr>`;
    }

    docContent += `</table><br><br>`;

    // Add separate table for last 4 rows
    if (lastFourRows.length > 0) {
        docContent += `
            <h3>Additional Information</h3>
            <table border="1" style="border-collapse: collapse; width: 100%;">
                <tr>
                    <th style="padding: 8px; border: 1px solid black;">Field</th>
                    <th style="padding: 8px; border: 1px solid black;">Data</th>
                </tr>`;

        for (let i = 0; i < lastFourRows.length; i++) {
            let index = mainTableRows.length + i;
            docContent += `
                <tr>
                    <td style="padding: 8px; border: 1px solid black;">${placeholders[index]}</td>
                    <td style="padding: 8px; border: 1px solid black;">${lastFourRows[i]}</td>
                </tr>`;
        }

        docContent += `</table>`;
    }

    docContent += `</body></html>`;

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

    // Cleanup
    setTimeout(() => {
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
    }, 100);
}
