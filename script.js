function generateWord() {
    let inputData = document.getElementById("excelData").value.trim();

    if (!inputData) {
        alert("Please paste some data first.");
        return;
    }

    // Split input by tab (Excel uses tabs when copying multiple cells)
    let rowData = inputData.split("\t");

    // Create an HTML structure for the Word document
    let docContent = `<!DOCTYPE html>
        <html xmlns:o="urn:schemas-microsoft-com:office:office"
              xmlns:w="urn:schemas-microsoft-com:office:word"
              xmlns="http://www.w3.org/TR/REC-html40">
        <head><meta charset="utf-8"></head>
        <body>
            <table border="1" style="border-collapse: collapse; width: 100%;">
                <tr>
                    ${rowData.map(cell => `<td style="padding: 8px; border: 1px solid black;">${cell}</td>`).join("")}
                </tr>
            </table>
        </body>
        </html>`;

    // Create a Blob object with proper MIME type
    let blob = new Blob(["\ufeff" + docContent], { type: "application/msword" });

    // Generate a URL for the blob
    let url = window.URL.createObjectURL(blob);

    // Create a temporary download link
    let link = document.createElement("a");
    link.href = url;
    link.download = "ExcelRowTable.doc";

    // Append to document and trigger click
    document.body.appendChild(link);
    link.click();

    // Cleanup: Remove the link and revoke the Blob URL after a delay
    setTimeout(() => {
        document.body.removeChild(link);
        window.URL.revokeObjectURL(url);
    }, 100);
}
