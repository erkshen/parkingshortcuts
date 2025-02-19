function generateWord() {
    let inputData = document.getElementById("excelData").value.trim();

    if (!inputData) {
        alert("Please paste some data first.");
        return;
    }

    // Split input by tab (Excel uses tabs when copying multiple cells)
    let rowData = inputData.split("\t");

    // Create a Word document using a basic Blob and HTML structure
    let docContent = `
        <html xmlns:o="urn:schemas-microsoft-com:office:office" 
              xmlns:w="urn:schemas-microsoft-com:office:word" 
              xmlns="http://www.w3.org/TR/REC-html40">
        <head><meta charset="utf-8"></head>
        <body>
            <table border="1" style="border-collapse: collapse; width: 100%;">
                <tr>
                    ${rowData.map(cell => `<td style="padding: 8px;">${cell}</td>`).join("")}
                </tr>
            </table>
        </body>
        </html>`;

    let blob = new Blob([docContent], { type: "application/msword" });
    let link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "ExcelRowTable.doc";
    document.body.appendChild(link
