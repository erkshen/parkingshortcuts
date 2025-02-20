async function generateWord() {
    let inputData = document.getElementById("excelData").value.trim();
    let author = document.getElementById("author").value.trim() || "Unknown Author";
    let imageFile = document.getElementById("imageUpload").files[0];

    if (!inputData) {
        alert("Please paste some data first.");
        return;
    }

    let rowData = inputData.split("\t");
    let fileName = rowData[0].replace(/[^a-zA-Z0-9]/g, "_") || "High_Risk_Manoeuvre_Report";

    const placeholders = [
        "Property", "Description", "Date & Time of Entry", "Entry Gate",
        "Date & Time of Exit", "Exit Gate", "Parking Fee", "Parking Fee Paid",
        "Serial Offender", "Report Author", "Photographs/ CCTV Footage",
        "Offender Name", "Contact Details", "Vehicle Details", "Store Name"
    ];

    while (placeholders.length < rowData.length) {
        placeholders.push(`Field ${placeholders.length + 1}`);
    }

    let tableHTML = `<table border="1" style="width: 100%; border-collapse: collapse;">`;

    for (let i = 0; i < rowData.length; i++) {
        tableHTML += `
            <tr>
                <td style="padding: 8px; border: 1px solid black; font-weight: bold;">${placeholders[i]}</td>
                <td style="padding: 8px; border: 1px solid black;">${rowData[i]}</td>
            </tr>`;
    }

    // If an image is uploaded, include it
    if (imageFile) {
        let imageBase64 = await toBase64(imageFile);
        tableHTML += `
            <tr>
                <td style="padding: 8px; border: 1px solid black; font-weight: bold;">Uploaded Image</td>
                <td style="padding: 8px; border: 1px solid black;">
                    <img src="${imageBase64}" width="300">
                </td>
            </tr>`;
    }

    tableHTML += `</table>`;

    let docContent = `
        <html xmlns:o="urn:schemas-microsoft-com:office:office"
              xmlns:w="urn:schemas-microsoft-com:office:word"
              xmlns="http://www.w3.org/TR/REC-html40">
        <head><meta charset="utf-8"></head>
        <body>
            <h2 style="text-align: center;">HIGH RISK MANOEUVRE REPORT</h2>
            <h3>Author: ${author}</h3>
            ${tableHTML}
        </body>
        </html>`;

    let blob = new Blob(["\ufeff" + docContent], { type: "application/msword" });
    let url = URL.createObjectURL(blob);

    let link = document.createElement("a");
    link.href = url;
    link.download = `${fileName}.doc`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

// Convert image file to base64
function toBase64(file) {
    return new Promise((resolve, reject) => {
        let reader = new FileReader();
        reader.readAsDataURL(file);
        reader.onload = () => resolve(reader.result);
        reader.onerror = error => reject(error);
    });
}
