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

    // Load the .docx template
    let response = await fetch("High Risk Manoeuvre Template.docx");
    let arrayBuffer = await response.arrayBuffer();

    // Load docx library
    let doc = await window.docx.load(arrayBuffer);

    // Define placeholders and replacements
    const placeholders = {
        "{{PROPERTY}}": rowData[0] || "N/A",
        "{{DESCRIPTION}}": rowData[1] || "N/A",
        "{{DATE_TIME_ENTRY}}": rowData[2] || "N/A",
        "{{ENTRY_GATE}}": rowData[3] || "N/A",
        "{{DATE_TIME_EXIT}}": rowData[4] || "N/A",
        "{{EXIT_GATE}}": rowData[5] || "N/A",
        "{{PARKING_FEE}}": rowData[6] || "N/A",
        "{{PARKING_FEE_PAID}}": rowData[7] || "N/A",
        "{{SERIAL_OFFENDER}}": rowData[8] || "N/A",
        "{{REPORT_AUTHOR}}": author,
        "{{OFFENDER_NAME}}": rowData[9] || "N/A",
        "{{CONTACT_DETAILS}}": rowData[10] || "N/A",
        "{{VEHICLE_DETAILS}}": rowData[11] || "N/A",
        "{{STORE_NAME}}": rowData[12] || "N/A"
    };

    // Replace placeholders in the document
    doc.replaceText(Object.keys(placeholders), Object.values(placeholders));

    // If an image is uploaded, insert it in the correct cell
    if (imageFile) {
        let imageBase64 = await toBase64(imageFile);
        let imageBuffer = Uint8Array.from(atob(imageBase64.split(",")[1]), c => c.charCodeAt(0));

        doc.insertImage({
            data: imageBuffer,
            width: 300, // Image width
            height: 150, // Image height
            targetPlaceholder: "{{PHOTOGRAPHS_CCTV_FOOTAGE}}"
        });
    }

    // Generate the final .docx file
    let newDocBlob = await doc.saveAsBlob();
    let downloadUrl = URL.createObjectURL(newDocBlob);

    // Create download link
    let link = document.createElement("a");
    link.href = downloadUrl;
    link.download = `${fileName}.docx`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

// Convert image file to Base64
function toBase64(file) {
    return new Promise((resolve, reject) => {
        let reader = new FileReader();
        reader.readAsDataURL(file);
        reader.onload = () => resolve(reader.result);
        reader.onerror = error => reject(error);
    });
}
