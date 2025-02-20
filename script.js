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

    // Load the .docx template
    let response = await fetch("High Risk Manoeuvre Template.docx");
    let blob = await response.blob();
    let zip = await JSZip.loadAsync(blob);

    // Read and modify document.xml (main text content)
    let docXml = await zip.file("word/document.xml").async("string");

    // Replace placeholders safely using Word-compatible XML formatting
    Object.keys(placeholders).forEach(key => {
        let safeKey = key.replace(/[\{\}]/g, ""); // Remove curly braces in XML searches
        docXml = docXml.replace(new RegExp(safeKey, "g"), placeholders[key]);
    });

    // Save modified XML back into the .docx
    zip.file("word/document.xml", docXml);

    // If an image is uploaded, add it to the document
    if (imageFile) {
        let imageBase64 = await toBase64(imageFile);
        let imgData = imageBase64.split(",")[1]; // Remove the data type prefix

        // Add the image inside the .docx archive
        zip.file("word/media/image1.png", imgData, { base64: true });

        // Update document.xml.rels to reference the new image
        let relsXml = await zip.file("word/_rels/document.xml.rels").async("string");
        let newRelId = `rId${Date.now()}`; // Unique ID for the image reference
        let imageRel = `
            <Relationship Id="${newRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>`;
        relsXml = relsXml.replace("</Relationships>", `${imageRel}</Relationships>`);
        zip.file("word/_rels/document.xml.rels", relsXml);

        // Insert the image reference inside document.xml
        let imageTag = `<w:drawing><wp:inline><wp:extent cx="5000000" cy="5000000"/>
            <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                    <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                        <pic:blipFill><a:blip r:embed="${newRelId}"/></pic:blipFill>
                    </pic:pic>
                </a:graphicData>
            </a:graphic>
        </wp:inline></w:drawing>`;

        // Insert image after the "Photographs/ CCTV Footage" placeholder
        docXml = docXml.replace("{{PHOTOGRAPHS_CCTV_FOOTAGE}}", imageTag);
        zip.file("word/document.xml", docXml);
    }

    // Generate and download the modified .docx file
    let modifiedBlob = await zip.generateAsync({ type: "blob" });
    let downloadUrl = URL.createObjectURL(modifiedBlob);

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
