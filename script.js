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

    // Placeholder mappings for template replacement
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

    // Read document.xml (main text content inside the .docx)
    let docXml = await zip.file("word/document.xml").async("string");

    // **Correctly replace placeholders inside Word's `<w:t>` elements**
    Object.keys(placeholders).forEach(key => {
        let safeKey = key.replace(/[\{\}]/g, ""); // Remove curly braces in XML searches
        let regex = new RegExp(`<w:t>\\s*${safeKey}\\s*</w:t>`, "g"); // Match full `<w:t>` blocks
        docXml = docXml.replace(regex, `<w:t>${placeholders[key]}</w:t>`);
    });

    // **Save modified document.xml back to the .docx**
    zip.file("word/document.xml", docXml);

    // **If an image is uploaded, insert it inside a table cell**
    if (imageFile) {
        let imageBase64 = await toBase64(imageFile);
        let imgData = imageBase64.split(",")[1]; // Remove the `data:image/png;base64,` prefix

        // **Embed the image inside word/media**
        let imageFileName = "word/media/uploadedImage.png";
        zip.file(imageFileName, imgData, { base64: true });

        // **Create a new relationship in document.xml.rels**
        let relsXml = await zip.file("word/_rels/document.xml.rels").async("string");
        let newRelId = `rId${Date.now()}`; // Unique ID for the new image
        let imageRel = `
            <Relationship Id="${newRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/uploadedImage.png"/>`;
        relsXml = relsXml.replace("</Relationships>", `${imageRel}</Relationships>`);
        zip.file("word/_rels/document.xml.rels", relsXml);

        // **Insert the image inside the correct table cell**
        let imageTag = `<w:tc>
            <w:p><w:r><w:drawing><wp:inline><wp:extent cx="5000000" cy="5000000"/>
                <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                        <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                            <pic:blipFill><a:blip r:embed="${newRelId}"/></pic:blipFill>
                        </pic:pic>
                    </a:graphicData>
                </a:graphic>
            </wp:inline></w:drawing></w:r></w:p>
        </w:tc>`;

        // **Insert image after the "Photographs/ CCTV Footage" table row**
        docXml = docXml.replace("</w:tr>", `${imageTag}</w:tr>`);
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
