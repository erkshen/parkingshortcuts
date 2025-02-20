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

    // Define placeholders that will be replaced in the document
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

    // Read the document.xml file
    let docXml = await zip.file("word/document.xml").async("string");

    // Ensure replacements are within `<w:t>` tags correctly
    Object.keys(placeholders).forEach(key => {
        let safeKey = key.replace(/[\{\}]/g, "");
        let regex = new RegExp(`(<w:t>\\s*)${safeKey}(\\s*</w:t>)`, "g");
        docXml = docXml.replace(regex, `$1${placeholders[key]}$2`);
    });

    // Save modified XML back into .docx
    zip.file("word/document.xml", docXml);

    // If an image is uploaded, insert it properly
    if (imageFile) {
        let imageBase64 = await toBase64(imageFile);
        let imgData = imageBase64.split(",")[1]; // Remove base64 prefix

        // Store the image inside "word/media/"
        zip.file("word/media/uploadedImage.png", imgData, { base64: true });

        // Modify relationships (.rels file) properly
        let relsXml = await zip.file("word/_rels/document.xml.rels").async("string");
        let newRelId = `rId${Date.now()}`;
        let imageRel = `
            <Relationship Id="${newRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/uploadedImage.png"/>`;
        relsXml = relsXml.replace("</Relationships>", `${imageRel}</Relationships>`);
        zip.file("word/_rels/document.xml.rels", relsXml);

        // Insert the image in a valid table cell inside `document.xml`
        let imageTag = `
            <w:p>
                <w:r>
                    <w:drawing>
                        <wp:inline>
                            <wp:extent cx="5000000" cy="5000000"/>
                            <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                                <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                    <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                        <pic:blipFill><a:blip r:embed="${newRelId}"/></pic:blipFill>
                                    </pic:pic>
                                </a:graphicData>
                            </a:graphic>
                        </wp:inline>
                    </w:drawing>
                </w:r>
            </w:p>`;

        // Ensure the image is added **inside the correct table cell**
        docXml = docXml.replace("{{PHOTOGRAPHS_CCTV_FOOTAGE}}", imageTag);
        zip.file("word/document.xml", docXml);
    }

    // Generate and download the final `.docx` file
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
