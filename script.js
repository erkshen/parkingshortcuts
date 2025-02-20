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

    const { Document, Packer, Paragraph, Table, TableRow, TableCell, TextRun, WidthType, ImageRun } = window.docx;

    function createRow(label, value) {
        return new TableRow({
            children: [
                new TableCell({ children: [new Paragraph(label)] }),
                new TableCell({ children: [new Paragraph(value)] })
            ]
        });
    }

    let tableRows = rowData.map((data, index) => createRow(placeholders[index], data));

    let imageBase64 = null;
    if (imageFile) {
        imageBase64 = await toBase64(imageFile);
        tableRows.push(new TableRow({
            children: [
                new TableCell({ children: [new Paragraph("Uploaded Image")] }),
                new TableCell({
                    children: [new ImageRun({
                        data: imageBase64,
                        transformation: { width: 300, height: 150 }
                    })]
                })
            ]
        }));
    }

    const doc = new Document({
        sections: [{
            children: [
                new Paragraph({ text: "HIGH RISK MANOEUVRE REPORT", heading: window.docx.HeadingLevel.HEADING_1 }),
                new Table({ rows: tableRows })
            ]
        }]
    });

    Packer.toBlob(doc).then(blob => {
        let url = URL.createObjectURL(blob);
        let link = document.createElement("a");
        link.href = url;
        link.download = `${fileName}.docx`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    });
}

function toBase64(file) {
    return new Promise((resolve, reject) => {
        let reader = new FileReader();
        reader.readAsArrayBuffer(file);
        reader.onload = () => resolve(reader.result);
        reader.onerror = error => reject(error);
    });
}
