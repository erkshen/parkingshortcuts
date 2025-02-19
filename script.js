async function generateWord() {
    let inputData = document.getElementById("excelData").value.trim();
    let author = document.getElementById("author").value.trim() || "Unknown Author";
    let imageElement = document.getElementById("previewImage");
    let imageFile = document.getElementById("imageUpload").files[0];

    if (!inputData) {
        alert("Please paste some data first.");
        return;
    }

    // Split input by tab (Excel uses tabs when copying multiple cells)
    let rowData = inputData.split("\t");

    // Set filename based on first cell
    let fileName = rowData[0].replace(/[^a-zA-Z0-9]/g, "_") || "ExcelRowTable";

    // Placeholder names for first column
    const placeholders = [
        "Description", "Entry Date and Time", "Entry Gate", "Exit Date and Time", "Exit Gate",
        "Fee Due", "Paid Amount", "Repeat Offender", "Report Author", "Supporting Images",
        "Name", "Vehicle Details", "Contact", "Store Name"
    ];

    // Ensure placeholders cover all fields
    while (placeholders.length < rowData.length) {
        placeholders.push(`Field ${placeholders.length + 1}`);
    }

    // Separate last 4 rows
    let mainTableRows = rowData.slice(0, -4);
    let lastFourRows = rowData.slice(-4);

    // Import docx library
    const { Document, Packer, Paragraph, Table, TableRow, TableCell, TextRun, WidthType, AlignmentType, ImageRun } = docx;

    // Function to create a table row
    function createRow(label, value) {
        return new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph(label)],
                    width: { size: 40, type: WidthType.PERCENTAGE }
                }),
                new TableCell({
                    children: [new Paragraph(value)],
                    width: { size: 60, type: WidthType.PERCENTAGE }
                })
            ]
        });
    }

    // Create main table rows
    let tableRows = [
        createRow("Author", author),
        ...mainTableRows.map((data, index) => createRow(placeholders[index], data))
    ];

    // Load image as base64 if provided
    let imageBase64 = null;
    if (imageFile) {
        imageBase64 = await toBase64(imageFile);
        tableRows.push(new TableRow({
            children: [
                new TableCell({ children: [new Paragraph("Uploaded Image")] }),
                new TableCell({
                    children: [new Paragraph(" "), new ImageRun({
                        data: imageBase64,
                        transformation: { width: 300, height: 150 }
                    })]
                })
            ]
        }));
    }

    // Create the additional table for last 4 rows
    let additionalTableRows = lastFourRows.map((data, index) => createRow(placeholders[mainTableRows.length + index], data));

    // Create the document
    const doc = new Document({
        sections: [
            {
                children: [
                    new Paragraph({ text: "Main Data", heading: docx.HeadingLevel.HEADING_2 }),
                    new Table({
                        rows: tableRows,
                        width: { size: 100, type: WidthType.PERCENTAGE }
                    }),
                    new Paragraph({ text: " " }), // Spacer
                    new Paragraph({ text: "Additional Information", heading: docx.HeadingLevel.HEADING_2 }),
                    new Table({
                        rows: additionalTableRows,
                        width: { size: 100, type: WidthType.PERCENTAGE }
                    })
                ]
            }
        ]
    });

    // Generate and download the document
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

// Convert image file to base64
function toBase64(file) {
    return new Promise((resolve, reject) => {
        let reader = new FileReader();
        reader.readAsArrayBuffer(file);
        reader.onload = () => resolve(reader.result);
        reader.onerror = error => reject(error);
    });
}
