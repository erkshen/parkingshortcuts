// script.js - With Image Module Support

async function read_text() {
    const worker = await Tesseract.createWorker('eng');
    const previewImage = document.getElementById('previewImage');
    /*await worker.setParameters({
        tessedit_pageseg_mode: PSM.SINGLE_BLOCK,
    });
    */

    if (previewImage) {
        (async () => {
            const { data: { text } } = await worker.recognize(previewImage.src);
            console.log(text);
            await worker.terminate();
        })();
    }
}
