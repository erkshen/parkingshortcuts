// script.js - With Image Module Support
async function read_text() {
	const worker = await Tesseract.createWorker('eng');
	const previewImage = document.getElementById('previewImage');

	if (previewImage !== undefined) {
		(async () => {
		  const { data: { text } } = await worker.recognize(previewImage.src);
		  console.log(text);
		  await worker.terminate();
		})();
	};
}
