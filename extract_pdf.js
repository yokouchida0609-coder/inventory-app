const fs = require('fs');
const pdf = require('pdf-parse');

async function extractPDF(filePath, label) {
    const dataBuffer = fs.readFileSync(filePath);
    const data = await pdf(dataBuffer);
    console.log(`\n${'='.repeat(60)}`);
    console.log(`${label} (${data.numpages}ページ)`);
    console.log('='.repeat(60));
    console.log(data.text);
}

(async () => {
    await extractPDF('C:\\Users\\yoko7\\Downloads\\心理１日目（前半）.pdf', '心理１日目（前半）');
    await extractPDF('C:\\Users\\yoko7\\Downloads\\心理１日目（後半）.pdf', '心理１日目（後半）');
})().catch(console.error);
