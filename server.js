const convert = require("docx-pdf");
const fs = require('fs');
const pdfParse = require('pdf-parse');
const { Document, Packer, Paragraph } = require('docx');
const pdf2excel = require('pdf-to-excel');
const XLSX = require('xlsx');
const PDF = require('pdfkit');
const pptxgen = require("pptxgenjs");
const officeParser = require("officeparser");
const { PDFDocument } = require('pdf-lib');
const { compressPDF } = require('pdf-compress-split-utility');
const sharp = require("sharp");
const poppler = require('pdf-poppler');
const imageToPdf = require('image-to-pdf');
const html = require('html-pdf');
const ffmpeg = require('fluent-ffmpeg');
const ffmpegPath = require('ffmpeg-static');
const axios = require('axios');
require('dotenv').config();

async function doctopdf() {
    try {
        await convert("./files/doc1.docx", "./files/doc1.pdf", function (err, result) {
            if (err) throw err;
            console.log(result);
        });
    } catch (error) {
        console.error("Error in doctopdf:", error.message);
    }
}
// doctopdf();

async function pdftodocx() {
    try {
        const dataBuffer = fs.readFileSync('./files/doc1.pdf');
        const data = await pdfParse(dataBuffer);
        const doc = new Document({
            sections: [
                {
                    properties: {
                        title: "Sample Document"
                    },
                    children: [new Paragraph(data.text)],
                },
            ],
        });
        const buffer = await Packer.toBuffer(doc);
        fs.writeFileSync('./files/doc3.docx', buffer);
        console.log('DOCX file created successfully.');
    } catch (error) {
        console.error("Error in pdftodocx:", error.message);
    }
}
// pdftodocx();

async function pdftoexcel() {
    try {
        const options = {
            type: 'buffer',
            bookType: 'xlsx'
        };
        await pdf2excel.genXlsx('./files/excel1.pdf', './files/excel2.xlsx', options);
    } catch (error) {
        console.error("Error in pdftoexcel:", error.message);
    }
}
// pdftoexcel();

async function excltopdf() {
    try {
        const workbook = XLSX.readFile('./files/excel1.xlsx');
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(sheet);
        const doc = new PDF();
        doc.pipe(fs.createWriteStream('./files/excel1.pdf'));
        doc.fontSize(25).text('Excel Data');
        data.forEach((row) => {
            doc.fontSize(12).text(`${JSON.stringify(row)}`);
        });
        doc.end();
        console.log('Excel converted to PDF successfully.');
    } catch (error) {
        console.error("Error in excltopdf:", error.message);
    }
}
// excltopdf();


async function pdftoppt() {
    try {
        function splitTextIntoSlides(text, maxLinesPerSlide, charsPerLine) {
            const words = text.split(' ');
            let slides = [];
            let slideText = '';
            let currentLines = 0;
            words.forEach(word => {
                const currentLineLength = slideText.split('\n').pop().length;
                if (currentLineLength + word.length > charsPerLine) {
                    slideText += '\n';
                    currentLines++;
                }
                if (currentLines >= maxLinesPerSlide) {
                    slides.push(slideText.trim());
                    slideText = '';
                    currentLines = 0;
                }
                slideText += word + ' ';
            });
            if (slideText.length > 0) {
                slides.push(slideText.trim());
            }
            return slides;
        }

        const dataBuffer = fs.readFileSync('./files/doc1.pdf');
        const data = await pdfParse(dataBuffer);
        const text = data.text;
        const ppt = new pptxgen();
        const maxLinesPerSlide = 8;
        const charsPerLine = 40;
        const textSlides = splitTextIntoSlides(text, maxLinesPerSlide, charsPerLine);
        textSlides.forEach(slideText => {
            let slide = ppt.addSlide();
            slide.addText(slideText, {
                x: 0.5, y: 0.5, w: '90%', h: '90%',
                fontSize: 18, wrap: true, valign: 'top'
            });
        });
        const fileName = './files/newpdf.pptx';
        await ppt.writeFile(fileName);
        console.log('PPT generated successfully:', fileName);
    } catch (error) {
        console.error("Error in pdftoppt:", error.message);
    }
}
// pdftoppt();



async function ppttopdf() {
    try {
        officeParser.parseOffice('./files/newpdf.pptx', async function (data, err) {
            if (err) throw err;
            const doc = new PDF();
            doc.pipe(fs.createWriteStream('./files/newppt.pdf'));
            doc.text(data, 10, 10);
            doc.end();
            console.log('PPT converted to PDF successfully.');
        });
    } catch (error) {
        console.error("Error in ppttopdf:", error.message);
    }
}
// ppttopdf();

async function merging() {
    try {
        const pdf1Bytes = fs.readFileSync('./files/doc1.pdf');
        const pdf2Bytes = fs.readFileSync('./files/excel1.pdf');
        const mergedPdf = await PDFDocument.create();
        const pdf1 = await PDFDocument.load(pdf1Bytes);
        const pdf2 = await PDFDocument.load(pdf2Bytes);
        const copiedPages1 = await mergedPdf.copyPages(pdf1, pdf1.getPageIndices());
        const copiedPages2 = await mergedPdf.copyPages(pdf2, pdf2.getPageIndices());
        copiedPages1.forEach(page => mergedPdf.addPage(page));
        copiedPages2.forEach(page => mergedPdf.addPage(page));
        const mergedPdfBytes = await mergedPdf.save();
        fs.writeFileSync('./files/merged.pdf', mergedPdfBytes);
        console.log('Successfully merged!');
    } catch (error) {
        console.error('Error during merge:', error);
    }
}
// merging();

async function comPDF() {
    const KB = 1024;
    try {
        await compressPDF('./files/doc1.pdf', 450 * KB);
        console.log('PDF pages compressed successfully');
    } catch (error) {
        console.error("Error in comPDF:", error.message);
    }
}
// comPDF();

async function comPIC() {
    try {
        await sharp('./files/1.jpg')
            .resize({ width: 800 })
            .jpeg({ quality: 80 })
            .toFile('./files/2.jpg');
        console.log('Image compressed successfully.');
    } catch (error) {
        console.error("Error in comPIC:", error.message);
    }
}
// comPIC();

async function pdftopic() {
    try {
        const pdfPath = './files/doc1.pdf';
        const outputPath = './files/';
        const options = {
            format: 'jpg',
            out_dir: outputPath,
            out_prefix: 'doc1',
            page: null
        };
        await poppler.convert(pdfPath, options);
        console.log('PDF converted to images successfully.');
    } catch (error) {
        console.error("Error in pdftopic:", error.message);
    }
}
// pdftopic();

async function pictopdf() {
    try {
        const images = ['./files/doc1-1.jpg', './files/doc1-2.jpg'];
        const pageSize = [595.28, 841.89];
        await imageToPdf(images, pageSize)
            .pipe(fs.createWriteStream('./files/output.pdf'))
            .on('finish', () => {
                console.log('PDF generated successfully.');
            });
    } catch (error) {
        console.error("Error in pictopdf:", error.message);
    }
}
// pictopdf();

async function htmltopdf() {
    try {
        const htpdf = fs.readFileSync('./files/index.html', 'utf8');
        await new Promise((resolve, reject) => {
            html.create(htpdf).toFile('./files/html.pdf', function (err, res) {
                if (err) reject(err);
                else {
                    console.log('PDF generated:', res.filename);
                    resolve();
                }
            });
        });
    } catch (error) {
        console.error("Error in htmltopdf:", error.message);
    }
}
// htmltopdf();

async function videototext() {
    const videoPath = './files/video.mp4';
    const outputPath = './files/audio.mp3';

    ffmpeg.setFfmpegPath(ffmpegPath);

    // Ensure ffmpeg completes before proceeding
    await new Promise((resolve, reject) => {
        ffmpeg(videoPath)
            .output(outputPath)
            .on('end', () => {
                console.log("Audio extraction completed.");
                resolve();
            })
            .on('error', (err) => {
                console.error("FFmpeg error:", err);
                reject(err);
            })
            .run();
    });

    // Confirm audio file exists before proceeding
    if (!fs.existsSync(outputPath)) {
        console.error("Audio file not found after FFmpeg process.");
        return;
    }

    const API_KEY = process.env.API_KEY;    // Here you have to use your's assembly.ai account's api.
    if (!API_KEY) {
        console.error("API key is missing. Please set it in the environment variables.");
        return;
    }

    try {
        // Upload audio to AssemblyAI
        const upload = await axios.post('https://api.assemblyai.com/v2/upload', fs.createReadStream(outputPath), {
            headers: {
                authorization: API_KEY,
            },
        });
        const audioUrl = upload.data.upload_url;

        // Request transcription
        const transcription = await axios.post('https://api.assemblyai.com/v2/transcript', {
            audio_url: audioUrl,
            punctuate: true,
            format_text: true,
        }, {
            headers: {
                authorization: API_KEY,
            },
        });

        const Id = transcription.data.id;
        let Result;

        // Poll for transcription status
        do {
            const Response = await axios.get(`https://api.assemblyai.com/v2/transcript/${Id}`, {
                headers: {
                    authorization: API_KEY,
                },
            });
            Result = Response.data;

            console.log('Transcription status:', Result.status);

            if (Result.status === 'completed') {
                const Text = formatTranscription(Result.text);
                fs.writeFileSync('./files/audio.txt', Text, 'utf8');
                console.log("Transcription saved to audio.txt");
            } else if (Result.status === 'failed') {
                console.error("Transcription failed:", Result.error);
                return;
            }

            // Wait 5 seconds before checking again
            await new Promise(resolve => setTimeout(resolve, 5000));
        } while (Result.status !== 'completed');
    } catch (err) {
        console.error("Error during transcription process:", err);
    }
}

function formatTranscription(text, Length = 80) {
    const words = text.split(' ');
    let current = '';
    const Lines = [];

    for (const word of words) {
        if ((current + word).length <= Length) {
            current += (current.length ? ' ' : '') + word;
        } else {
            Lines.push(current);
            current = word;
        }
    }

    if (current) {
        Lines.push(current);
    }

    return Lines.join('\n');
}
// videototext();
