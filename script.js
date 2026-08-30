// ======================================================
// PDF2EXCEL - PDF TO EXCEL CONVERTER
// ======================================================

// ------------------------------
// ELEMENTS
// ------------------------------

const pdfFile = document.getElementById("pdfFile");
const browseBtn = document.getElementById("browseBtn");
const dropArea = document.getElementById("dropArea");

const fileInfo = document.getElementById("fileInfo");
const fileName = document.getElementById("fileName");
const fileSize = document.getElementById("fileSize");

const removeFile = document.getElementById("removeFile");

const convertBtn = document.getElementById("convertBtn");

const progressBox = document.getElementById("progressBox");
const progress = document.getElementById("progress");
const progressPercent = document.getElementById("progressPercent");
const statusText = document.getElementById("statusText");

const downloadBox = document.getElementById("downloadBox");
const downloadBtn = document.getElementById("downloadBtn");
const newFileBtn = document.getElementById("newFileBtn");

const themeBtn = document.getElementById("themeBtn");


// ------------------------------
// VARIABLES
// ------------------------------

let selectedFile = null;
let generatedWorkbook = null;


// ------------------------------
// PDF.JS WORKER
// ------------------------------

if (typeof pdfjsLib !== "undefined") {

    pdfjsLib.GlobalWorkerOptions.workerSrc =
        "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

}


// ======================================================
// BROWSE FILE
// ======================================================

browseBtn.addEventListener("click", () => {

    pdfFile.click();

});


// ======================================================
// SELECT FILE
// ======================================================

pdfFile.addEventListener("change", () => {

    if (pdfFile.files && pdfFile.files.length > 0) {

        handleFile(pdfFile.files[0]);

    }

});


// ======================================================
// DRAG OVER
// ======================================================

dropArea.addEventListener("dragover", (event) => {

    event.preventDefault();

    dropArea.classList.add("dragover");

});


// ======================================================
// DRAG LEAVE
// ======================================================

dropArea.addEventListener("dragleave", () => {

    dropArea.classList.remove("dragover");

});


// ======================================================
// DROP
// ======================================================

dropArea.addEventListener("drop", (event) => {

    event.preventDefault();

    dropArea.classList.remove("dragover");

    const file = event.dataTransfer.files[0];

    handleFile(file);

});


// ======================================================
// HANDLE FILE
// ======================================================

function handleFile(file) {

    if (!file) {
        return;
    }


    const isPDF =
        file.type === "application/pdf" ||
        file.name.toLowerCase().endsWith(".pdf");


    if (!isPDF) {

        alert("Please select a PDF file.");

        return;

    }


    selectedFile = file;
    generatedWorkbook = null;


    fileName.textContent = file.name;

    fileSize.textContent =
        formatFileSize(file.size);


    fileInfo.classList.remove("hidden");

    convertBtn.classList.remove("hidden");

    progressBox.classList.add("hidden");

    downloadBox.classList.add("hidden");


    progress.style.width = "0%";

    progressPercent.textContent = "0%";

    statusText.textContent = "Ready to convert";

}


// ======================================================
// FILE SIZE
// ======================================================

function formatFileSize(bytes) {

    if (bytes === 0) {
        return "0 Bytes";
    }


    const units = [
        "Bytes",
        "KB",
        "MB",
        "GB"
    ];


    const index =
        Math.floor(
            Math.log(bytes) / Math.log(1024)
        );


    return (
        parseFloat(
            (bytes / Math.pow(1024, index))
                .toFixed(2)
        )
        + " "
        + units[index]
    );

}


// ======================================================
// REMOVE FILE
// ======================================================

removeFile.addEventListener("click", resetApp);


// ======================================================
// CONVERT PDF
// ======================================================

convertBtn.addEventListener("click", async () => {

    if (!selectedFile) {

        alert("Please select a PDF file.");

        return;

    }


    if (
        typeof pdfjsLib === "undefined" ||
        typeof XLSX === "undefined"
    ) {

        alert(
            "Required conversion libraries could not be loaded. Please refresh the page and try again."
        );

        return;

    }


    try {

        convertBtn.disabled = true;

        progressBox.classList.remove("hidden");

        downloadBox.classList.add("hidden");


        progress.style.width = "5%";

        progressPercent.textContent = "5%";

        statusText.textContent = "Reading PDF...";


        // ------------------------------
        // READ PDF
        // ------------------------------

        const arrayBuffer =
            await selectedFile.arrayBuffer();


        const pdf =
            await pdfjsLib
                .getDocument({
                    data: arrayBuffer
                })
                .promise;


        // ------------------------------
        // CREATE WORKBOOK
        // ------------------------------

        generatedWorkbook =
            XLSX.utils.book_new();


        // ------------------------------
        // PROCESS EACH PAGE
        // ------------------------------

        for (
            let pageNumber = 1;
            pageNumber <= pdf.numPages;
            pageNumber++
        ) {

            const pageProgress =
                10 +
                Math.round(
                    (pageNumber / pdf.numPages) * 80
                );


            progress.style.width =
                pageProgress + "%";

            progressPercent.textContent =
                pageProgress + "%";


            statusText.textContent =
                `Processing page ${pageNumber} of ${pdf.numPages}`;


            const page =
                await pdf.getPage(pageNumber);


            const textContent =
                await page.getTextContent();


            // ------------------------------
            // GROUP TEXT BY Y POSITION
            // ------------------------------

            const lines = [];


            textContent.items.forEach((item) => {

                const text =
                    (item.str || "").trim();


                if (!text) {
                    return;
                }


                const x =
                    item.transform[4];


                const y =
                    item.transform[5];


                // Find an existing line
                // with similar Y position

                let line =
                    lines.find(
                        (existingLine) =>
                            Math.abs(existingLine.y - y) < 4
                    );


                if (!line) {

                    line = {
                        y: y,
                        items: []
                    };

                    lines.push(line);

                }


                line.items.push({
                    text: text,
                    x: x
                });

            });


            // ------------------------------
            // SORT LINES TOP TO BOTTOM
            // ------------------------------

            lines.sort(
                (a, b) => b.y - a.y
            );


            // ------------------------------
            // CREATE EXCEL DATA
            // ------------------------------

            const excelData = [];


            lines.forEach((line) => {

                line.items.sort(
                    (a, b) => a.x - b.x
                );


                const row =
                    line.items.map(
                        (item) => item.text
                    );


                if (row.length > 0) {

                    excelData.push(row);

                }

            });


            // ------------------------------
            // EMPTY PAGE HANDLING
            // ------------------------------

            if (excelData.length === 0) {

                excelData.push([
                    "No selectable text found on this page."
                ]);

            }


            // ------------------------------
            // CREATE WORKSHEET
            // ------------------------------

            const worksheet =
                XLSX.utils.aoa_to_sheet(
                    excelData
                );


            // ------------------------------
            // COLUMN WIDTH
            // ------------------------------

            const maxColumns =
                Math.max(
                    ...excelData.map(
                        row => row.length
                    ),
                    1
                );


            worksheet["!cols"] = [];


            for (
                let i = 0;
                i < maxColumns;
                i++
            ) {

                worksheet["!cols"].push({
                    wch: 22
                });

            }


            // ------------------------------
            // ADD SHEET
            // ------------------------------

            XLSX.utils.book_append_sheet(
                generatedWorkbook,
                worksheet,
                `Page ${pageNumber}`
            );

        }


        // ------------------------------
        // COMPLETE
        // ------------------------------

        progress.style.width = "100%";

        progressPercent.textContent = "100%";

        statusText.textContent =
            "Conversion Complete!";


        setTimeout(() => {

            downloadBox.classList.remove(
                "hidden"
            );

        }, 400);

    }


    catch (error) {

        console.error(
            "PDF conversion error:",
            error
        );


        generatedWorkbook = null;


        progressBox.classList.add("hidden");


        alert(
            "Unable to convert this PDF. Please try another PDF file."
        );

    }


    finally {

        convertBtn.disabled = false;

    }

});


// ======================================================
// DOWNLOAD EXCEL
// ======================================================

downloadBtn.addEventListener("click", () => {

    if (!generatedWorkbook) {

        alert(
            "Please convert a PDF first."
        );

        return;

    }


    let outputName =
        selectedFile
            ? selectedFile.name
            : "converted";


    outputName =
        outputName.replace(
            /\.pdf$/i,
            ""
        );


    outputName += ".xlsx";


    XLSX.writeFile(
        generatedWorkbook,
        outputName
    );

});


// ======================================================
// NEW FILE
// ======================================================

newFileBtn.addEventListener(
    "click",
    resetApp
);


// ======================================================
// RESET
// ======================================================

function resetApp() {

    selectedFile = null;

    generatedWorkbook = null;


    pdfFile.value = "";


    fileName.textContent =
        "filename.pdf";


    fileSize.textContent =
        "0 MB";


    fileInfo.classList.add("hidden");

    convertBtn.classList.add("hidden");

    progressBox.classList.add("hidden");

    downloadBox.classList.add("hidden");


    progress.style.width = "0%";

    progressPercent.textContent = "0%";

    statusText.textContent =
        "Preparing...";

}


// ======================================================
// DARK MODE
// ======================================================

function applyTheme(isDark) {

    if (isDark) {

        document.body.classList.add("dark");

        themeBtn.textContent = "☀️";

    }
    else {

        document.body.classList.remove("dark");

        themeBtn.textContent = "🌙";

    }

}


themeBtn.addEventListener("click", () => {

    const isDark =
        !document.body.classList.contains("dark");


    applyTheme(isDark);


    localStorage.setItem(
        "pdf2excel-theme",
        isDark ? "dark" : "light"
    );

});


// ======================================================
// LOAD SAVED THEME
// ======================================================

const savedTheme =
    localStorage.getItem(
        "pdf2excel-theme"
    );


if (savedTheme === "dark") {

    applyTheme(true);

}
else {

    applyTheme(false);

}