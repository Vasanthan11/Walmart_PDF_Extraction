document.getElementById('extractButton').addEventListener('click', extractComments);

document.getElementById('fileInput').addEventListener('change', function(event) {
    var fileCount = event.target.files.length;
    var fileCountText = fileCount > 0 ? fileCount + ' file(s) chosen' : 'No files chosen';
    document.getElementById('fileCount').textContent = fileCountText;

    // Set the upload date when files are selected
    if (fileCount > 0) {
        const uploadDate = formatDate(new Date()); // Format as DD.MM.YYYY
        document.getElementById('uploadDate').value = uploadDate;
    }
});

async function extractComments() {
    const fileInput = document.getElementById('fileInput').files;
    if (fileInput.length === 0) {
        alert('Please upload at least one PDF file.');
        return;
    }

    const uploadDate = document.getElementById('uploadDate').value || formatDate(new Date());

    let comments = [];
    for (const file of fileInput) {
        const pdfData = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: pdfData }).promise;

        // Extract information from the filename
        const fileName = file.name;

        // Extract week number (e.g., WK38) and format it to Week-38
        const weekMatch = fileName.match(/(WK\d+)/); // Match WK followed by digits
        const week = weekMatch ? `Week-${weekMatch[0].substring(2)}` : 'Unknown';

        // Set default banner name to Walmart
        const bannerName = 'Walmart'; 

        // Extract meaningful content between WK..._24_ and (optionally) _CPR or .pdf
        const pageMatch = fileName.match(/^WK\d+_24_(.+?)(?:(_CPR|_PI\s*\d*|_CF|_PR\s*\d*|_PR[1-4](_CF)?|\.pdf))?$/);
        let page = pageMatch ? pageMatch[1] : 'Unknown'; // Get the captured group without unwanted suffixes

        // Log the extracted page for debugging
        console.log(`Extracted page name from "${fileName}": "${page}"`);

        // Remove unwanted suffixes using a regular expression
        page = page.replace(/(_CPR|_PI\s*\d*|_CF|_PR\s*\d*|_PR[1-4](_CF)?|_CF)?$/, '').trim(); // Removing unwanted patterns and trimming whitespace

        // Determine the proof name based on file name contents
        let proofName = 'Press'; // Default proof name now set to 'Press'
        if (fileName.includes('_CPR')) {
            proofName = 'CPR';
        }else if (/(_PR\d)/.test(fileName)) {
            const prMatch = fileName.match(/_PR(\d)/); // Extract the number following _PR
            proofName = `Proof ${prMatch[1]}`; //
        } else if (fileName.includes('_PI')) {
            proofName = 'Press';
        }

        // Log the determined proof name
        console.log(`Proof name for "${fileName}": "${proofName}"`);

        // Determine BI_ENG_All based on file name contents
        let biEngAll = 'All Zones'; // Default value
        const bilZonesPatterns = ['B_ON', 'B_NB', 'B_QC', '_B_MTL_RADDAR'];
        const engZonesPatterns = ['E_ON', 'NB', 'NS', 'PE', 'NL', 'MB', 'SK', 'AB', 'BC', 'E_NAT', 'E_ATL', 'E_WEST', '_E_VAN_RADDAR'];

        // Check if the file name contains any of the Bil Zones patterns
        if (bilZonesPatterns.some(pattern => fileName.includes(pattern))) {
            biEngAll = 'Bil Zones';
        } 
        // Check if the file name contains any of the Eng Zones patterns
        else if (engZonesPatterns.some(pattern => fileName.includes(pattern))) {
            biEngAll = 'Eng Zones';
        }

        // Log the determined BI_ENG_All for debugging
        console.log(`BI_ENG_All for "${fileName}": "${biEngAll}"`);

        for (let i = 0; i < pdf.numPages; i++) {
            const pdfPage = await pdf.getPage(i + 1);
            const annotations = await pdfPage.getAnnotations();

            annotations.forEach(annotation => {
                if (annotation.subtype !== 'Popup') {
                    // Determine error type based on the comment's content
                    let errorType = 'Product_Description'; // Default category
                    let content = annotation.contents || 'No content';
            
                    if (content.toLowerCase().includes('price')) {
                        errorType = 'Price_Point';
                    } else if (content.toLowerCase().includes('alignment')) {
                        errorType = 'Overall_Layout';
                    } else if (content.toLowerCase().includes('image')) {
                        errorType = 'Image_Usage';
                    }
            
                    let errorsContent = '';
                    let gdContent = '';
            
                    if (content.startsWith('GD:')) {
                        gdContent = content.substring(3).trim(); // Remove 'GD:' and trim the content
                    } else {
                        errorsContent = content;
                    }
            
                    // Collect comments with fields matching the desired format
                    comments.push({
                        Date: uploadDate,
                        Banner: bannerName,
                        Week: week,
                        Page: page, // The final extracted page name
                        Proof: proofName, // Use the determined proof name
                        BI_ENG_All: biEngAll, // Use the determined BI_ENG_All
                        PageAssembler: '', // Placeholder, adjust logic as needed
                        QC: annotation.title || 'Unknown',
                        SJC_QC: '', // Placeholder for SJC QC, adjust as needed
                        Correction_Revision: gdContent,
                        NO_OF_ERRORS: 1, // Example placeholder, adjust logic as needed
                        ERROR_CATEGORY: errorType,
                        REMARKS: errorsContent  // Place GD content in the REMARKS column
                    });
                }
            });
        }

        // Add a blank row after each PDF's comments
        comments.push({
            Date: '',
            Banner: '',
            Week: '',
            Page: '',
            Proof: '',
            BI_ENG_All: '',
            PageAssembler: '',
            QC: '',
            SJC_QC: '',
            Correction_Revision: '',
            NO_OF_ERRORS: '',
            ERROR_CATEGORY: '',
            REMARKS: ''
        });
    }

    exportToExcel(comments);
}

function formatDate(date) {
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0'); // Months are zero-based
    const year = date.getFullYear();
    return `${day}.${month}.${year}`; // Format as DD.MM.YYYY
}

function exportToExcel(comments) {
    // Define the worksheet and workbook
    const worksheet = XLSX.utils.json_to_sheet(comments, { header: ["Date", "Banner", "Week", "Page", "Proof", "BI_ENG_All", "PageAssembler", "QC", "SJC_QC", "Correction_Revision", "NO_OF_ERRORS", "ERROR_CATEGORY", "REMARKS"] });
    const workbook = XLSX.utils.book_new();

    // Append the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbook, worksheet, "Comments");

    // Create the Excel file and trigger a download
    XLSX.writeFile(workbook, "comments.xlsx");
}
