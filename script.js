
const dragArea = document.getElementById('dragArea');
const fileInput = document.getElementById('excelFile');
const fileList = document.getElementById('fileList');

['dragenter', 'dragover'].forEach(eventType => {
    dragArea.addEventListener(eventType, (e) => {
        e.preventDefault();
        dragArea.classList.add('dragging');
    });
});

['dragleave', 'drop'].forEach(eventType => {
    dragArea.addEventListener(eventType, (e) => {
        e.preventDefault();
        dragArea.classList.remove('dragging');
    });
});

dragArea.addEventListener('drop', (e) => {
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        handleFiles(files);
    }
});

dragArea.addEventListener('click', () => fileInput.click());

fileInput.addEventListener('change', () => {
    if (fileInput.files.length > 0) {
        handleFiles(fileInput.files);
    }
});

// Function to handle multiple files
function handleFiles(files) {
    for (let file of files) {
        processFile(file);
    }
}

function processFile(file) {
    const listItem = document.createElement('li');
    listItem.innerHTML = `
      <span>${file.name}</span>
      <div class="file-status">
        <div class="loading-spinner"></div>
        <span class="status-text processing">Elaborazione...</span>
      </div>
    `;
    fileList.appendChild(listItem);

    // Read the file as binary
    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        // Get the first worksheet
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];


        const columns = [0, 2, 5, 3, 4];
        const contentColumn = [];
        const range = XLSX.utils.decode_range(worksheet['!ref']);


        columns.forEach(() => contentColumn.push([]));



        columns.forEach(col => {
            for (let row = range.s.r + 1; row <= range.e.r; row++) {
                const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                const cell = worksheet[cellAddress];
                let isToPushValue = true;
                if (cell) {
                    let cellExtracted = cell.v;
                    if (cellExtracted.length > 0) {
                        cellExtracted = cellExtracted.trim();
                        if (col === 0) {
                            cellExtracted = cellExtracted.split(' ').slice(0, -1).join(' ').trim();
                        }
                        else if (col === 2) {
                            cellExtracted = cellExtracted.slice(2);
                            if (cellExtracted.length > 0) {

                                let indexFind = contentColumn[0].findIndex(element => {
                                    const parts = element.split(' - ');
                                    return parts.length > 1 && cellExtracted.slice(0).trim().includes(parts[1].trim());
                                }) + 1;

                                if (indexFind > 0 && contentColumn[0][row] === '') {
                                    isToPushValue = false;
                                    contentColumn[0][row] = '@@@';
                                }

                                cellExtracted = cellExtracted.charAt(0).toUpperCase() + cellExtracted.slice(1);

                            }

                        }
                    }
                    if (isToPushValue)
                        contentColumn[columns.indexOf(col)].push(cellExtracted);
                }
            }
        });

        const indexToDelete = contentColumn[0].reduce((acc, item, index) => {
            if (item === '@@@') {
                acc.push(index);
            }
            return acc;
        }, []);

        for (let index of indexToDelete) {
            for (let col of [3, 4]) {
                const columnIndex = columns.indexOf(col);
                contentColumn[columnIndex][index] = '@@@'
            }
        }

        for (let col of [0, 3, 4]) {
            const columnIndex = columns.indexOf(col);
            contentColumn[columnIndex]=  contentColumn[columnIndex].filter(elem => elem!=='@@@');
        }

        
        // Ensure all columns have the same length by filling with empty strings
        const maxRows = Math.max(...contentColumn.map(col => col.length));
        const finalContent = [];

        for (let i = 0; i < maxRows; i++) {
            const row = columns.map((col, colIndex) => contentColumn[colIndex][i] || "");
            finalContent.push(row);
        }
        
        // Create a new worksheet with the extracted columns
        const newWorksheet = XLSX.utils.aoa_to_sheet(finalContent);
        const newWorkbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'ProcessedData');

        // Generate a new Excel file
        const wbout = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'array' });

        // Create a download link for the processed file
        const blob = new Blob([wbout], { type: 'application/octet-stream' });
        const downloadLink = document.createElement('a');
        const url = window.URL.createObjectURL(blob);
        downloadLink.href = url;
        downloadLink.download = file.name.replace(/\.[^/.]+$/, "") + '_processed.xlsx';
        downloadLink.classList.add('btn', 'btn-success', 'btn-sm');
        downloadLink.textContent = 'Scarica';

        // Update file status to "Complete" and append download link
        listItem.querySelector('.file-status').innerHTML = `<span class="status-text complete">Elaborato</span>`;
        listItem.appendChild(downloadLink);

        // Revoke the object URL after download to free memory
        downloadLink.addEventListener('click', () => {
            setTimeout(() => window.URL.revokeObjectURL(url), 100);
        });
    };

    // Read the file as binary string
    reader.readAsArrayBuffer(file);
}



