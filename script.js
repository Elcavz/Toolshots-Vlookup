let file1RawData = []; // Store raw values
let file1CleanData = []; // Store cleaned values
let file2Data = {};

// Handle file 1 upload
document.getElementById('file1Input').addEventListener('change', function (e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        processFile1(data);
    };
    reader.readAsArrayBuffer(file);
});

// Handle file 2 upload
document.getElementById('file2Input').addEventListener('change', function (e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        processFile2(data);
    };
    reader.readAsArrayBuffer(file);
});

// Process button click
document.getElementById('processBtn').addEventListener('click', performVlookup);

function cleanString(input) {
    if (!input || typeof input !== 'string') return null;

    // Remove everything from first A-Z letter onward (case insensitive)
    const cleaned = input.replace(/[A-Za-z].*$/i, '').trim();

    return cleaned || null;
}

function processFile1(data) {
    try {
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        // Reset arrays
        file1RawData = [];
        file1CleanData = [];

        jsonData.forEach(row => {
            const rawValue = row[0];
            if (rawValue !== undefined) {
                const cleanValue = cleanString(rawValue);
                if (cleanValue !== null) {
                    file1RawData.push(rawValue);
                    file1CleanData.push(cleanValue);
                }
            }
        });

        updateProcessButton();
    } catch (error) {
        console.error('Error processing File 1:', error);
        alert('Error processing first file. Please make sure it\'s a valid Excel file.');
    }
}

function processFile2(data) {
    try {
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        // Create lookup dictionary with cleaned keys
        file2Data = {};

        jsonData.forEach(row => {
            const key = cleanString(row[0]);
            if (key && row[1] !== undefined) {
                file2Data[key] = row[1];
            }
        });

        updateProcessButton();
    } catch (error) {
        console.error('Error processing File 2:', error);
        alert('Error processing second file. Please make sure it\'s a valid Excel file.');
    }
}

function updateProcessButton() {
    const btn = document.getElementById('processBtn');
    btn.disabled = !(file1CleanData.length > 0 && Object.keys(file2Data).length > 0);
}

function performVlookup() {
    const resultBody = document.getElementById('resultBody');
    resultBody.innerHTML = '';
    
    if (file1CleanData.length === 0 || Object.keys(file2Data).length === 0) {
        resultBody.innerHTML = '<tr><td colspan="3">Please upload both files first</td></tr>';
        return;
    }
    
    file1CleanData.forEach((cleanValue, index) => {
        const row = document.createElement('tr');
        
        // Raw value cell
        const rawCell = document.createElement('td');
        rawCell.textContent = file1RawData[index];
        row.appendChild(rawCell);
        
        // Cleaned value cell
        const cleanCell = document.createElement('td');
        cleanCell.textContent = cleanValue;
        row.appendChild(cleanCell);
        
        // Result cell (VLOOKUP)
        const resultCell = document.createElement('td');
        resultCell.textContent = file2Data[cleanValue] || ''; // Blank if not found
        row.appendChild(resultCell);
        
        resultBody.appendChild(row);
    });
}