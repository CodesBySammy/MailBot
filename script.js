class MailBot {
    constructor() {
        this.workbook = null;
        this.initEventListeners();
    }

    initEventListeners() {
        document.getElementById('excelFile').addEventListener('change', this.handleFileUpload.bind(this));
        document.getElementById('sheetSelect').addEventListener('change', this.loadSheetPreview.bind(this));
        document.getElementById('sendEmails').addEventListener('click', this.sendEmails.bind(this));
    }

    handleFileUpload(e) {
        const file = e.target.files[0];
        const reader = new FileReader();

        reader.onload = (event) => {
            try {
                const data = new Uint8Array(event.target.result);
                this.workbook = XLSX.read(data, {type: 'array'});

                const sheetSelect = document.getElementById('sheetSelect');
                const sheetSelectContainer = document.getElementById('sheetSelectContainer');
                
                sheetSelect.innerHTML = '<option>Select Sheet</option>';
                this.workbook.SheetNames.forEach(sheetName => {
                    const option = document.createElement('option');
                    option.value = sheetName;
                    option.textContent = sheetName;
                    sheetSelect.appendChild(option);
                });
                
                sheetSelectContainer.style.display = 'block';
            } catch (error) {
                this.showError('Error processing Excel file');
                console.error(error);
            }
        };

        reader.readAsArrayBuffer(file);
    }

    loadSheetPreview() {
        const sheetName = document.getElementById('sheetSelect').value;
        if (!sheetName) return;

        try {
            const worksheet = this.workbook.Sheets[sheetName];
            const data = XLSX.utils.sheet_to_json(worksheet, { 
                defval: '', 
                header: 1 // Use the first row as headers
            });

            const headers = data[0];
            const previewBody = document.getElementById('previewBody');
            previewBody.innerHTML = '';

            // Update table headers dynamically
            const previewHeader = document.querySelector('#excelPreview thead tr');
            previewHeader.innerHTML = headers.map(header => 
                `<th class="border p-2">${header}</th>`
            ).join('');

            // Skip the header row and render data rows
            data.slice(1).forEach((row) => {
                const tr = document.createElement('tr');
                headers.forEach(header => {
                    const td = document.createElement('td');
                    // Find the index of the current header in the headers array
                    const columnIndex = headers.indexOf(header);
                    td.textContent = row[columnIndex] || 'N/A';
                    td.className = 'border p-2';
                    tr.appendChild(td);
                });
                previewBody.appendChild(tr);
            });
        } catch (error) {
            this.showError('Error loading sheet preview');
            console.error(error);
        }
    }

    async sendEmails() {
        const fromEmail = document.getElementById('fromEmail').value;
        const appPassword = document.getElementById('appPassword').value;
        const subject = document.getElementById('emailSubject').value;
        const message = document.getElementById('fixedMessage').value;
        const htmlMessage = document.getElementById('htmlMessage').value;

        const sheetName = document.getElementById('sheetSelect').value;
        const worksheet = this.workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet, { 
            defval: '', 
            header: 1 // Use the first row as headers
        });

        const headers = data[0];
        const emailData = data.slice(1).map(row => {
            // Convert row to an object with headers as keys
            const rowObj = {};
            headers.forEach((header, index) => {
                rowObj[header] = row[index] || '';
            });
            return rowObj;
        });

        const progressStatus = document.getElementById('progressStatus');
        const sendButton = document.getElementById('sendEmails');

        sendButton.disabled = true;
        progressStatus.innerHTML = 'Preparing to send emails...';
        progressStatus.className = 'mt-4 text-center p-3 rounded bg-yellow-100';

        try {
            const response = await fetch('https://mail-bot-be-pink.vercel.app/send-emails', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    fromEmail,
                    appPassword,
                    subject,
                    message: message || undefined,
                    html: htmlMessage || undefined,
                    emails: emailData
                })
            });

            if (!response.ok) {
                throw new Error('Network response was not ok');
            }
            const result = await response.json();

            if (result.summary) {
                progressStatus.innerHTML = `
                    Email Sending Complete:
                    Total Emails: ${result.summary.total}
                    Successful: ${result.summary.success}
                    Failed: ${result.summary.failed}
                `;
                progressStatus.className = 'mt-4 text-center p-3 rounded bg-green-100';
            } else {
                progressStatus.innerHTML = result.message || 'Unknown result';
                progressStatus.className = 'mt-4 text-center p-3 rounded bg-red-100';
            }
        } catch (error) {
            progressStatus.innerHTML = `Error: ${error.message}`;
            progressStatus.className = 'mt-4 text-center p-3 rounded bg-red-100';
            console.error('Email sending error:', error);
        } finally {
            sendButton.disabled = false;
        }
    }

    showError(message) {
        const progressStatus = document.getElementById('progressStatus');
        progressStatus.innerHTML = message;
        progressStatus.className = 'mt-4 text-center p-3 rounded bg-red-100';
    }
}

document.addEventListener('DOMContentLoaded', () => {
    new MailBot();
});
