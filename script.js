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

    handleFileUpload(event) {
        const file = event.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                this.workbook = XLSX.read(data, { type: 'array' });

                const sheetSelect = document.getElementById('sheetSelect');
                sheetSelect.innerHTML = '<option value="">Select Sheet</option>';
                this.workbook.SheetNames.forEach(sheetName => {
                    const option = document.createElement('option');
                    option.value = sheetName;
                    option.textContent = sheetName;
                    sheetSelect.appendChild(option);
                });

                document.getElementById('sheetSelectContainer').style.display = 'block';
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
            const data = XLSX.utils.sheet_to_json(worksheet, { defval: '', header: 1 });

            if (data.length === 0) {
                this.showError('Empty sheet. Please select a valid sheet.');
                return;
            }

            const headers = data[0];
            const previewBody = document.getElementById('previewBody');
            const previewHeader = document.querySelector('#excelPreview thead tr');

            previewHeader.innerHTML = headers.map(header => `<th class="border p-2">${header}</th>`).join('');
            previewBody.innerHTML = '';

            data.slice(1).forEach(row => {
                const tr = document.createElement('tr');
                headers.forEach((header, index) => {
                    const td = document.createElement('td');
                    td.textContent = row[index] || 'N/A';
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
        const fromEmail = document.getElementById('fromEmail').value.trim();
        const appPassword = document.getElementById('appPassword').value.trim();
        const subject = document.getElementById('emailSubject').value.trim();
        const message = document.getElementById('fixedMessage').value.trim();
        const htmlMessage = document.getElementById('htmlMessage').value.trim();
        const fromName = document.getElementById('fromName').value.trim();
        const sheetName = document.getElementById('sheetSelect').value;

        if (!fromEmail || !appPassword || !subject || !sheetName) {
            this.showError('Please fill all required fields and select a sheet.');
            return;
        }

        const worksheet = this.workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet, { defval: '', header: 1 });

        if (data.length < 2) {
            this.showError('No recipient data found in the selected sheet.');
            return;
        }

        const headers = data[0];
        const emailData = data.slice(1).map(row => {
            let rowObj = {};
            headers.forEach((header, index) => {
                rowObj[header] = row[index] || '';
            });
            return rowObj;
        });

        const progressStatus = document.getElementById('progressStatus');
        const sendButton = document.getElementById('sendEmails');

        sendButton.disabled = true;
        progressStatus.innerHTML = 'Sending emails...';
        progressStatus.className = 'mt-4 text-center p-3 rounded bg-yellow-100';

        try {
            const response = await fetch('http://localhost:3000/send-emails', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    fromEmail,
                    appPassword,
                    subject,
                    message: message || undefined,
                    html: htmlMessage || undefined,
                    emails: emailData,
                    fromName
                })
            });

            if (!response.ok) throw new Error('Network response was not ok');
            const result = await response.json();

            progressStatus.innerHTML = `
                Email Sending Complete:<br>
                Total Emails: ${result.summary.total}<br>
                Successful: ${result.summary.success}<br>
                Failed: ${result.summary.failed}
            `;
            progressStatus.className = 'mt-4 text-center p-3 rounded bg-green-100';
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

document.addEventListener('DOMContentLoaded', () => new MailBot());
