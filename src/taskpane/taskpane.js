Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.addEventListener("DOMContentLoaded", function() {
            initializeAddIn();
        });
    }
});

let currentEmailData = null;

function initializeAddIn() {
    // Set up event listeners
    setupEventListeners();
    
    // Load email data
    loadEmailData();
    
    // Update filename preview initially
    updateFilenamePreview();
}

function setupEventListeners() {
    // Checkbox and input change listeners for preview updates
    document.getElementById('includeDate').addEventListener('change', updateFilenamePreview);
    document.getElementById('dateFormat').addEventListener('change', updateFilenamePreview);
    document.getElementById('includeSender').addEventListener('change', updateFilenamePreview);
    document.getElementById('includeSubject').addEventListener('change', updateFilenamePreview);
    document.getElementById('subjectLength').addEventListener('input', updateFilenamePreview);
    document.getElementById('separator').addEventListener('change', updateFilenamePreview);
    
    // Convert button
    document.getElementById('convertToPDF').addEventListener('click', convertToPDF);
}

function loadEmailData() {
    Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            // Get email properties
            const item = Office.context.mailbox.item;
            
            currentEmailData = {
                subject: item.subject || 'No Subject',
                sender: item.sender ? item.sender.displayName || item.sender.emailAddress : 'Unknown Sender',
                dateTimeReceived: item.dateTimeCreated || new Date(),
                body: null // Will be loaded when converting
            };
            
            updateFilenamePreview();
        } else {
            showStatus('Error loading email data: ' + result.error.message, 'error');
        }
    });
}

function updateFilenamePreview() {
    if (!currentEmailData) {
        document.getElementById('filenamePreview').textContent = 'Loading...';
        return;
    }
    
    const includeDate = document.getElementById('includeDate').checked;
    const dateFormat = document.getElementById('dateFormat').value;
    const includeSender = document.getElementById('includeSender').checked;
    const includeSubject = document.getElementById('includeSubject').checked;
    const subjectLength = parseInt(document.getElementById('subjectLength').value) || 50;
    const separator = document.getElementById('separator').value;
    
    const parts = [];
    
    // Add date if selected
    if (includeDate) {
        const date = new Date(currentEmailData.dateTimeReceived);
        let formattedDate = '';
        
        switch (dateFormat) {
            case 'YYYY-MM-DD':
                formattedDate = date.getFullYear() + '-' + 
                               String(date.getMonth() + 1).padStart(2, '0') + '-' + 
                               String(date.getDate()).padStart(2, '0');
                break;
            case 'DD-MM-YYYY':
                formattedDate = String(date.getDate()).padStart(2, '0') + '-' +
                               String(date.getMonth() + 1).padStart(2, '0') + '-' +
                               date.getFullYear();
                break;
            case 'MM-DD-YYYY':
                formattedDate = String(date.getMonth() + 1).padStart(2, '0') + '-' +
                               String(date.getDate()).padStart(2, '0') + '-' +
                               date.getFullYear();
                break;
        }
        parts.push(formattedDate);
    }
    
    // Add sender if selected
    if (includeSender) {
        const cleanSender = sanitizeFileName(currentEmailData.sender);
        parts.push(cleanSender);
    }
    
    // Add subject if selected
    if (includeSubject) {
        let subject = currentEmailData.subject;
        if (subject.length > subjectLength) {
            subject = subject.substring(0, subjectLength) + '...';
        }
        const cleanSubject = sanitizeFileName(subject);
        parts.push(cleanSubject);
    }
    
    const filename = parts.join(separator) + '.pdf';
    document.getElementById('filenamePreview').textContent = filename || 'email.pdf';
}

function sanitizeFileName(fileName) {
    // Remove or replace invalid characters for file names
    return fileName
        .replace(/[<>:"/\\|?*]/g, '') // Remove invalid characters
        .replace(/\s+/g, ' ') // Replace multiple spaces with single space
        .trim();
}

async function convertToPDF() {
    try {
        showProgress(true);
        showStatus('', '');
        
        // Disable the convert button
        document.getElementById('convertToPDF').disabled = true;
        
        // Get the email body
        const emailBody = await getEmailBody();
        
        // Create PDF content
        await createPDF(emailBody);
        
        // Add category to email
        await addEmailCategory();
        
        showStatus('Email successfully converted to PDF and categorized!', 'success');
        
    } catch (error) {
        console.error('Error converting to PDF:', error);
        showStatus('Error converting to PDF: ' + error.message, 'error');
    } finally {
        showProgress(false);
        document.getElementById('convertToPDF').disabled = false;
    }
}

function getEmailBody() {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.body.getAsync(
            Office.CoercionType.Html,
            (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    resolve(result.value);
                } else {
                    reject(new Error('Failed to get email body: ' + result.error.message));
                }
            }
        );
    });
}

function loadJsPDF() {
    return new Promise((resolve, reject) => {
        if (window.jsPDF) {
            resolve();
            return;
        }
        
        const script = document.createElement('script');
        script.src = 'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js';
        script.onload = () => {
            resolve();
        };
        script.onerror = (error) => {
            reject(new Error('Failed to load jsPDF library'));
        };
        document.head.appendChild(script);
    });
}

async function createPDF(emailBody) {
    return new Promise(async (resolve, reject) => {
        try {
            // Load jsPDF from CDN if not already loaded
            if (!window.jsPDF) {
                await loadJsPDF();
            }
            
            const doc = new window.jsPDF.jsPDF();
            
            // Set up PDF metadata
            const filename = document.getElementById('filenamePreview').textContent;
            doc.setProperties({
                title: filename,
                subject: currentEmailData.subject,
                author: currentEmailData.sender,
                creator: 'Outlook PDF Converter'
            });
            
            // Add email header information
            doc.setFontSize(16);
            doc.text('Email Details', 20, 20);
            
            doc.setFontSize(12);
            let yPosition = 35;
            
            doc.text(`From: ${currentEmailData.sender}`, 20, yPosition);
            yPosition += 10;
            
            doc.text(`Subject: ${currentEmailData.subject}`, 20, yPosition);
            yPosition += 10;
            
            const dateStr = new Date(currentEmailData.dateTimeReceived).toLocaleString();
            doc.text(`Date: ${dateStr}`, 20, yPosition);
            yPosition += 20;
            
            // Add email body
            doc.setFontSize(10);
            
            // Convert HTML to plain text for PDF (basic conversion)
            const tempDiv = document.createElement('div');
            tempDiv.innerHTML = emailBody;
            const plainText = tempDiv.textContent || tempDiv.innerText || '';
            
            // Split text into lines that fit the PDF width
            const maxLineWidth = 170;
            const lines = doc.splitTextToSize(plainText, maxLineWidth);
            
            // Add lines to PDF, handling page breaks
            const pageHeight = doc.internal.pageSize.height;
            const lineHeight = 5;
            
            for (let i = 0; i < lines.length; i++) {
                if (yPosition + lineHeight > pageHeight - 20) {
                    doc.addPage();
                    yPosition = 20;
                }
                doc.text(lines[i], 20, yPosition);
                yPosition += lineHeight;
            }
            
            // Save the PDF
            doc.save(filename);
            resolve();
            
        } catch (error) {
            reject(error);
        }
    });
}

function addEmailCategory() {
    return new Promise((resolve, reject) => {
        // Get current categories
        Office.context.mailbox.item.categories.getAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const currentCategories = result.value;
                
                // Check if PDF category already exists
                const pdfCategoryExists = currentCategories.some(cat => cat.displayName === 'PDF');
                
                if (!pdfCategoryExists) {
                    // Add the PDF category
                    const pdfCategory = {
                        displayName: 'PDF',
                        color: Office.MailboxEnums.CategoryColor.Preset1 // Red color
                    };
                    
                    const updatedCategories = [...currentCategories, pdfCategory];
                    
                    Office.context.mailbox.item.categories.setAsync(
                        updatedCategories,
                        (setResult) => {
                            if (setResult.status === Office.AsyncResultStatus.Succeeded) {
                                resolve();
                            } else {
                                reject(new Error('Failed to add PDF category: ' + setResult.error.message));
                            }
                        }
                    );
                } else {
                    resolve(); // Category already exists
                }
            } else {
                reject(new Error('Failed to get current categories: ' + result.error.message));
            }
        });
    });
}

function showProgress(show) {
    const progressIndicator = document.getElementById('progressIndicator');
    if (show) {
        progressIndicator.classList.remove('hidden');
    } else {
        progressIndicator.classList.add('hidden');
    }
}

function showStatus(message, type) {
    const statusSection = document.getElementById('statusSection');
    const statusMessage = document.getElementById('statusMessage');
    
    if (message) {
        statusMessage.textContent = message;
        statusSection.className = 'status';
        if (type) {
            statusSection.classList.add(type);
        }
        statusSection.classList.remove('hidden');
    } else {
        statusSection.classList.add('hidden');
    }
}