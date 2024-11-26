// Set the worker source for PDF.js
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.10.377/pdf.worker.min.js';

let currentPDF = null;
let selectionBox = null;
let isSelecting = false;
let notificationTimeout = null;

Office.onReady(() => {
    $(document).ready(() => {
        console.log('Office.js and jQuery initialized');
        $('#pdf-file-input').on('change', handleFileSelect);
        
        // Add mouse event listeners for selection
        const container = document.getElementById('pdf-viewer-container');
        container.addEventListener('mousedown', handleMouseDown);
        document.addEventListener('mousemove', handleMouseMove);
        document.addEventListener('mouseup', handleMouseUp);
    });
});

function handleFileSelect(event) {
    const file = event.target.files[0];
    if (file && file.type === 'application/pdf') {
        const fileReader = new FileReader();
        fileReader.onload = function() {
            const typedarray = new Uint8Array(this.result);

            pdfjsLib.getDocument(typedarray).promise.then(pdf => {
                currentPDF = pdf;
                $('#pdf-viewer-container').empty(); // Clear previous content
                renderAllPages();
            }).catch(error => {
                console.error('Error loading PDF:', error);
            });
        };
        fileReader.readAsArrayBuffer(file);
    } else {
        console.error('Please select a valid PDF file.');
    }
}

async function renderAllPages() {
    if (!currentPDF) return;

    const container = document.getElementById('pdf-viewer-container');
    container.style.display = 'block';

    // Show loading message
    container.innerHTML = '<div class="loading">Loading PDF...</div>';

    try {
        for (let pageNum = 1; pageNum <= currentPDF.numPages; pageNum++) {
            const page = await currentPDF.getPage(pageNum);
            const scale = 1.5;
            const viewport = page.getViewport({ scale });

            // Create canvas for this page
            const canvas = document.createElement('canvas');
            canvas.className = 'pdf-page';
            const context = canvas.getContext('2d');
            canvas.height = viewport.height;
            canvas.width = viewport.width;

            // Add canvas to container
            container.appendChild(canvas);

            // Render page
            await page.render({
                canvasContext: context,
                viewport: viewport
            }).promise;
        }
    } catch (error) {
        console.error('Error rendering PDF:', error);
        container.innerHTML = `<div class="error">Error rendering PDF: ${error.message}</div>`;
    }
}

function handleMouseDown(e) {
    const container = document.getElementById('pdf-viewer-container');
    const rect = container.getBoundingClientRect();
    
    if (e.clientX >= rect.left && e.clientX <= rect.right &&
        e.clientY >= rect.top && e.clientY <= rect.bottom) {
        
        if (selectionBox) {
            document.body.removeChild(selectionBox);
        }

        selectionBox = document.createElement('div');
        selectionBox.style.position = 'absolute';
        selectionBox.style.border = '2px dashed #0078d4';
        selectionBox.style.backgroundColor = 'rgba(0, 120, 212, 0.1)';
        selectionBox.style.zIndex = '1000';
        document.body.appendChild(selectionBox);

        isSelecting = true;
        selectionBox.style.left = `${e.pageX}px`;
        selectionBox.style.top = `${e.pageY}px`;
    }
}

function handleMouseMove(e) {
    if (!isSelecting || !selectionBox) return;

    const startX = parseInt(selectionBox.style.left);
    const startY = parseInt(selectionBox.style.top);
    
    const width = Math.abs(e.pageX - startX);
    const height = Math.abs(e.pageY - startY);

    selectionBox.style.width = `${width}px`;
    selectionBox.style.height = `${height}px`;
    selectionBox.style.left = `${Math.min(e.pageX, startX)}px`;
    selectionBox.style.top = `${Math.min(e.pageY, startY)}px`;
}

function handleMouseUp(e) {
    if (!isSelecting || !selectionBox) return;
    isSelecting = false;

    const width = parseInt(selectionBox.style.width);
    const height = parseInt(selectionBox.style.height);

    if (width > 10 && height > 10) {
        captureScreenshot(selectionBox.getBoundingClientRect());
    }

    document.body.removeChild(selectionBox);
    selectionBox = null;
}

async function captureScreenshot(rect) {
    const container = document.getElementById('pdf-viewer-container');
    const containerRect = container.getBoundingClientRect();

    try {
        // Calculate relative position
        const x = rect.left - containerRect.left;
        const y = rect.top - containerRect.top;

        // First capture the area
        const canvas = await html2canvas(container, {
            x: x,
            y: y,
            width: rect.width,
            height: rect.height,
            backgroundColor: null,
            useCORS: true,
            scale: window.devicePixelRatio * 2 // Increase resolution
        });

        // Enhanced preprocessing (similar to Python version)
        const ctx = canvas.getContext('2d');
        let imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
        
        // Convert to grayscale and enhance contrast
        const data = imageData.data;
        for (let i = 0; i < data.length; i += 4) {
            // Convert to grayscale
            const gray = (data[i] + data[i + 1] + data[i + 2]) / 3;
            
            // Enhance contrast (similar to Python's enhance(3.0))
            const contrast = 3.0;
            let enhanced = 128 + contrast * (gray - 128);
            
            // Binary threshold (similar to Python's cv2.threshold)
            const threshold = 150;
            const binary = enhanced > threshold ? 255 : 0;
            
            data[i] = binary;     // R
            data[i + 1] = binary; // G
            data[i + 2] = binary; // B
        }
        ctx.putImageData(imageData, 0, 0);

        // Perform OCR with better configuration
        const worker = await Tesseract.createWorker();
        await worker.loadLanguage('eng');
        await worker.initialize('eng');
        await worker.setParameters({
            tessedit_char_whitelist: '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz.,()-:/ ',
            preserve_interword_spaces: '1',
        });
        
        const { data: { text } } = await worker.recognize(canvas);
        await worker.terminate();

        // Format text for Excel (similar to Python version)
        const lines = text.split('\n').filter(line => line.trim() !== '');
        const formattedText = lines.join('\t');

        // Copy to clipboard
        await navigator.clipboard.writeText(formattedText);
        showNotification('✓ Copied to clipboard');

        // Save image for verification
        canvas.toBlob(blob => {
            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            link.download = 'snippet.png';
            link.click();
            showNotification('✓ Image saved');
        });

        // Try to insert into Excel
        try {
            await Excel.run(async (context) => {
                const range = context.workbook.getSelectedRange();
                const values = lines.map(line => line.split(/\s+/));
                range.values = values;
                await context.sync();
                showNotification('✓ Inserted into Excel');
            });
        } catch (excelError) {
            console.error('Excel insertion failed:', excelError);
            // At least we have it in clipboard
        }

    } catch (error) {
        console.error('Error processing image:', error);
        showNotification('❌ Error: ' + error.message, 'error');
    }
}

// Add this helper function for better image processing
function enhanceImageData(imageData) {
    const data = imageData.data;
    const contrast = 3.0;
    const threshold = 150;

    for (let i = 0; i < data.length; i += 4) {
        const gray = (data[i] + data[i + 1] + data[i + 2]) / 3;
        const enhanced = 128 + contrast * (gray - 128);
        const binary = enhanced > threshold ? 255 : 0;
        
        data[i] = binary;
        data[i + 1] = binary;
        data[i + 2] = binary;
        data[i + 3] = 255; // Alpha
    }
    return imageData;
}

// Handle window resize
$(window).resize(() => {
    if (currentPDF) {
        renderAllPages();
    }
});

// Add this function for notifications
function showNotification(message, type = 'success') {
    // Clear any existing notification
    if (notificationTimeout) {
        clearTimeout(notificationTimeout);
        const existingNotification = document.querySelector('.notification');
        if (existingNotification) {
            existingNotification.remove();
        }
    }

    // Create notification element
    const notification = document.createElement('div');
    notification.className = `notification ${type}`;
    notification.innerHTML = `
        <div class="notification-content">
            <span class="notification-message">${message}</span>
            <div class="notification-progress"></div>
        </div>
    `;

    // Add to document
    document.body.appendChild(notification);

    // Animate in
    setTimeout(() => notification.classList.add('show'), 10);

    // Remove after 3 seconds
    notificationTimeout = setTimeout(() => {
        notification.classList.remove('show');
        setTimeout(() => notification.remove(), 300);
    }, 3000);
}
