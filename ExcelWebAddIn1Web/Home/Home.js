// Set the worker source for PDF.js
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.10.377/pdf.worker.min.js';

let currentPDF = null;
let selectionBox = null;
let isSelecting = false;

Office.onReady(() => {
    $(document).ready(() => {
        console.log('Office.js and jQuery initialized');
        $('#pdf-file-input').on('change', handleFileSelect);
        
        // Add event listeners for selection
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

// Handle window resize
$(window).resize(() => {
    if (currentPDF) {
        renderAllPages();
    }
});

function handleMouseDown(e) {
    const container = document.getElementById('pdf-viewer-container');
    const rect = container.getBoundingClientRect();
    
    // Only start selection if click is inside the PDF viewer
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

    // Clean up
    document.body.removeChild(selectionBox);
    selectionBox = null;
}

function captureScreenshot(rect) {
    const container = document.getElementById('pdf-viewer-container');
    const containerRect = container.getBoundingClientRect();

    // Calculate relative position to container
    const x = rect.left - containerRect.left;
    const y = rect.top - containerRect.top;

    html2canvas(container, {
        x: x,
        y: y,
        width: rect.width,
        height: rect.height,
        backgroundColor: null,
        useCORS: true,
        scale: window.devicePixelRatio
    }).then(canvas => {
        // Save to file
        canvas.toBlob(blob => {
            // Save to file
            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            link.download = 'snippet.png';
            link.click();

            // Copy to clipboard
            if (navigator.clipboard && navigator.clipboard.write) {
                const item = new ClipboardItem({ 'image/png': blob });
                navigator.clipboard.write([item]).catch(err => {
                    console.error('Error copying to clipboard:', err);
                });
            }
        });
    }).catch(error => {
        console.error('Error capturing screenshot:', error);
    });
}
