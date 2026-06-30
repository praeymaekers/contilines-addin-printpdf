/*
 * Print als PDF - Conti-Lines NV
 * Office JS Add-in: converteert Word naar PDF en opent printdialoog
 */

Office.onReady(() => {
    // Office JS is geladen
});

/**
 * Hoofdfunctie - gekoppeld aan de ribbon knop
 */
function printAsPdf(event) {
    getDocumentAsPdf()
        .then(function (pdfBlob) {
            var url = URL.createObjectURL(pdfBlob);
            openPrintWindow(url, event);
        })
        .catch(function (err) {
            // Toon foutmelding via Office notificatie
            try {
                Office.context.document.settings.set('lastError', err.message);
            } catch (e) {}
            // Laat Word weten dat de functie klaar is (ook bij fout)
            event.completed({ allowEvent: false });
            alert('Fout bij aanmaken PDF: ' + err.message);
        });
}

/**
 * Haalt het volledige document op als PDF (in slices van 64KB)
 */
function getDocumentAsPdf() {
    return new Promise(function (resolve, reject) {
        Office.context.document.getFileAsync(
            Office.FileType.Pdf,
            { sliceSize: 65536 },
            function (result) {
                if (result.status !== Office.AsyncResultStatus.Succeeded) {
                    reject(new Error(result.error.message));
                    return;
                }

                var file = result.value;
                var total = file.sliceCount;
                var slices = new Array(total);
                var received = 0;

                for (var i = 0; i < total; i++) {
                    // Slices worden parallel opgehaald
                    (function (index) {
                        file.getSliceAsync(index, function (sliceResult) {
                            if (sliceResult.status !== Office.AsyncResultStatus.Succeeded) {
                                file.closeAsync();
                                reject(new Error(sliceResult.error.message));
                                return;
                            }

                            slices[sliceResult.value.index] = new Uint8Array(sliceResult.value.data);
                            received++;

                            if (received === total) {
                                file.closeAsync();

                                // Alle slices samenvoegen tot één PDF Blob
                                var totalBytes = slices.reduce(function (n, s) { return n + s.length; }, 0);
                                var combined = new Uint8Array(totalBytes);
                                var offset = 0;
                                for (var j = 0; j < slices.length; j++) {
                                    combined.set(slices[j], offset);
                                    offset += slices[j].length;
                                }

                                resolve(new Blob([combined], { type: 'application/pdf' }));
                            }
                        });
                    })(i);
                }
            }
        );
    });
}

/**
 * Opent de PDF in een nieuw venster en stuurt meteen het print-commando
 */
function openPrintWindow(pdfUrl, event) {
    var printWindow = window.open(pdfUrl, '_blank');

    if (!printWindow) {
        // Popup geblokkeerd: fallback naar downloaden
        downloadFallback(pdfUrl);
        event.completed();
        return;
    }

    var printed = false;

    function doPrint() {
        if (printed) return;
        printed = true;
        try {
            printWindow.focus();
            printWindow.print();
        } catch (e) {
            // Sommige PDF-viewers staan geen window.print() toe vanuit extern script
            // Gebruiker kan zelf Ctrl+P gebruiken - venster is wel geopend
        }
        // Geheugenbeheer: URL vrijgeven na 5 minuten
        setTimeout(function () { URL.revokeObjectURL(pdfUrl); }, 300000);
        event.completed();
    }

    // Wacht op load event, met timeout als fallback
    printWindow.addEventListener('load', function () {
        setTimeout(doPrint, 800);
    });

    // Fallback als load event niet afvuurt (bijv. bij Edge PDF viewer)
    setTimeout(doPrint, 4000);
}

/**
 * Fallback: download de PDF als het venster geblokkeerd wordt
 */
function downloadFallback(pdfUrl) {
    var docUrl = '';
    try { docUrl = Office.context.document.url || ''; } catch (e) {}
    var name = docUrl
        ? decodeURIComponent(docUrl.split('/').pop().replace(/\.[^.]+$/, '')) + '.pdf'
        : 'document.pdf';

    var a = document.createElement('a');
    a.href = pdfUrl;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(pdfUrl); }, 60000);
}

// Functie registreren bij Office
if (typeof Office !== 'undefined') {
    Office.actions.associate('printAsPdf', printAsPdf);
}
