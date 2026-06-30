/*
 * Print als PDF - Conti-Lines NV
 * Office JS Add-in: converteert Word naar PDF en opent printdialoog
 */

Office.onReady(() => {});

/**
 * Hoofdfunctie - gekoppeld aan de ribbon knop
 */
function printAsPdf(event) {
    getDocumentAsPdf()
        .then(function (pdfBlob) {
            var url = URL.createObjectURL(pdfBlob);
            printViaIframe(url, event);
        })
        .catch(function (err) {
            event.completed({ allowEvent: false });
            alert('Fout bij aanmaken PDF: ' + err.message);
        });
}

/**
 * Laadt de PDF in een verborgen iframe en roept window.print() op
 * (vermijdt het blob:// protocol probleem met window.open in WebView2)
 */
function printViaIframe(pdfUrl, event) {
    var iframe = document.createElement('iframe');
    iframe.style.cssText = 'width:1px;height:1px;position:absolute;top:-9999px;left:-9999px;border:none;';
    iframe.src = pdfUrl;
    document.body.appendChild(iframe);

    var done = false;

    function tryPrint() {
        if (done) return;
        done = true;
        try {
            iframe.contentWindow.focus();
            iframe.contentWindow.print();
        } catch (e) {
            // WebView2 staat soms geen cross-frame print toe - fallback naar parent
            try { window.print(); } catch (e2) {}
        }
        event.completed();
        // Opruimen na 5 minuten
        setTimeout(function () {
            try { document.body.removeChild(iframe); } catch (e) {}
            URL.revokeObjectURL(pdfUrl);
        }, 300000);
    }

    iframe.onload = function () { setTimeout(tryPrint, 800); };
    // Fallback als onload niet afvuurt (bijv. bij sommige PDF-viewers in WebView2)
    setTimeout(tryPrint, 5000);
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

// Functie registreren bij Office
if (typeof Office !== 'undefined') {
    Office.actions.associate('printAsPdf', printAsPdf);
}
