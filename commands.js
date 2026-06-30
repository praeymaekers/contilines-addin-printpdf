/*
 * Print als PDF - Conti-Lines NV
 * v1.2 - dialoog meldt 'ready' voor data verstuurd wordt
 */

Office.onReady(() => {});

function printAsPdf(event) {
    getDocumentAsPdf()
        .then(function (pdfBlob) {
            return blobToBase64(pdfBlob);
        })
        .then(function (base64DataUrl) {
            Office.context.ui.displayDialogAsync(
                'https://praeymaekers.github.io/contilines-addin-printpdf/printdialog.html',
                { height: 80, width: 60, displayInIframe: false },
                function (result) {
                    if (result.status === Office.AsyncResultStatus.Failed) {
                        alert('Kon printvenster niet openen: ' + result.error.message);
                        event.completed();
                        return;
                    }

                    var dialog = result.value;

                    dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (args) {
                        if (args.message === 'ready') {
                            // Dialoog is geïnitialiseerd — nu pas data sturen
                            dialog.messageChild(base64DataUrl);
                        } else if (args.message === 'printed' || args.message === 'closed') {
                            dialog.close();
                            event.completed();
                        }
                    });

                    dialog.addEventHandler(Office.EventType.DialogEventReceived, function () {
                        event.completed();
                    });
                }
            );
        })
        .catch(function (err) {
            event.completed({ allowEvent: false });
            alert('Fout bij aanmaken PDF: ' + err.message);
        });
}

function blobToBase64(blob) {
    return new Promise(function (resolve, reject) {
        var reader = new FileReader();
        reader.onload = function (e) { resolve(e.target.result); };
        reader.onerror = reject;
        reader.readAsDataURL(blob);
    });
}

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

if (typeof Office !== 'undefined') {
    Office.actions.associate('printAsPdf', printAsPdf);
}
