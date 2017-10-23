(function () {
        "use strict";

        // The initialize function is run each time the page is loaded.
        Office.initialize = function (reason) {
            $(document).ready(function () {

                // Use this to check whether the API is supported in the Word client.
                if (Office.context.requirements.isSetSupported('WordApi', 1.1)) {
                    // Do something that is only available via the new APIs
                    $('#insert').click(function() {insertImage();});
                    $('#getSelection').click(function() {getSelection();});
                    $('#getOoxml').click(function() {getOoxml();});
                } else {
                    // Just letting you know that this code will not work with your version of Word.
                    $('#supportedVersion').html('This code requires Word 2016 or greater.');
                }
            });
        };

        function insertImage() {
            Word.run(function (context) {

                // Create a proxy object for the document.
                var thisDocument = context.document;

                // Queue a command to get the current selection.
                // Create a proxy range object for the selection.
                var range = thisDocument.getSelection();
				
                var request = new XMLHttpRequest();
                request.open('GET', 'formula.base64', false);
                request.send(null);
                document.getElementById("output").innerHTML = request.responseText;

                // Queue a command to replace the selected text.
                var image = range.insertInlinePictureFromBase64(request.responseText, Word.InsertLocation.replace);
                image.height = 50;

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added an image.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }
        
        function getSelection() {
            Word.run(function (context) {

                // Create a proxy object for the document.
                var thisDocument = context.document;

                // Queue a command to get the current selection.
                // Create a proxy range object for the selection.
                var range = thisDocument.getSelection();
                var xml = range.getOoxml();
                return context.sync().then(function () {
                    var v=xml.value;
                    v=v.replaceAll("<","&lt;")
                    document.getElementById("output").innerHTML = v;
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
            
        }

        function getSelectionOoxml() {
            Word.run(function (context) {

                // Create a proxy object for the document.
                var thisDocument = context.document;

                // Queue a command to get the current selection.
                // Create a proxy range object for the selection.
                var range = thisDocument.getSelection();
                var html = range.getHtml();
                return context.sync().then(function () {
                    document.getElementById("output").innerHTML = html.value;
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
            
        }

        String.prototype.replaceAll = function(find, replace) {
            var str = this;
            return str.replace(new RegExp(find.replace(/([.*+?^=!:${}()|\[\]\/\\])/g, "\\$1"), 'g'), replace);
        };
        
        
        })();