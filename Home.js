
    let cellToHighlight;
    let messageBanner;

    // Initialization when Office JS and JQuery are ready.
    Office.onReady(() => {
        $(() => {
            // Initialize he Office Fabric UI notification mechanism and hide it.
            const element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();
            
            // If not using Excel 2016 or later, use fallback logic.
            if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
                $("#template-description").text("This sample will display the value of the cells that you have selected in the spreadsheet.");
                $('#button-text').text("Display!");
                $('#button-desc').text("Display the selection");

                $('#highlight-button').on('click',displaySelectedCells);
                return;
            }

            $("#template-description").text("This sample highlights the highest value from the cells you have selected in the spreadsheet.");
            $('#button-text').text("Highlight!");
            $('#button-desc').text("Highlights the largest number.");
                
            loadSampleData();

            // Add a click event handler for the highlight button.
            $('#highlight-button').on('click',highlightHighestValue);
        });
    });

    async function loadSampleData() {
        const values = [
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)]
        ];

        try {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                // Write sample values to a range in the active worksheet
                sheet.getRange("B3:D5").values = values;
                await context.sync();
            });
        } catch (error) {
            errorHandler(error);
        }
    }

    async function highlightHighestValue() {
        try {
            await Excel.run(async (context) => {
                const sourceRange = context.workbook.getSelectedRange().load("values, rowCount, columnCount");

                await context.sync();
                let highestRow = 0;
                let highestCol = 0;
                let highestValue = sourceRange.values[0][0];

                // Find the cell to highlight
                for (let i = 0; i < sourceRange.rowCount; i++) {
                    for (let j = 0; j < sourceRange.columnCount; j++) {
                        if (!isNaN(sourceRange.values[i][j]) && sourceRange.values[i][j] > highestValue) {
                            highestRow = i;
                            highestCol = j;
                            highestValue = sourceRange.values[i][j];
                        }
                    }
                }

                cellToHighlight = sourceRange.getCell(highestRow, highestCol);
                sourceRange.worksheet.getUsedRange().format.fill.clear();
                sourceRange.worksheet.getUsedRange().format.font.bold = false;

                // Highlight the cell
                cellToHighlight.format.fill.color = "orange";
                cellToHighlight.format.font.bold = true;
                await context.sync;
            });
        } catch (error) {
            errorHandler(error);
        }
    }

    async function displaySelectedCells() {
        try {
            await Excel.run(async (context) => {
                const range = context.workbook.getSelectedRange();
                range.load("text");
                await context.sync();
                const textValue = range.text.toString();
                showNotification('The selected text is:', '"' + textValue + '"');
            });
        } catch (error) {
            errorHandler(error);
        }
    }

    // Helper function for treating errors
    function errorHandler(error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        showNotification("Error", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
