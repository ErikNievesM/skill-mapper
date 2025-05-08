Office.onReady(() => {
    console.log("Office.js ready");

    $(document).ready(() => {
        console.log("DOM ready, loading reference data...");
        loadReferenceData();

        $('#platformSelect').on('change', () => {
            console.log("Platform changed");
            onPlatformChange();
        });

        $('#productSelect').on('change', () => {
            console.log("Product changed");
            onProductChange();
        });

        $('#moduleSelect').on('change', () => {
            console.log("Module changed");
            onModuleChange();
        });

        $('#insertButton').on('click', () => {
            console.log("Insert button clicked");
            insertToDeliverySheet();
        });
    });
});

let referenceData = [];

function notify(message, isError = false) {
    const container = document.getElementById("notification-area");
    if (container) {
        container.innerHTML = `<div style="color:${isError ? 'red' : 'green'}; margin-top:10px;">${message}</div>`;
    }
    console.log((isError ? "ERROR: " : "INFO: ") + message);
}

async function loadReferenceData() {
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem("Reference");
            const range = sheet.getUsedRange();
            range.load("values");
            await context.sync();

            referenceData = range.values.slice(1); // skip header
            console.log("Loaded Reference data:", referenceData);

            const platforms = [...new Set(referenceData.map(row => row[0]))].sort();
            populateDropdown("#platformSelect", platforms);
            notify("Reference data loaded.");
        });
    } catch (e) {
        console.error("Failed to load Reference sheet data", e);
        notify("Error: Could not load 'Reference' sheet. Please check the sheet name and format.", true);
    }
}

function onPlatformChange() {
    const selected = $('#platformSelect').val();
    console.log("Selected platforms:", selected);

    const filtered = referenceData.filter(row => selected.includes(row[0]));
    const products = [...new Set(filtered.map(row => row[1]))].sort();

    populateDropdown("#productSelect", products);
    $('#moduleSelect').empty();
    $('#skillSelect').empty();
}

function onProductChange() {
    const platforms = $('#platformSelect').val();
    const products = $('#productSelect').val();
    console.log("Selected products:", products);

    const filtered = referenceData.filter(row =>
        platforms.includes(row[0]) && products.includes(row[1])
    );
    const modules = [...new Set(filtered.map(row => row[2]))].sort();

    populateDropdown("#moduleSelect", modules);
    $('#skillSelect').empty();
}

function onModuleChange() {
    const platforms = $('#platformSelect').val();
    const products = $('#productSelect').val();
    const modules = $('#moduleSelect').val();
    console.log("Selected modules:", modules);

    const filtered = referenceData.filter(row =>
        platforms.includes(row[0]) &&
        products.includes(row[1]) &&
        modules.includes(row[2])
    );
    const skills = [...new Set(filtered.map(row => row[3]))].sort();

    populateDropdown("#skillSelect", skills);
}

function populateDropdown(selector, items) {
    const $el = $(selector);
    $el.empty();
    items.forEach(item => {
        $el.append($('<option>', { value: item, text: item }));
    });
    console.log(`Populated ${selector} with:`, items);
}

async function insertToDeliverySheet() {
    const platforms = $('#platformSelect').val();
    const products = $('#productSelect').val();
    const modules = $('#moduleSelect').val();
    const skills = $('#skillSelect').val();

    if (!platforms?.length || !products?.length || !modules?.length || !skills?.length) {
        notify("Please make all selections before inserting.", true);
        return;
    }

    const platformStr = platforms.join(", ");
    const productStr = products.join(", ");
    const moduleStr = modules.join(", ");
    const skillStr = skills.join(", ");

    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem("Delivery");
            const activeCell = context.workbook.getSelectedRange();
            activeCell.load("rowIndex");
            await context.sync();

            const row = activeCell.rowIndex + 1;
            const range = sheet.getRange(`E${row}:H${row}`);
            range.values = [[platformStr, productStr, moduleStr, skillStr]];

            await context.sync();
            notify("Selections inserted into current row.");
        });
    } catch (e) {
        console.error("Error inserting into Delivery sheet:", e);
        notify("Error: Could not insert row. Check if 'Delivery' sheet exists.", true);
    }
}
