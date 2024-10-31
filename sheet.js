document.addEventListener("DOMContentLoaded", function () {
    const urlParams = new URLSearchParams(window.location.search);
    const fileUrl = urlParams.get("fileUrl");
    const sheetContentDiv = document.getElementById("sheet-content");

    let workbook;
    let activeSheet;
    
    if (fileUrl) {
        fetch(fileUrl)
            .then(response => response.arrayBuffer())
            .then(data => {
                workbook = XLSX.read(data, { type: 'array' });
                displaySheet(workbook.SheetNames[0]); // Display the first sheet initially
                addSheetTabs(workbook.SheetNames);
            })
            .catch(error => {
                alert("Failed to load Excel file. Please check the URL.");
                console.error(error);
            });
    } else {
        alert("No file URL provided.");
    }

    function displaySheet(sheetName) {
        activeSheet = workbook.Sheets[sheetName];
        const htmlString = XLSX.utils.sheet_to_html(activeSheet, { id: "sheet-table", editable: false });
        sheetContentDiv.innerHTML = htmlString;
    }

    function addSheetTabs(sheetNames) {
        const sheetTabs = document.createElement("div");
        sheetTabs.classList.add("sheet-tabs");

        sheetNames.forEach(sheetName => {
            const tabButton = document.createElement("button");
            tabButton.classList.add("sheet-tab");
            tabButton.innerText = sheetName;
            tabButton.addEventListener("click", () => displaySheet(sheetName));
            sheetTabs.appendChild(tabButton);
        });

        document.body.insertBefore(sheetTabs, sheetContentDiv);
    }

    document.getElementById("apply-operation").addEventListener("click", function () {
        const primaryColumn = document.getElementById("primary-column").value;
        const operationColumns = document.getElementById("operation-columns").value.split(",");
        const operationType = document.getElementById("operation-type").value;
        const operation = document.getElementById("operation").value;

        if (primaryColumn && operationColumns.length) {
            const filteredData = applyFilterOperation(primaryColumn, operationColumns, operationType, operation);
            const worksheet = XLSX.utils.aoa_to_sheet(filteredData);
            const htmlString = XLSX.utils.sheet_to_html(worksheet, { id: "sheet-table", editable: false });
            sheetContentDiv.innerHTML = htmlString;
        } else {
            alert("Please enter valid columns for filtering.");
        }
    });

    function applyFilterOperation(primaryColumn, operationColumns, operationType, operation) {
        const jsonSheetData = XLSX.utils.sheet_to_json(activeSheet, { header: 1 });
        const primaryIndex = getColumnIndex(primaryColumn);

        return jsonSheetData.filter((row, rowIndex) => {
            if (rowIndex === 0) return true;

            return operationColumns.every(col => {
                const colIndex = getColumnIndex(col);
                return operationType === "and"
                    ? (operation === "null" ? !row[colIndex] : row[colIndex])
                    : (operation === "null" ? !row[colIndex] || !row[primaryIndex] : row[colIndex] || row[primaryIndex]);
            });
        });
    }

    function getColumnIndex(col) {
        return col.toUpperCase().charCodeAt(0) - 65;
    }

    document.getElementById("download-button").addEventListener("click", function () {
        document.getElementById("download-modal").style.display = "flex";
    });

    document.getElementById("confirm-download").addEventListener("click", function () {
        const filename = document.getElementById("filename").value || "Sheet";
        const fileFormat = document.getElementById("file-format").value;
        const sheetTable = document.getElementById("sheet-table");

        const data = XLSX.utils.table_to_sheet(sheetTable);
        const newWorkbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWorkbook, data, "Sheet1");

        if (fileFormat === "xlsx") {
            XLSX.writeFile(newWorkbook, `${filename}.xlsx`);
        } else {
            XLSX.writeFile(newWorkbook, `${filename}.csv`);
        }

        document.getElementById("download-modal").style.display = "none";
    });

    document.getElementById("close-modal").addEventListener("click", function () {
        document.getElementById("download-modal").style.display = "none";
    });
});
