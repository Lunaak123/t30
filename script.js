document.getElementById('submit-link').addEventListener('click', () => {
    const excelUrl = document.getElementById('excel-url').value;
    if (!excelUrl) {
        alert("Please enter a valid Excel file URL.");
        return;
    }
    window.location.href = `sheet.html?fileUrl=${encodeURIComponent(excelUrl)}`;
});
