let records = [];

// Format Date
function formatDate(dateValue) {

    if (!dateValue) return "-";

    // Excel serial date
    if (!isNaN(dateValue)) {
        const date = new Date((Number(dateValue) - 25569) * 86400 * 1000);

        const day = String(date.getDate()).padStart(2, "0");
        const month = String(date.getMonth() + 1).padStart(2, "0");
        const year = date.getFullYear();

        return `${day}-${month}-${year}`;
    }

    return dateValue;
}

// Upload Excel
document.getElementById("excelFile").addEventListener("change", function (e) {

    const file = e.target.files[0];

    if (!file) return;

    const reader = new FileReader();

    reader.onload = function (event) {

        const data = new Uint8Array(event.target.result);

        const workbook = XLSX.read(data, {
            type: "array"
        });

        const sheet =
            workbook.Sheets["OOP Cases"] ||
            workbook.Sheets[workbook.SheetNames[0]];

        records = XLSX.utils.sheet_to_json(sheet, {
            header: 1,
            defval: ""
        });

        console.log("Excel Loaded");
        console.log(records[0]);

        alert("Excel Loaded Successfully");
    };

    reader.readAsArrayBuffer(file);
});

// Search Unit
function searchUnit() {

    if (records.length === 0) return;

    const searchText = document
        .getElementById("unitSearch")
        .value
        .trim()
        .toUpperCase();

    if (!searchText) return;

    const row = records.slice(1).find(r =>
        String(r[0] || "")
            .trim()
            .toUpperCase() === searchText
    );

    if (!row) {

        document.getElementById("customer").innerText = "No Record Found";
        document.getElementById("crm").innerText = "-";
        document.getElementById("frm").innerText = "-";
        document.getElementById("invoice").innerText = "-";
        document.getElementById("due").innerText = "-";

        document.getElementById("oldbsp").innerText = "₹0";
        document.getElementById("newbsp").innerText = "₹0";
        document.getElementById("totalamount").innerText = "₹0";
        document.getElementById("received").innerText = "₹0";
        document.getElementById("interest").innerText = "₹0";
        document.getElementById("payable").innerText = "₹0";

        return;
    }

    // Customer Information
    document.getElementById("customer").innerText = row[4] || "-";
    document.getElementById("crm").innerText = row[76] || "-";
    document.getElementById("frm").innerText = row[77] || "-";

    document.getElementById("invoice").innerText =
        formatDate(row[58]);

    document.getElementById("due").innerText =
        formatDate(row[59]);

    // Calculation Details
    document.getElementById("oldbsp").innerText =
        "₹ " + (parseFloat(row[18]) || 0).toLocaleString("en-IN");

    document.getElementById("newbsp").innerText =
        "₹ " + (parseFloat(row[25]) || 0).toLocaleString("en-IN");

    document.getElementById("totalamount").innerText =
        "₹ " + (parseFloat(row[41]) || 0).toLocaleString("en-IN");

    document.getElementById("received").innerText =
        "₹ " + (parseFloat(row[42]) || 0).toLocaleString("en-IN");

    document.getElementById("interest").innerText =
        "₹ " + (parseFloat(row[43]) || 0).toLocaleString("en-IN");

    // Net Payable
   const payableValue = row[44];

console.log("Net Payable Raw Value:", payableValue);

document.getElementById("payable").innerText =
    "₹ " + Number(payableValue || 0).toLocaleString("en-IN");

    
console.log("Selected Unit:", searchText);
console.log("Matched Row:", row);
console.log("Column 44:", row[44]);
}