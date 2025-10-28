// Map Kintone fieldCodes → Excel Sheet + Cell
const fieldToExcelMap = {
  date: { sheet: "INPUT DATA", cell: "B1" }, // Date

  customer: { sheet: "INPUT DATA", cell: "B3" }, // Customer
  address: { sheet: "INPUT DATA", cell: "B4" }, // Address

  colour: { sheet: "INPUT DATA", cell: "B9" }, // No. of Side
  contactPerson: { sheet: "INPUT DATA", cell: "B10" }, // Contact Person

  contactNumber: {
    sheet: "INPUT DATA",
    cell: "B11",
    extract: (value, ws, cell) => {
      // Convert to number if possible
      const numericValue = Number(value);
      ws.getCell(cell).value = isNaN(numericValue) ? value : numericValue;
      return null; // handled manually
    },
  },

  emailAddress: { sheet: "INPUT DATA", cell: "B12" }, // Email Add

  paper: { sheet: "INPUT DATA", cell: "B18" }, // Cover-Front

  orderQuantity: {
    sheet: "INPUT DATA",
    cell: "B22",
    extract: (value, ws, cell) => {
      const numericValue = Number(value);
      ws.getCell(cell).value = isNaN(numericValue) ? value : numericValue;
      return null;
    },
  },

  itemDescription: { sheet: "INPUT DATA", cell: "B23" }, // Item Description

  size: { sheet: "INPUT DATA", cell: "B25" }, // Size

  unitPrice: {
    sheet: "INPUT DATA",
    cell: "B27",
    extract: (value, ws, cell) => {
      const numericValue = Number(value);
      ws.getCell(cell).value = isNaN(numericValue) ? value : numericValue;
      if (!isNaN(numericValue)) {
        ws.getCell(cell).numFmt = "#,##0.00"; // Format as number with 2 decimals
      }
      return null;
    },
  },
};

module.exports = fieldToExcelMap;
