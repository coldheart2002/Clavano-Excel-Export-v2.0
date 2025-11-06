// Map Kintone fieldCodes → Excel Sheet + Cell
const fieldToExcelMap = {
  date: { sheet: "QUOTATION TEMPLATE", cell: "H11" }, // Date B1
  quotationNo: { sheet: "QUOTATION TEMPLATE", cell: "Q11" }, // Quote No.
  customer: { sheet: "QUOTATION TEMPLATE", cell: "H17" }, // Customer B3
  address: { sheet: "QUOTATION TEMPLATE", cell: "H18" }, // Address B4

  colour: { sheet: "QUOTATION TEMPLATE", cell: "J25" }, // No. of Side B9
  contactPerson: { sheet: "QUOTATION TEMPLATE", cell: "H13" }, // Contact Person B10

  contactNumber: {
    sheet: "QUOTATION TEMPLATE",
    cell: "H14",
    extract: (value, ws, cell) => {
      // Convert to number if possible
      const numericValue = Number(value);
      ws.getCell(cell).value = isNaN(numericValue) ? value : numericValue;
      return null;
    },
  }, // Contact Number

  emailAddress: { sheet: "QUOTATION TEMPLATE", cell: "H15" }, // Email Add

  paper: { sheet: "QUOTATION TEMPLATE", cell: "J24" }, // Cover-Front

  orderQuantity: {
    sheet: "QUOTATION TEMPLATE",
    cell: "P21",
    extract: (value, ws, cell) => {
      const numericValue = Number(value);
      ws.getCell(cell).value = isNaN(numericValue) ? value : numericValue;
      return null;
    },
  }, // Order Quantity

  itemDescription: { sheet: "QUOTATION TEMPLATE", cell: "G21" }, // Item Description

  size: { sheet: "QUOTATION TEMPLATE", cell: "J23" }, // Size

  officialUnitPrice: {
    sheet: "QUOTATION TEMPLATE",
    cell: "Q21",
    extract: (value, ws, cell) => {
      const numericValue = Number(value);
      ws.getCell(cell).value = isNaN(numericValue) ? value : numericValue;
      if (!isNaN(numericValue)) {
        ws.getCell(cell).numFmt = "#,##0.00";
      }
      return null;
    },
  }, // Unit Price

  totalAmount: {
    sheet: "QUOTATION TEMPLATE",
    cell: "R21",
    extract: (value, ws, cell) => {
      const numericValue = Number(value);
      ws.getCell(cell).value = isNaN(numericValue) ? value : numericValue;
      return null;
    },
  },

  signature: {
    sheet: "QUOTATION TEMPLATE",
    cell: "G51",
  },
};

module.exports = fieldToExcelMap;
