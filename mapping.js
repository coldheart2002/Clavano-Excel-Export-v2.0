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
    // extract: (value, ws, cell) => {
    //   // Convert to number if possible
    //   const numericValue = Number(value);
    //   ws.getCell(cell).value = isNaN(numericValue) ? value : numericValue;
    //   return null;
    // },
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
    cell: "F49",
    isImage: true,
    width: 120, // width in pixels
    height: 50, // height in pixels
  },

  Created_by: {
    sheet: "QUOTATION TEMPLATE",
    cell: "F52",
    extract: (value, ws, cell) => {
      // Some user fields return an array, some a single object
      const name =
        Array.isArray(value) && value.length > 0
          ? value[0].name
          : value.name || "";
      ws.getCell(cell).value = name;
      return null; // prevents default assignment
    },
  },
};

module.exports = fieldToExcelMap;
