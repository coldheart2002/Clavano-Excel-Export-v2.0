// Map Kintone fieldCodes → Excel Sheet + Cell
const fieldToExcelMap = {
  customer: { sheet: "INPUT DATA", cell: "B3" },
  date: { sheet: "INPUT DATA", cell: "B1" },
  requested_delivery_date: { sheet: "INPUT DATA", cell: "B2" },
  address: { sheet: "INPUT DATA", cell: "B4" },
  number_of_pages: { sheet: "INPUT DATA", cell: "B7" },
  number_of_side: { sheet: "INPUT DATA", cell: "B9" },
  order_qty: { sheet: "INPUT DATA", cell: "B22" },
  final_size_w: { sheet: "INPUT DATA", cell: "C25" },
  final_size_l: { sheet: "INPUT DATA", cell: "D25" },
  finishing: { sheet: "INPUT DATA", cell: "B21" },
  contact_person: { sheet: "INPUT DATA", cell: "B10" },
  contact_number: { sheet: "INPUT DATA", cell: "B11" },
  email_address: { sheet: "INPUT DATA", cell: "B12" },
  remarks: { sheet: "INPUT DATA", cell: "B8" },
  input_size: { sheet: "INPUT DATA", cell: "B25" },
  item_description: { sheet: "INPUT DATA", cell: "B23" },

  // Special case: mark_up → percentage
  mark_up: {
    sheet: "INPUT DATA",
    cell: "D17",
    extract: (val, ws, cell) => {
      if (!val) return "";
      const num = parseFloat(val);
      if (isNaN(num)) return val;

      // Write to Excel
      ws.getCell(cell).value = num / 100; // convert 15 → 0.15
      ws.getCell(cell).numFmt = "0%"; // show as 15%
      return null; // we already set it directly
    },
  },

  // ✅ Sales Representative (Concatenate)
  // sales_representative: {
  //   sheet: "INPUT DATA",
  //   cell: "B14",
  //   concat: true,
  //   extract: (val, concat) => {
  //     if (!Array.isArray(val) || val.length === 0) return "";
  //     if (concat) {
  //       return val.map((u) => `${u.name} (${u.code})`).join(", ");
  //     } else {
  //       return val[0]?.name || "";
  //     }
  //   },
  // },

  //Default
  sales_representative: {
    sheet: "INPUT DATA",
    cell: "B14",
    extract: (val) => (Array.isArray(val) && val[0]?.name) || "",
  },

  // ✅ Front Cover Papers
  front_cover_papers: {
    sheet: "INPUT DATA",
    cell: "B18",
    concat: true,
    extract: (val, concat) => {
      if (!Array.isArray(val) || val.length === 0) return "";
      if (concat) {
        return val
          .map(
            (u) =>
              `${u.value.front_cover_paper?.value || ""} (${
                u.value.fcp_print_output?.value || ""
              })`
          )
          .join(", ");
      } else {
        return val[0]?.value?.front_cover_paper?.value || "";
      }
    },
  },

  // ✅ Back Cover Papers
  back_cover_papers: {
    sheet: "INPUT DATA",
    cell: "B19",
    concat: true,
    extract: (val, concat) => {
      if (!Array.isArray(val) || val.length === 0) return "";
      if (concat) {
        return val
          .map(
            (u) =>
              `${u.value.back_cover_paper?.value || ""} (${
                u.value.bcp_print_output?.value || ""
              })`
          )
          .join(", ");
      } else {
        return val[0]?.value?.back_cover_paper?.value || "";
      }
    },
  },

  // ✅ Inside Papers
  inside_papers: {
    sheet: "INPUT DATA",
    cell: "B20",
    concat: true,
    extract: (val, concat) => {
      if (!Array.isArray(val) || val.length === 0) return "";
      if (concat) {
        return val
          .map(
            (u) =>
              `${u.value.inside_paper?.value || ""} (${
                u.value.inside_paper_print_output?.value || ""
              })`
          )
          .join(", ");
      } else {
        return val[0]?.value?.inside_paper?.value || "";
      }
    },
  },
};

module.exports = fieldToExcelMap;
