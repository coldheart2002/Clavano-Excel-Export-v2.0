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
  cover_papers: { sheet: "INPUT DATA", cell: "B18" },
  // cover_back: { sheet: "INPUT DATA", cell: "B19" },
  inside_page: { sheet: "INPUT DATA", cell: "B20" },
  finishing: { sheet: "INPUT DATA", cell: "B21" },
  contact_person: { sheet: "INPUT DATA", cell: "B10" },
  contact_number: { sheet: "INPUT DATA", cell: "B11" },
  email_address: { sheet: "INPUT DATA", cell: "B12" },
  remarks: { sheet: "INPUT DATA", cell: "B8" },
  representative: { sheet: "INPUT DATA", cell: "B14" },
  input_size: { sheet: "INPUT DATA", cell: "B25" },
};

module.exports = fieldToExcelMap;
