const baseMap = {
  date: { left: 181.5, top: 151.5, width: 250, height: 10 },
  quotationNo: { left: 503, top: 151.5, width: 250, height: 10 },
  customer: { left: 181.5, top: 176, width: 250, height: 10 },
  contactNumber: { left: 181.5, top: 187.5, width: 250, height: 10 },
  emailAddress: { left: 181.5, top: 199, width: 250, height: 10 },
  contactPerson: { left: 181.5, top: 222, width: 250, height: 10 },
  address: { left: 181.5, top: 233.5, width: 250, height: 10 },
  itemDescription: { left: 181.5, top: 266, width: 250, height: 10 },
  size: { left: 290, top: 288.25, width: 250, height: 10 },
  cover_paper_description: { left: 290, top: 299.5, width: 250, height: 10 },
  inside_paper_description: { left: 290, top: 322.5, width: 250, height: 10 },
  number_of_pages: { left: 290, top: 334, width: 250, height: 10 },
  specificationBinding: { left: 290, top: 345.5, width: 250, height: 10 },
  order_qty: { left: 458, top: 266, width: 250, height: 10 },

  signature: { left: 125, top: 585, width: 120, height: 50, isImage: true },
  salesRepresentative: { left: 125, top: 610, width: 250, height: 10 },
};

// 🔥 Dynamic fields for offset/digital
const dynamicFields = {
  offset: {
    unit_price_field: "offset_unit_selling_price_official",
    total_amount_field: "offset_total_amount",
  },

  digital: {
    unit_price_field: "digital_unit_selling_price_official",
    total_amount_field: "digital_total_amount",
  },
};

/**
 * Returns final mapping based on export type
 */
function getPdfMap(type) {
  const selected = dynamicFields[type] || dynamicFields.offset;

  return {
    ...baseMap,

    // dynamic mixing
    [selected.unit_price_field]: {
      left: 503,
      top: 266,
      width: 250,
      height: 10,
    },

    [selected.total_amount_field]: {
      left: 543,
      top: 266,
      width: 250,
      height: 10,
    },
  };
}

module.exports = getPdfMap;
