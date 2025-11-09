const fieldToPdfMap = {
  date: { left: 181.5, top: 151.5, width: 250, height: 10 },
  quotationNo: { left: 503, top: 151.5, width: 250, height: 10 },
  customer: { left: 181.5, top: 176, width: 250, height: 10 },
  contactNumber: { left: 181.5, top: 187.5, width: 250, height: 10 },
  emailAddress: { left: 181.5, top: 199, width: 250, height: 10 },
  contactPerson: { left: 181.5, top: 222, width: 250, height: 10 },
  address: { left: 181.5, top: 233.5, width: 250, height: 10 },
  itemDescription: { left: 181.5, top: 266, width: 250, height: 10 },
  unitOfMeasure: { left: 290, top: 277, width: 250, height: 10 },
  size: { left: 290, top: 288.25, width: 250, height: 10 },
  paper: { left: 290, top: 299.5, width: 250, height: 10 },
  weight: { left: 290, top: 311, width: 250, height: 10 },
  colour: { left: 290, top: 322.5, width: 250, height: 10 },
  numberOfPages: { left: 290, top: 334, width: 250, height: 10 },
  specificationBinding: { left: 290, top: 345.5, width: 250, height: 10 },
  orderQuantity: { left: 458, top: 266, width: 250, height: 10 },
  officialUnitPrice: { left: 503, top: 266, width: 250, height: 10 },
  totalAmount: { left: 543, top: 266, width: 250, height: 10 },
  signature: { left: 125, top: 585, width: 120, height: 50, isImage: true },
  salesRepresentative: {
    left: 125,
    top: 610,
    width: 250,
    height: 10,
  },
};

module.exports = fieldToPdfMap;
