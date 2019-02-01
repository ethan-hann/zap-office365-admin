// get a single sku
const getSku = (z, bundle) => {
  const responsePromise = z.request({
    url: `https://graph.microsoft.com/v1.0/subscribedSkus/${bundle.inputData.id}`,
  });
  return responsePromise
    .then(response => z.JSON.parse(response.content));
};

// get a list of skus
const listSkus = (z) => {
  const responsePromise = z.request({
    url: 'https://graph.microsoft.com/v1.0/subscribedSkus',
  });
  return responsePromise
    .then(response => {
      return z.JSON.parse(response.content).value;
    });
};

module.exports = {
  key: 'sku',
  noun: 'Sku',

  get: {
    display: {
      label: 'Get Sku',
      description: 'Retrieve a specific commercial subscription that your organization has acquired'
    },
    operation: {
      inputFields: [
        {key: 'id', required: true, label: "SKU ID"}
      ],
      perform: getSku
    }
  },

  list: {
    display: {
      label: 'New Sku',
      description: 'Retrieve the list of commercial subscriptions that your organization has acquired'
    },
    operation: {
      perform: listSkus
    }
  },

  sample: {
    capabilityStatus: "Enabled",
    consumedUnits: 14,
    id: "48a80680-7326-48cd-9935-b556b81d3a4e_c7df2760-2c81-4ef7-b578-5b5392b571df",
    prepaidUnits: {
      enabled: 25,
      suspended: 0,
      warning: 0
    },
    servicePlans: [
      {
        servicePlanId: "8c098270-9dd4-4350-9b30-ba4703f3b36b",
        servicePlanName: "ADALLOM_S_O365",
        provisioningStatus: "Success",
        appliesTo: "User"
      }
    ],
    skuId: "c7df2760-2c81-4ef7-b578-5b5392b571df",
    skuPartNumber: "ENTERPRISEPREMIUM",
    appliesTo: "User"
  },

  outputFields: [
    {key: 'capabilityStatus', label: 'Capability Status'},
    {key: 'consumedUnits', label: 'Consumed Units', type: 'integer'},
    {key: 'id', label: 'ID'},
    {key: 'prepaidUnits__enabled', label: 'Prepaid Units Enabled', type: 'integer'},
    {key: 'prepaidUnits__suspended', label: 'Prepaid Units Suspended', type: 'integer'},
    {key: 'prepaidUnits__warning', label: 'Prepaid Units Warning', type: 'integer'},
    {key: 'servicePlans[]servicePlanId', label: 'Service Plan ID'},
    {key: 'servicePlans[]servicePlanName', label: 'Service Plan Name'},
    {key: 'servicePlans[]provisioningStatus', label: 'Provisioning Status'},
    {key: 'servicePlans[]appliesTo', label: 'Applies To'},
    {key: 'skuId', label: 'SKU ID'},
    {key: 'skuPartNumber', label: 'SKU Part Number'},
    {key: 'appliesTo', label: 'Applies To'}
  ]
};
