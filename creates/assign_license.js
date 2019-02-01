// create a particular licenseuser by name
const createAssignLicense = (z, bundle) => {
  const responsePromise = z.request({
    method: 'POST',
    url: `https://graph.microsoft.com/v1.0/users/${bundle.inputData.userPrincipalName}/assignLicense`,
    body: {
        addLicenses: [
        {
          skuId: bundle.inputData.skuId
        }
      ], removeLicenses: []}
  });
  return responsePromise
    .then(response => z.JSON.parse(response.content));
};

module.exports = {
  key: 'assignLicense',
  noun: 'license',

  display: {
    label: 'Assign License',
    description: 'Assign a purchased license to a user.'
  },

  operation: {
    inputFields: [
      {
        key: 'userPrincipalName',
        required: true,
        label: 'User Principal Name',
        dynamic: 'userList.userPrincipalName.displayname'
      },
      {
        key: 'skuId',
        required: true,
        label: 'Product',
        dynamic: 'skuList.skuId.skuPartNumber'
      }
    ],
    perform: createAssignLicense,

    sample: {
      accountEnabled: true,
      assignedLicenses: [
        {
          disabledPlans: [ "11b0131d-43c8-4bbb-b2c8-e80f9a50834a" ],
          skuId: "0118A350-71FC-4EC3-8F0C-6A1CB8867561"
        }
      ],
      assignedPlans: [
        {
          assignedDateTime: "2016-10-02T12:13:14Z",
          capabilityStatus: "capabilityStatus-value",
          service: "service-value",
          servicePlanId: "bea13e0c-3828-4daa-a392-28af7ff61a0f"
        }
      ],
      businessPhones: [
        "businessPhones-value"
      ],
      city: "city-value",
      companyName: "companyName-value"
    },
    outputFields: [
      {key: 'id', label: 'ID'},
      {key: 'accountEnabled', label: 'ID'},
      {key: 'assignedLicenses[]disabledPlans[]', label: 'ID'},
      {key: 'assignedLicenses[]skuId', label: 'ID'},
      {key: 'assignedPlans[]assignedDateTime', label: 'ID'},
      {key: 'assignedPlans[]capabilityStatus', label: 'ID'},
      {key: 'assignedPlans[]service', label: 'ID'},
      {key: 'assignedPlans[]servicePlanId', label: 'ID'},
      {key: 'businessPhones[]', label: 'ID'},
      {key: 'city', label: 'ID'},
      {key: 'companyName', label: 'ID'}
    ]
  }
};
