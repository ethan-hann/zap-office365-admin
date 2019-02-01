// get a single user
const getUser = (z, bundle) => {
  const responsePromise = z.request({
    url: `https://graph.microsoft.com/v1.0/users/${bundle.inputData.id}`,
  });
  return responsePromise
    .then(response => z.JSON.parse(response.content));
};

// get a list of users
const listUsers = (z) => {
  const responsePromise = z.request({
    url: 'https://graph.microsoft.com/v1.0/users'
  });
  return responsePromise
    .then(response => z.JSON.parse(response.content).value);
};

// create a user
const createUser = (z, bundle) => {
  const responsePromise = z.request({
    method: 'POST',
    url: 'https://graph.microsoft.com/v1.0/users',
    body: {
      accountEnabled: bundle.inputData.accountEnabled,
      displayName: bundle.inputData.displayName,
      givenName: bundle.inputData.givenName,
      surname: bundle.inputData.surname,
      onPremisesImmutableId: bundle.inputData.onPremisesImmutableId,
      mailNickname: bundle.inputData.mailNickname,
      passwordProfile: {
        forceChangePasswordNextSignIn: bundle.inputData.forceChangePasswordNextSignIn,
        password: bundle.inputData.password
      },
      userPrincipalName: bundle.inputData.userPrincipalName,
      usageLocation: bundle.inputData.usageLocation
    }
  });
  return responsePromise
    .then(response => z.JSON.parse(response.content));
};

module.exports = {
  key: 'user',
  noun: 'User',

  get: {
    display: {
      label: 'Get User',
      description: 'Gets a user.'
    },
    operation: {
      inputFields: [
        {key: 'id', required: true}
      ],
      perform: getUser
    }
  },

  list: {
    display: {
      label: 'New User',
      description: 'Lists the users.'
    },
    operation: {
      perform: listUsers
    }
  },

  create: {
    display: {
      label: 'Create User',
      description: 'Creates a new user.'
    },
    operation: {
      inputFields: [
        {key: 'accountEnabled', required: true, type: 'boolean', helpText: 'true if the account is enabled; otherwise, false', label: 'Account Enabled'},
        {key: 'displayName', required: true, type: 'string', helpText: 'The name to display in the address book for the user', label: 'Display name'},
        {key: 'givenName', required: false, type: 'string', helpText: 'The given name (first name) of the user', label: 'Given name'},
        {key: 'surname', required: false, type: 'string', helpText: 'The user\'s surname (family name or last name)', label: 'Surname'},
        {key: 'onPremisesImmutableId', required: false, type: 'string', helpText: 'Only needs to be specified when creating a new user account if you are using a federated domain for the user\'s userPrincipalName (UPN) property', label: 'On Premises Immutable ID'},
        {key: 'mailNickname', required: true, type: 'string', helpText: 'The mail alias for the user', label: 'Mail Nickname'},
        {key: 'forceChangePasswordNextSignIn', required: true, type: 'boolean', helpText: 'true if the user must change her password on the next login; otherwise false', label: 'Force Change Password on Next Sign In'},
        {key: 'password', required: true, type: 'string', helpText: 'The password for the user. This property is required when a user is created. It can be updated, but the user will be required to change the password on the next login. The password must satisfy minimum requirements as specified by the user’s passwordPolicies property. By default, a strong password is required', label: 'Password'},
        {key: 'userPrincipalName', required: true, type: 'string', helpText: 'The user principal name (someuser@contoso.com)', label: 'User Principal Name'},
        {key: 'usageLocation', required: false, type: 'string', helpText: 'A two letter country code (ISO standard 3166). Required for users that will be assigned licenses due to legal requirement to check for availability of services in countries', label: 'Usage Location'}
      ],
      perform: createUser
    },
  },

  sample: {
      businessPhones: [
        "businessPhones-value"
      ],
      displayName: "displayName-value",
      givenName: "givenName-value",
      jobTitle: "jobTitle-value",
      mail: "mail-value",
      mobilePhone: "mobilePhone-value",
      officeLocation: "officeLocation-value",
      preferredLanguage: "preferredLanguage-value",
      surname: "surname-value",
      userPrincipalName: "userPrincipalName-value",
      id: "id-value"
    },

  outputFields: [
    {key: 'businessPhones:[]', label: "Business Phone"},
    {key: 'displayName', label: 'Display Name'},
    {key: 'givenName', label: 'Given Name'},
    {key: 'jobTitle', label: 'Job Title'},
    {key: 'mail', label: 'Mail'},
    {key: 'mobilePhone', label: 'Mobile Phone'},
    {key: 'officeLocation', label: 'Office Location'},
    {key: 'preferredLanguage', label: 'Preferred Language'},
    {key: 'surname', label: 'Surname'},
    {key: 'userPrincipalName', label: 'User Principal Name'}
  ]
};
