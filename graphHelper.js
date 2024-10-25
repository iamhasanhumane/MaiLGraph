const azure = require("@azure/identity");
const graph = require("@microsoft/microsoft-graph-client");
require("isomorphic-fetch");
const settings = require("./appSetting.js");
const {
  TokenCredentialAuthenticationProvider,
} = require("@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js");

let _settings = undefined;
let _deviceCodeCredential = undefined;
let _userClient = undefined;

function initializeGraphForUserAuth(settings, deviceCodePrompt) {
  //Ensure settings isn't null
  if (!settings) {
    throw new Error("Settings cannot be undefined");
  }

  _settings = settings;

  _deviceCodeCredential = new azure.DeviceCodeCredential({
    clientId: settings.clientId,
    tenantId: settings.tenantId,
    userPromptCallback: deviceCodePrompt,
  });

  const authProvider = new TokenCredentialAuthenticationProvider(
    _deviceCodeCredential,
    {
      scopes: settings.graphUserScopes,
    }
  );

  _userClient = graph.Client.initWithMiddleware({
    authProvider: authProvider,
  });
}

async function getUserTokenAsync() {
  //Ensure credential is'nt undefined
  if (!_deviceCodeCredential) {
    throw new Error("Graph has not been initialized for user auth");
  }

  //Ensure scopes isn't undefined
  if (!settings?.graphUserScopes) {
    throw new Error('Setting "scopes" cannot be undefined ');
  }

  //Request token with given scopes
  const response = await _deviceCodeCredential.getToken(
    _settings?.graphUserScopes
  );
  return response.token;
}

async function getUserAsync() {
  // Ensure client isn't undefined
  if (!_userClient) {
    throw new Error("Graph has not been initialized for user auth");
  }

  //Only request specific properties with .select()
  return _userClient
    .api("/me")
    .select(["displayName", "mail", "userPrincipalName"])
    .get();
}

async function getInboxAsync() {
  // Ensure client isn't undefined
  if (!_userClient) {
    throw new Error("Graph has not been initialized for user auth");
  }

  console.log(_userClient.api("/me").select(["mail"]).get());
  return _userClient
    .api("/me/mailFolders/inbox/messages")
    .select(["from", "isRead", "receivedDateTime", "subject"])
    .top(25)
    .orderby("receivedDateTime DESC")
    .get();
}

async function sendMailAsync(subject, body, recipient) {
  // Ensure client isn't undefined
  if (!_userClient) {
    throw new Error("Graph has not been initialized for user auth");
  }

  // Create a new message
  const message = {
    subject: subject,
    body: {
      content: body,
      contentType: "text",
    },
    toRecipients: [
      {
        emailAddress: {
          address: recipient,
        },
      },
    ],
  };

  // Send the message
  return _userClient.api("me/sendMail").post({
    message: message,
  });
}

async function makeGraphCallAsync() {
  // Ensure client isn't undefined
  if (!_userClient) {
    throw new Error("Graph has not been initialized for user auth");
  }

  const event = {
    subject: "Let's go for lunch",
    body: {
      contentType: "HTML",
      content: "Does noon work for you?",
    },
    start: {
      dateTime: "2017-04-15T12:00:00",
      timeZone: "Pacific Standard Time",
    },
    end: {
      dateTime: "2017-04-15T14:00:00",
      timeZone: "Pacific Standard Time",
    },
    location: {
      displayName: "Selva's Bar",
    },
    attendees: [
      {
        emailAddress: {
          address: "santhoshkumar@hoodshub.com",
          name: "Santhosh",
        },
        type: "required",
      },
    ],
    allowNewTimeProposals: true,
    transactionId: "7E163156-7762-4BEB-A1C6-729EA81755A7",
  };

  await _userClient.api("/me/events").post(event);
}

module.exports = {
  initializeGraphForUserAuth,
  getUserTokenAsync,
  getUserAsync,
  getInboxAsync,
  sendMailAsync,
  makeGraphCallAsync,
};
