const readline = require("readline-sync");
const settings = require("./appSetting.js");
const {
  initializeGraphForUserAuth,
  getUserAsync,
  getUserTokenAsync,
  getInboxAsync,
  sendMailAsync,
  makeGraphCallAsync,
} = require("./graphHelper.js");

async function main() {
  console.log("Javascript Graph Tutorial");

  let choice = 0;

  //Initialize Graph
  initializeGraph(settings);

  //Greet the user by name
  await greetUserAsync();

  const choices = [
    "Display access token",
    "List my inbox",
    "Send Mail",
    "Make a Graph Call",
  ];

  while (choice != -1) {
    choice = readline.keyInSelect(choices, "Select an option ", {
      cancel: "Exit",
    });
    switch (choice) {
      case -1:
        //Exit
        console.log("GoodBye");
        break;
      case 0:
        //Display access token
        await displayAccessTokenAsync();
        break;
      case 1:
        // List emails from user's inbox
        await listInboxAsync();
        break;
      case 2:
        // send an email message
        await sendMailToSelfAsync();
        break;
      case 3:
        // Run any graph code
        await doGraphCallAsync();
        break;
      default:
        console.log("Invalid Choice! Please Try Again!");
    }
  }
}

main();

function initializeGraph(settings) {
  initializeGraphForUserAuth(settings, (info) => {
    //Display the device code message to
    // the user. This tells them
    // where to go to signin and provides the
    // code to use.
    console.log(info.message);
  });
}

async function greetUserAsync() {
  try {
    const user = await getUserAsync();
    console.log(`Hello, ${user?.displayName}!`);
    //For work/school accounts email is in Mail Property
    //For personal accounts , email is in userPrincipalName
    console.log(`Email: ${user?.mail ?? user?.userPrincipalName ?? ""}`);
  } catch (error) {
    console.log(`Error getting user ${error}`);
  }
}

async function displayAccessTokenAsync() {
  try {
    const userToken = await getUserTokenAsync();
    console.log(`User Token: ${userToken}`);
  } catch (error) {
    console.log(`Error getting user access token: ${error}`);
  }
}

async function listInboxAsync() {
  try {
    const messagePage = await getInboxAsync();
    console.log(messagePage);
    const messages = await messagePage.value;

    console.log(messages);

    //output each message's details
    for (const message of messages) {
      console.log(`Message: ${message.subject ?? "No Subject"}`);
      console.log(`   From: ${message.from?.emailaddress?.name ?? "UNKNOWN"}`);
      console.log(`   Status: ${message.isRead ? "READ" : "Unread"}`);
      console.log(`   Received: ${message.receivedDateTime}`);
    }

    // If @odata.nextLink is not undefined, there are more messages
    // available on the server
    const moreAvailable = messagePage["@odata.nextLink"] != undefined;
    console.log(`\nMore messages available? ${moreAvailable}`);
  } catch (error) {
    console.log(`Error getting users mail ${error}`);
  }
}

async function sendMailToSelfAsync() {
  try {
    // Send mail to the signed-in user
    // Get the user for their email address
    const user = await getUserAsync();
    const userEmail = user?.mail ?? user?.userPrincipalName;

    if (!userEmail) {
      console.log("Couldn't get your email address, canceling...");
      return;
    }

    await sendMailAsync("Testing Microsoft Graph", "Hello world", userEmail);
    console.log("Mail Sent.");
  } catch (error) {
    console.log(`Error in sending mail ${error}`);
  }
}

async function doGraphCallAsync() {
  try {
    await makeGraphCallAsync();
  } catch (err) {
    console.log(`Error making Graph call: ${err}`);
  }
}
