require("dotenv").config();

const settings = {
  clientId: process.env.clientId,
  tenantId: "common",
  graphUserScopes: ["user.read", "mail.read", "mail.send"],
};

module.exports = settings;
