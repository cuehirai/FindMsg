import { PublicClientApplication } from "@azure/msal-browser";

const appId = "67dc74f0-fb8c-47a4-8435-63253a3fc3bb"; // FindMsg
// const appId = "8fe75eb2-476b-43e1-8e36-a09cf42fcb42"; // FindMsg-NoConsentTest

export const client = new PublicClientApplication({
    auth: {
        clientId: appId,
        authority: "https://login.microsoftonline.com/organizations/",
        navigateToLoginRequestUrl: false,
    },
    cache: { cacheLocation: "localStorage" },
});
