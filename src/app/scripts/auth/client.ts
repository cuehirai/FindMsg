import { PublicClientApplication } from "@azure/msal-browser";
import { AppConfig } from "../../../config/AppConfig";

// const appId = "2e9f73d3-65f6-4ed6-8fbb-6ec025d9cde7"; // FindMsg
// const appId = "8fe75eb2-476b-43e1-8e36-a09cf42fcb42"; // FindMsg-NoConsentTest
// const appId = "6f5036ec-ad6d-4b29-87cb-010b0702bce4"; // SeaChaSt
const appId = AppConfig.AuthClient.AppId;

export const client = new PublicClientApplication({
    auth: {
        clientId: appId,
        authority: "https://login.microsoftonline.com/organizations/",
        navigateToLoginRequestUrl: false,
    },
    cache: { cacheLocation: "localStorage" },
});
