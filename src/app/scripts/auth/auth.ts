import * as msal from "@azure/msal-browser";
import * as msalC from "@azure/msal-common";
import * as teams from "@microsoft/teams-js";
import * as log from '../logger';
import { isUnknownObject } from '../utils';
import { BrowserAuthErrorMessage } from "@azure/msal-browser";
import { IAuthMessages } from "./IAuthMessages";
import { client } from './client';
import { AI } from '../appInsights';

/*
  The assumed workflow made here is that the tenant's global admin has consented to all permission that the app requires before normal users get to see it.
  Flow is:
  - log in which user-consentable scopes only
  - check the returned scopes
  - if all are present, consent was given
  - if not, show an error with a button to redirect to admin consent flow
  - this should be ok for normal users and admins
*/

/*
    Note: DO NOT request 'offline_access' scope when calling acquireTokenSilent().
    'offline_access' is not returned as a scope in the access token,
    so acquireTokenSilent() assumes it cannot use the cached token, forcing a token refresh.
    This leads to 2 extra requests (OPTIONS and POST) for every single MsGraph request.
*/
const requestScopesSilent = ["email", "openid", "profile"];
const requestScopesLogin = [...requestScopesSilent, "offline_access"];
const extraScopes = [
    "User.Read",                // read user profile, needed to signin?
    "User.ReadBasic.All",       // read basic profile information (name, email) of other users
    "Team.ReadBasic.All",       // Read the names and descriptions of joined teams, on behalf of the signed-in user.
    "TeamMember.Read.All",
    "Channel.ReadBasic.All",    // Read channel names and channel descriptions, on behalf of the signed-in user.
    "ChannelMember.Read.All",
    "ChannelMessage.Read.All",  // Allows an app to read a channel's messages in Microsoft Teams, on behalf of the signed-in user.
    "Chat.Read",                // Read personal 1-to-1 and group chats
];

const consentScopes = [...requestScopesLogin, ...extraScopes];
const expectedScopes = [...requestScopesSilent, ...extraScopes].map(s => s.toLowerCase());

const msalErrors = [
    nameof<msalC.AuthError>(),
    nameof<msal.BrowserAuthError>(),
    nameof<msal.BrowserConfigurationAuthError>(),
    nameof<msalC.ClientAuthError>(),
    nameof<msalC.ClientConfigurationError>(),
    nameof<msalC.ServerError>(),
    nameof<msalC.InteractionRequiredAuthError>(),
];


export const redirectUrlSilent = window.location.origin + "/empty.html";
export const redirectUrlPopup = window.location.origin + "/auth-end.html";



/**
 * Note: can not use instanceOf with stuff extending from builtin types like Error because MSAL in compiled to ES5
 * see https://github.com/Microsoft/TypeScript/wiki/Breaking-Changes#extending-built-ins-like-error-array-and-map-may-no-longer-work
 * @param err
 */
const isProbablyMsalError = (err: unknown): err is msalC.AuthError => {
    return isUnknownObject(err) && typeof err.name === "string" && msalErrors.includes(err.name) && 'errorCode' in err && 'errorMessage' in err;
}


/**
 * Checks, if there is cached account info for the user.
 * If this function return false, full interactive login is necessary
 * @param hint loginHint
 */
export const haveUserInfo = (hint?: string | null): boolean => !!hint && !!client.getAccountByUsername(hint);

/**
 * Get the logged in user's id if available
 * @param hint loginHint
 */
export const userId = (hint?: string | null): string | null => hint ? (client.getAccountByUsername(hint)?.homeAccountId.split(".").shift() ?? null) : null;


export class ScopesNotGranted extends Error {
    constructor(missingScopes: string[]) {
        const message = `Scopes not granted: ${missingScopes.join(' ')}`;
        super(message);
        this.name = nameof(ScopesNotGranted);
        log.error(message);
    }
}


export class AuthError extends Error {
    readonly inner?: unknown;
    readonly isMsalError: boolean;
    readonly isRecoverable: boolean;
    readonly userInteractionRequired: boolean;
    readonly adminConsentRequired: boolean;

    constructor(inner: unknown, messages: IAuthMessages) {
        if (inner instanceof ScopesNotGranted) {
            super(messages.needConsent);
            this.isMsalError = false;
            this.isRecoverable = true;
            this.userInteractionRequired = true;
            this.adminConsentRequired = true;
            AI.trackException({ error: inner });

        } else if (isProbablyMsalError(inner)) {
            if (inner.name === "InteractionRequiredAuthError") {
                super(messages.needServerInteraction);
                this.isMsalError = true;
                this.isRecoverable = true;
                this.userInteractionRequired = true;
                this.adminConsentRequired = false;

            } else {
                switch (inner.errorCode) {
                    case msalC.ClientAuthErrorMessage.accountMismatchError.code:
                    case msalC.ClientAuthErrorMessage.noTokensFoundError.code:
                    case msalC.ClientAuthErrorMessage.userLoginRequiredError.code:
                    case BrowserAuthErrorMessage.popUpWindowError.code:
                    case BrowserAuthErrorMessage.userCancelledError.code:
                        super(messages.loginMessage);
                        this.isMsalError = true;
                        this.isRecoverable = true;
                        this.userInteractionRequired = true;
                        this.adminConsentRequired = false;
                        break;

                    case msalC.ClientAuthErrorMessage.endpointResolutionError.code:
                        super(messages.serverError);
                        this.isMsalError = true;
                        this.isRecoverable = true;
                        this.userInteractionRequired = false;
                        this.adminConsentRequired = false;
                        break;

                    default:
                        super(`${inner.errorCode}: ${inner.errorMessage}`);
                        this.isMsalError = true;
                        this.isRecoverable = false;
                        this.userInteractionRequired = true;
                        this.adminConsentRequired = false;
                }
            }
        } else if (inner === "CancelledByUser" || inner === "FailedToOpenWindow") {
            super(messages.loginMessage);
            this.isMsalError = false;
            this.isRecoverable = true;
            this.userInteractionRequired = true;
            this.adminConsentRequired = false;

        } else if (inner === nameof(AuthError.silentAuthFailed)) {
            super(messages.unkownError);
            this.isMsalError = false;
            this.isRecoverable = true;
            this.userInteractionRequired = true;
            this.adminConsentRequired = false;

        } else {
            super(messages.unkownError);
            this.isMsalError = false;
            this.isRecoverable = false;
            this.userInteractionRequired = false;
            this.adminConsentRequired = false;
        }

        this.name = nameof(AuthError);
        this.inner = inner;
    }


    public static silentAuthFailed(messages: IAuthMessages): never {
        throw new AuthError(nameof(AuthError.silentAuthFailed), messages);
    }
}


/**
 * Verify that all required scopes are granted
 * @param res
 */
function checkScopes(res: msal.AuthenticationResult): void {
    const missing = expectedScopes.filter(es => !res.scopes.includes(es));

    if (missing.length) {
        throw new ScopesNotGranted(missing);
    }
}


/**
 * class is only there to be able to use decorators
 */
class Auth {
    @log.traceAsync()
    static hostedInTeams() {
        return new Promise<boolean>((resolve, reject) => {
            try {
                teams.initialize(() => resolve(true));
                setTimeout(() => resolve(false), 10000);
            } catch (e) {
                reject(e);
            }
        });
    }


    @log.traceAsync()
    static async loginPopup(messages: IAuthMessages, hint: string | null | undefined, needAdminConsent: boolean | undefined): Promise<msal.AuthenticationResult> {
        const params: msal.AuthorizationUrlRequest = { scopes: requestScopesLogin };

        if (hint) {
            params.loginHint = hint;
        }

        if (needAdminConsent) {
            delete params.loginHint;
            params.prompt = "consent";
            params.scopes = consentScopes;
        }

        try {
            let res: msalC.AuthenticationResult;

            if (await hostedInTeams) {
                params.redirectUri = redirectUrlPopup;
                const url = new URL(window.location.origin + '/auth-start.html');
                url.searchParams.append('params', JSON.stringify(params));

                const json = await new Promise<string>((resolve, reject) => teams.authentication.authenticate({
                    url: url.toString(),
                    width: 600,
                    height: 535,
                    successCallback: resolve,
                    failureCallback: reject
                }));

                res = JSON.parse(json);
            } else {
                params.redirectUri = redirectUrlSilent;
                res = await client.acquireTokenPopup(params);
            }

            AI.setAuthenticatedUserContext(res.account.username, res.account.tenantId);

            checkScopes(res);

            if (needAdminConsent && res.account.username !== hint) {
                location.reload();
            }

            return res;
        } catch (err) {
            throw new AuthError(err, messages);
        }
    }


    //@log.traceAsync
    static async getAuthTokenSilent(messages: IAuthMessages, hint: string | null | undefined): Promise<string> {
        // allow silent auth only, if we have a login hint
        if (!hint) {
            AuthError.silentAuthFailed(messages);
        }

        try {
            const account: msalC.AccountInfo | null = client.getAccountByUsername(hint);
            let res: msal.AuthenticationResult | null = null;

            if (account) {
                // try to get the token silently via cache or refresh
                //log.info(`Trying acquireTokenSilent`);
                try {
                    res = await client.acquireTokenSilent({
                        scopes: requestScopesSilent,
                        forceRefresh: false,
                        account,
                        redirectUri: redirectUrlSilent,
                    });
                } catch (err) {
                    const ae = new AuthError(err, messages);
                    if (ae.adminConsentRequired || ae.userInteractionRequired) {
                        throw ae;
                    }
                }
            } else {
                log.info(`No account info cached for user [${hint}]`);
            }

            if (!res) {
                // this will probably succeed as long as the user is accessing the page from inside Teams
                log.info(`Trying ssoSilent`);
                res = await client.ssoSilent({
                    scopes: requestScopesLogin,
                    loginHint: hint,
                    redirectUri: redirectUrlSilent,
                });

                AI.setAuthenticatedUserContext(res.account.username, res.account.tenantId);
            }

            checkScopes(res);

            return res.accessToken;
        } catch (err) {
            throw err instanceof AuthError ? err : new AuthError(err, messages);
        }
    }
}


export const hostedInTeams = Auth.hostedInTeams();
export const logout = (): Promise<void> => client.logout();
export const getAuthTokenSilent = Auth.getAuthTokenSilent;
export const loginPopup = Auth.loginPopup;

