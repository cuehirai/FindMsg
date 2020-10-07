# TODO LIST

- maybe: customize retry logic to handle http 500 errors from MsGraph
- maybe: customize http 429 handling to include more aggressive back-off (especially for responses with retry-after header)
         maybe even cancel sync after a certain amount of 429 responses and tell the user to try again later.
- sync: is there a good way to respond to "410 Gone"?
        Just wait for the next team/channel update, or proactive check if the channel still exists?
        Getting the channel info itself will result it "404 not found"
        https://graph.microsoft.com/beta/teams/a8a8d517-b2a0-4576-a6d7-c326a6d03eab/channels/19:21832bd615be456dbee476cf887c2a5a@thread.skype
        Getting a subresource like messages results in "410 gone"
        https://graph.microsoft.com/beta/teams/a8a8d517-b2a0-4576-a6d7-c326a6d03eab/channels/19:21832bd615be456dbee476cf887c2a5a@thread.skype/messages

## For Kacoms internal release

### Test auth logic for all possible cases

Think of all possible cases!

- Admin/User
- Admin consent given or not
- User consent given or not
- Teams: desktop app/web
- refresh token fresh/expired (after 24h)
- more cases?

### Security/Privacy

#### Messages stored on local PC

The app stores messages and user information in IndexedDb.
This data is protected as much/little as any other browser data.

If a different user gets access to the files, they can extract messages
they would not otherwise have access to.

While it should be clear that Teams is not a good fit for shared terminals, especially when every user uses the same user account, this limitation must be documented.

##### Application Insights

Currently, detailed logs are sent to Application Insights.
Including, but not limited to:

- User ID and name
- Tenant ID
- Team ID and name
- Channel ID and name
- Message IDs


## For public release

### Get some better web-hosting than github pages

(no text)

### Logout button

Review store submission guidelines and implement necessary changes.
The app will probably need an logout function somewhere unobtrusive.

Problem: using MSAL.logout() seems to logout the complete microsoft account, not just this app.

### Onboarding Experience

### First Sync

Should probably force the user to do the first sync in the tab configuration dialog.

### Guest users

Check user guest status and display an error message or something.

Guest users are not supported. Not sure if it is possible to support them.

Reason: The app needs admin consent permissions. It seems, those permissions must be consented in the user's home tenant.
Without these, the user can't login to the app.

It *might* be possible to work around this by asking first for only user account info permissions to make a user status check.
Then, when confirmed that user is not a guest, ask for more permissions

Docs <https://docs.microsoft.com/en-us/microsoftteams/platform/tabs/how-to/access-teams-context> state
that context.userObjectId contains the user id in the "current tenant".
graph GET /users/{id}
should return a user object with 'userType' field set to "Guest" or something else.


## Possible additional features

### User Preferences / state --> DEFER

should the sort order of lists or the search conditions be stored?
automatically or on user request?


### Back button support --> DEFER

Somehow restore the previous state, when the tab is accessed via that back button.
This seems to be not supported at the time of writing.
Maybe microsoftTeams.setFrameContext() will handle this case in the future.
Other users have asked about this.
<https://github.com/OfficeDev/microsoft-teams-library-js/issues/347>
<https://stackoverflow.com/questions/62150589>


## Technical

### Dropdown search IME support is borked --> FIXED in v0.51

In at least chrome and firefox, possibly other browsers as well.
Having Japanese input in would allow searching for users in the dropdown
on the message search tab.

Reported at <https://github.com/microsoft/fluentui/issues/14052> on 2020/7/16.

### Remove `office-ui-fabric-react` dependency

Mainly `@fluentui/react-northstar` is recommended by MS for Teams app UI.
`office-ui-fabric-react` is only added for the `DatePicker` and `Link` components which do not (yet?) exist in `@fluentui/react-northstar`.
Once `DatePicker` and `Link` are available in `@fluentui/react-northstar`,
consider removing `office-ui-fabric-react` to decrease bundle size and loading speed.

UPDATE: there is now a DatePicker in v0.51, but it is marked unstable and not yet documented.

### Search logic --> DEFER

Current implementation is rather basic.

- No indexes are used for filtering and behaviour with many messages is unknown.
- UI might benefit from search being done in a webworker.

Ideally this should have more options to filter based on

- search in subject, main text, or both
- all words, any words (maybe difficult for japanese)

