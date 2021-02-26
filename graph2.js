const msalConfig = {
  auth: {
    clientId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
    // comment out if you use a multi-tenant AAD app
    authority: 'https://login.microsoftonline.com/0be187e2-aa5c-464a-bc8b-74b0416b4c3a',
    redirectUri: 'http://localhost:8080'
  }
};
const scopes = ['https://graph.microsoft.com/.default'];

const msalApplication = new Msal.UserAgentApplication(msalConfig);
const options = new MicrosoftGraph.MSALAuthenticationProviderOptions(scopes);
const authProvider = new MicrosoftGraph.ImplicitMSALAuthenticationProvider(msalApplication, options);
const graphClient = MicrosoftGraph.Client.initWithMiddleware({ authProvider });

console.log(msalApplication);
console.log(graphClient);

async function signIn() {
  const authResult = await msalApplication.loginPopup({ scopes });
  sessionStorage.setItem('msalAccount', authResult.account.username);
  // Get the user's profile from Graph
  const user = await getUser();
  displayProfile(user);
}

function loadEmails() {
  if (scopes.indexOf('mail.read') < 0) {
    scopes.push('mail.read');
  }

  // call graph here to request emails
}
