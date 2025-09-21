const { ConfidentialClientApplication } = require('@azure/msal-node');
const { Client } = require('@microsoft/microsoft-graph-client');
const { AuthenticationProvider } = require('@microsoft/microsoft-graph-client');

class GraphAuthProvider {
  constructor(clientId, tenantId) {
    this.clientId = clientId;
    this.tenantId = tenantId;
    this.accessToken = null;

    // MSAL configuration
    this.clientConfig = {
      auth: {
        clientId: clientId,
        authority: `https://login.microsoftonline.com/${tenantId}`,
      }
    };

    this.cca = new ConfidentialClientApplication(this.clientConfig);
  }

  async getAccessToken() {
    try {
      // For desktop applications, we'll use device code flow
      const deviceCodeRequest = {
        scopes: [
          'https://graph.microsoft.com/Mail.ReadWrite',
          'https://graph.microsoft.com/Mail.Send',
          'https://graph.microsoft.com/Calendars.ReadWrite',
          'https://graph.microsoft.com/Contacts.ReadWrite',
          'https://graph.microsoft.com/User.Read'
        ],
        deviceCodeCallback: (response) => {
          console.log('\n=== Microsoft Graph Authentication Required ===');
          console.log(`\nTo authenticate with Microsoft Graph API:`);
          console.log(`1. Open this URL in your browser: ${response.verificationUri}`);
          console.log(`2. Enter this code: ${response.userCode}`);
          console.log(`3. Sign in with your Microsoft account\n`);
        }
      };

      const response = await this.cca.acquireTokenByDeviceCode(deviceCodeRequest);
      this.accessToken = response.accessToken;
      return this.accessToken;
    } catch (error) {
      console.error('Authentication failed:', error);
      throw error;
    }
  }

  async getClient() {
    if (!this.accessToken) {
      await this.getAccessToken();
    }

    const authProvider = {
      getAccessToken: async () => {
        return this.accessToken;
      }
    };

    return Client.initWithMiddleware({ authProvider });
  }
}

module.exports = { GraphAuthProvider };