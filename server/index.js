#!/usr/bin/env node

const { Server } = require('@modelcontextprotocol/sdk/server/index.js');
const { StdioServerTransport } = require('@modelcontextprotocol/sdk/server/stdio.js');
const { CallToolRequestSchema, ListToolsRequestSchema } = require('@modelcontextprotocol/sdk/types.js');
const { GraphAuthProvider } = require('./auth.js');
const { OutlookAPI } = require('./outlook-api.js');

class OutlookMCPServer {
  constructor() {
    this.server = new Server(
      {
        name: 'outlook-desktop-extension',
        version: '1.0.0',
      },
      {
        capabilities: {
          tools: {},
        },
      }
    );

    this.authProvider = null;
    this.outlookAPI = null;
    this.setupToolHandlers();
  }

  async initializeAuth() {
    if (!this.authProvider) {
      // First try command line arguments (similar to Softeria implementation)
      const args = process.argv;
      let clientId = null;
      let tenantId = null;

      for (let i = 0; i < args.length; i++) {
        if (args[i] === '--client-id' && i + 1 < args.length) {
          clientId = args[i + 1];
        }
        if (args[i] === '--tenant-id' && i + 1 < args.length) {
          tenantId = args[i + 1];
        }
      }

      // Then try environment variables (various formats Claude Desktop might use)
      if (!clientId) {
        clientId = process.env.CLIENT_ID ||
                  process.env.client_id ||
                  process.env.MCP_CLIENT_ID ||
                  process.env.USER_CONFIG_CLIENT_ID ||
                  process.env.MS365_MCP_CLIENT_ID;
      }

      if (!tenantId) {
        tenantId = process.env.TENANT_ID ||
                  process.env.tenant_id ||
                  process.env.MCP_TENANT_ID ||
                  process.env.USER_CONFIG_TENANT_ID ||
                  process.env.MS365_MCP_TENANT_ID;
      }

      // Fallback to default values for testing
      if (!clientId) {
        clientId = 'your-client-id-here';
      }
      if (!tenantId) {
        tenantId = 'your-tenant-id-here';
      }

      console.error('=== Outlook Extension Debug Info ===');
      console.error('CLIENT_ID:', clientId ? (clientId.length > 10 ? clientId.substring(0, 10) + '...' : clientId) : 'Missing');
      console.error('TENANT_ID:', tenantId ? (tenantId.length > 10 ? tenantId.substring(0, 10) + '...' : tenantId) : 'Missing');
      console.error('Command line args:', args.slice(2));
      console.error('Env vars with CLIENT/TENANT:', Object.keys(process.env).filter(k => k.toUpperCase().includes('CLIENT') || k.toUpperCase().includes('TENANT')));

      if (!clientId || clientId === 'your-client-id-here' || !tenantId || tenantId === 'your-tenant-id-here') {
        throw new Error(`
=== CONFIGURATION REQUIRED ===
Please configure your Azure App Registration details:
1. CLIENT_ID: ${clientId === 'your-client-id-here' ? 'NOT SET' : 'Set'}
2. TENANT_ID: ${tenantId === 'your-tenant-id-here' ? 'NOT SET' : 'Set'}

Add these to Claude Desktop extension configuration or set as environment variables.
        `);
      }

      this.authProvider = new GraphAuthProvider(clientId, tenantId);
      const graphClient = await this.authProvider.getClient();
      this.outlookAPI = new OutlookAPI(graphClient);
    }
  }

  setupToolHandlers() {
    this.server.setRequestHandler(ListToolsRequestSchema, async () => {
      return {
        tools: [
          {
            name: 'send_email',
            description: 'Send an email through Outlook',
            inputSchema: {
              type: 'object',
              properties: {
                to: {
                  type: ['string', 'array'],
                  description: 'Recipient email address(es)',
                },
                subject: {
                  type: 'string',
                  description: 'Email subject',
                },
                body: {
                  type: 'string',
                  description: 'Email body content',
                },
                isHtml: {
                  type: 'boolean',
                  description: 'Whether the body is HTML formatted',
                  default: false,
                },
              },
              required: ['to', 'subject', 'body'],
            },
          },
          {
            name: 'read_emails',
            description: 'Read and search emails from Outlook',
            inputSchema: {
              type: 'object',
              properties: {
                folderId: {
                  type: 'string',
                  description: 'Folder to read from (inbox, sent, etc.)',
                  default: 'inbox',
                },
                top: {
                  type: 'number',
                  description: 'Number of emails to retrieve',
                  default: 10,
                },
                search: {
                  type: 'string',
                  description: 'Search query to filter emails',
                },
              },
            },
          },
          {
            name: 'create_calendar_event',
            description: 'Create a new calendar event in Outlook',
            inputSchema: {
              type: 'object',
              properties: {
                subject: {
                  type: 'string',
                  description: 'Event title',
                },
                start: {
                  type: 'string',
                  description: 'Start date/time (ISO 8601 format)',
                },
                end: {
                  type: 'string',
                  description: 'End date/time (ISO 8601 format)',
                },
                attendees: {
                  type: 'array',
                  items: { type: 'string' },
                  description: 'List of attendee email addresses',
                  default: [],
                },
                body: {
                  type: 'string',
                  description: 'Event description',
                  default: '',
                },
                location: {
                  type: 'string',
                  description: 'Event location',
                  default: '',
                },
              },
              required: ['subject', 'start', 'end'],
            },
          },
          {
            name: 'get_calendar_events',
            description: 'Retrieve calendar events from Outlook',
            inputSchema: {
              type: 'object',
              properties: {
                startDate: {
                  type: 'string',
                  description: 'Start date for event range (ISO 8601 format)',
                },
                endDate: {
                  type: 'string',
                  description: 'End date for event range (ISO 8601 format)',
                },
                top: {
                  type: 'number',
                  description: 'Maximum number of events to retrieve',
                  default: 25,
                },
              },
              required: ['startDate', 'endDate'],
            },
          },
          {
            name: 'search_contacts',
            description: 'Search for contacts in Outlook',
            inputSchema: {
              type: 'object',
              properties: {
                searchTerm: {
                  type: 'string',
                  description: 'Search term to find contacts',
                },
                top: {
                  type: 'number',
                  description: 'Maximum number of contacts to return',
                  default: 10,
                },
              },
              required: ['searchTerm'],
            },
          },
          {
            name: 'create_contact',
            description: 'Create a new contact in Outlook',
            inputSchema: {
              type: 'object',
              properties: {
                displayName: {
                  type: 'string',
                  description: 'Contact display name',
                },
                email: {
                  type: 'string',
                  description: 'Contact email address',
                },
                phone: {
                  type: 'string',
                  description: 'Contact phone number',
                  default: '',
                },
                company: {
                  type: 'string',
                  description: 'Contact company name',
                  default: '',
                },
                jobTitle: {
                  type: 'string',
                  description: 'Contact job title',
                  default: '',
                },
              },
              required: ['displayName'],
            },
          },
        ],
      };
    });

    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      await this.initializeAuth();

      const { name, arguments: args } = request.params;

      try {
        switch (name) {
          case 'send_email':
            const emailResult = await this.outlookAPI.sendEmail(
              args.to,
              args.subject,
              args.body,
              args.isHtml
            );
            return {
              content: [
                {
                  type: 'text',
                  text: JSON.stringify(emailResult, null, 2),
                },
              ],
            };

          case 'read_emails':
            const readResult = await this.outlookAPI.readEmails(
              args.folderId,
              args.top,
              args.search
            );
            return {
              content: [
                {
                  type: 'text',
                  text: JSON.stringify(readResult, null, 2),
                },
              ],
            };

          case 'create_calendar_event':
            const eventResult = await this.outlookAPI.createCalendarEvent(
              args.subject,
              args.start,
              args.end,
              args.attendees,
              args.body,
              args.location
            );
            return {
              content: [
                {
                  type: 'text',
                  text: JSON.stringify(eventResult, null, 2),
                },
              ],
            };

          case 'get_calendar_events':
            const calendarResult = await this.outlookAPI.getCalendarEvents(
              args.startDate,
              args.endDate,
              args.top
            );
            return {
              content: [
                {
                  type: 'text',
                  text: JSON.stringify(calendarResult, null, 2),
                },
              ],
            };

          case 'search_contacts':
            const searchResult = await this.outlookAPI.searchContacts(
              args.searchTerm,
              args.top
            );
            return {
              content: [
                {
                  type: 'text',
                  text: JSON.stringify(searchResult, null, 2),
                },
              ],
            };

          case 'create_contact':
            const contactResult = await this.outlookAPI.createContact(
              args.displayName,
              args.email,
              args.phone,
              args.company,
              args.jobTitle
            );
            return {
              content: [
                {
                  type: 'text',
                  text: JSON.stringify(contactResult, null, 2),
                },
              ],
            };

          default:
            throw new Error(`Unknown tool: ${name}`);
        }
      } catch (error) {
        console.error(`Error executing tool ${name}:`, error);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ success: false, error: error.message }, null, 2),
            },
          ],
        };
      }
    });
  }

  async run() {
    const transport = new StdioServerTransport();
    await this.server.connect(transport);
    console.error('Outlook MCP Server running on stdio');
  }
}

const server = new OutlookMCPServer();
server.run().catch(console.error);