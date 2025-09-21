class OutlookAPI {
  constructor(graphClient) {
    this.graphClient = graphClient;
  }

  async sendEmail(to, subject, body, isHtml = false) {
    try {
      const message = {
        subject: subject,
        body: {
          contentType: isHtml ? 'HTML' : 'Text',
          content: body
        },
        toRecipients: Array.isArray(to)
          ? to.map(email => ({ emailAddress: { address: email } }))
          : [{ emailAddress: { address: to } }]
      };

      await this.graphClient.api('/me/sendMail').post({ message });
      return { success: true, message: 'Email sent successfully' };
    } catch (error) {
      console.error('Send email error:', error);
      return { success: false, error: error.message };
    }
  }

  async readEmails(folderId = 'inbox', top = 10, search = null) {
    try {
      let query = this.graphClient.api(`/me/mailFolders/${folderId}/messages`)
        .top(top)
        .select('id,subject,from,receivedDateTime,body,isRead,importance')
        .orderby('receivedDateTime desc');

      if (search) {
        query = query.search(search);
      }

      const messages = await query.get();

      return {
        success: true,
        emails: messages.value.map(email => ({
          id: email.id,
          subject: email.subject,
          from: email.from?.emailAddress?.address || 'Unknown',
          fromName: email.from?.emailAddress?.name || 'Unknown',
          receivedDateTime: email.receivedDateTime,
          body: email.body?.content || '',
          isRead: email.isRead,
          importance: email.importance
        }))
      };
    } catch (error) {
      console.error('Read emails error:', error);
      return { success: false, error: error.message };
    }
  }

  async createCalendarEvent(subject, start, end, attendees = [], body = '', location = '') {
    try {
      const event = {
        subject: subject,
        body: {
          contentType: 'HTML',
          content: body
        },
        start: {
          dateTime: start,
          timeZone: 'UTC'
        },
        end: {
          dateTime: end,
          timeZone: 'UTC'
        },
        location: {
          displayName: location
        },
        attendees: attendees.map(email => ({
          emailAddress: {
            address: email,
            name: email
          }
        }))
      };

      const createdEvent = await this.graphClient.api('/me/events').post(event);

      return {
        success: true,
        event: {
          id: createdEvent.id,
          subject: createdEvent.subject,
          start: createdEvent.start.dateTime,
          end: createdEvent.end.dateTime,
          webLink: createdEvent.webLink
        }
      };
    } catch (error) {
      console.error('Create calendar event error:', error);
      return { success: false, error: error.message };
    }
  }

  async getCalendarEvents(startDate, endDate, top = 25) {
    try {
      const events = await this.graphClient
        .api('/me/events')
        .filter(`start/dateTime ge '${startDate}' and end/dateTime le '${endDate}'`)
        .select('id,subject,start,end,location,attendees,organizer')
        .orderby('start/dateTime')
        .top(top)
        .get();

      return {
        success: true,
        events: events.value.map(event => ({
          id: event.id,
          subject: event.subject,
          start: event.start.dateTime,
          end: event.end.dateTime,
          location: event.location?.displayName || '',
          organizer: event.organizer?.emailAddress?.address || '',
          attendees: event.attendees?.map(a => a.emailAddress?.address) || []
        }))
      };
    } catch (error) {
      console.error('Get calendar events error:', error);
      return { success: false, error: error.message };
    }
  }

  async searchContacts(searchTerm, top = 10) {
    try {
      const contacts = await this.graphClient
        .api('/me/contacts')
        .search(searchTerm)
        .select('id,displayName,emailAddresses,businessPhones,mobilePhone,jobTitle,companyName')
        .top(top)
        .get();

      return {
        success: true,
        contacts: contacts.value.map(contact => ({
          id: contact.id,
          displayName: contact.displayName,
          emailAddresses: contact.emailAddresses?.map(e => e.address) || [],
          businessPhones: contact.businessPhones || [],
          mobilePhone: contact.mobilePhone,
          jobTitle: contact.jobTitle,
          companyName: contact.companyName
        }))
      };
    } catch (error) {
      console.error('Search contacts error:', error);
      return { success: false, error: error.message };
    }
  }

  async createContact(displayName, email, phone = '', company = '', jobTitle = '') {
    try {
      const contact = {
        displayName: displayName,
        emailAddresses: email ? [{ address: email }] : [],
        businessPhones: phone ? [phone] : [],
        companyName: company,
        jobTitle: jobTitle
      };

      const createdContact = await this.graphClient.api('/me/contacts').post(contact);

      return {
        success: true,
        contact: {
          id: createdContact.id,
          displayName: createdContact.displayName,
          emailAddresses: createdContact.emailAddresses?.map(e => e.address) || []
        }
      };
    } catch (error) {
      console.error('Create contact error:', error);
      return { success: false, error: error.message };
    }
  }
}

module.exports = { OutlookAPI };