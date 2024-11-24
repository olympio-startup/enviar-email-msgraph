const fetch = require("node-fetch");
const { AUTH_URL, GRAPH_URL } = require("./constants");

class GraphEmailClient {
  constructor(clientId, clientSecret, tenantId, userId) {
    this.clientId = clientId;
    this.clientSecret = clientSecret;
    this.tenantId = tenantId;
    this.userId = userId;
  }

  /**
   * Obtains an access token for the Microsoft Graph API.
   * @returns {Promise<string>} The access token.
   * @throws {Error} If the request fails or the response is not OK.
   */
  async getAccessToken() {
    const url = `${AUTH_URL}/${this.tenantId}/oauth2/v2.0/token`;
    const params = new URLSearchParams();

    params.append("client_id", this.clientId);
    params.append("client_secret", this.clientSecret);
    params.append("scope", `${GRAPH_URL}/.default`);
    params.append("grant_type", "client_credentials");

    try {
      const response = await fetch(url, {
        method: "POST",
        headers: {
          "Content-Type": "application/x-www-form-urlencoded",
        },
        body: params,
      });

      if (!response.ok) {
        throw new Error(
          `Failed to obtain access token: ${await response.text()}`
        );
      }

      const data = await response.json();
      return data.access_token;
    } catch (error) {
      console.error("Error getting access token:", error);
      throw error;
    }
  }

  /**
   * Sends an email using the Microsoft Graph API.
   * @param {string} toEmailAddress - The recipient's email address.
   * @param {string} subject - The subject of the email.
   * @param {string} bodyContent - The HTML content of the email body.
   * @param {Array<Object>} attachments - An array of attachment objects, each with a name and contentBytes.
   * @param {Array<string>} bccAddresses - An array of email addresses to BCC.
   * @param {Array<string>} replyToAddresses - An array of email addresses for the reply-to field.
   * @returns {Promise<void>} A promise that resolves when the email is sent successfully.
   * @throws {Error} If the email fails to send.
   */
  async sendEmail(
    toEmailAddress,
    subject,
    bodyContent,
    attachments,
    bccAddresses,
    replyToAddresses
  ) {
    const accessToken = await this.getAccessToken();
    const url = `${GRAPH_URL}/v1.0/users/${this.userId}/sendMail`;

    const email = {
      message: {
        subject: subject,
        body: {
          contentType: "HTML",
          content: bodyContent,
        },
        toRecipients: [
          {
            emailAddress: {
              address: toEmailAddress,
            },
          },
        ],
        bccRecipients: bccAddresses.map((bcc) => ({
          emailAddress: {
            address: bcc,
          },
        })),
        attachments: attachments.map((attachment) => ({
          "@odata.type": "#microsoft.graph.fileAttachment",
          name: attachment.name,
          contentBytes: attachment.contentBytes,
        })),
        replyTo: replyToAddresses
          ? replyToAddresses.map((replyTo) => ({
              emailAddress: {
                address: replyTo,
              },
            }))
          : [],
      },
    };

    try {
      const response = await fetch(url, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify(email),
      });

      if (!response.ok) {
        throw new Error(`Failed to send email: ${await response.text()}`);
      }

      console.log("Email sent successfully!");
    } catch (error) {
      console.error("Error sending email:", error);
      throw error;
    }
  }

  /**
   * Lists the emails in the user's mailbox.
   * @returns {Promise<Object[]>} An array of email objects.
   * @throws {Error} If the request fails or the response is not OK.
   */
  async listMails() {
    const accessToken = await this.getAccessToken();
    const url = `${GRAPH_URL}/v1.0/users/${this.userId}/messages`;
    try {
      const response = await fetch(url, {
        method: "GET",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
      });

      if (!response.ok) {
        throw new Error(`Failed to list emails: ${await response.text()}`);
      }

      const data = await response.json();
      return data;
    } catch (error) {
      console.error("Error listing emails:", error);
      throw error;
    }
  }

  /**
   * Gets a mail by its ID.
   * @param {string} id The ID of the mail to retrieve.
   * @returns {Promise<Object>} The mail object with the given ID.
   * @throws {Error} If the request fails or the response is not OK.
   */
  async getMailById(id) {
    const accessToken = await this.getAccessToken();
    const url = `${GRAPH_URL}/v1.0/users/${this.userId}/messages/${id}`;
    try {
      const response = await fetch(url, {
        method: "GET",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
      });

      if (!response.ok) {
        throw new Error(`Failed to get email: ${await response.text()}`);
      }

      const data = await response.json();
      return data;
    } catch (error) {
      console.error("Error getting email:", error);
      throw error;
    }
  }

  /**
   * Gets the attachments for a given mail.
   * @param {string} mailId The ID of the mail to get attachments for.
   * @returns {Promise<Object[]>} An array of attachment objects.
   * @throws {Error} If the request fails or the response is not OK.
   */
  async getMailAttachments(mailId) {
    const accessToken = await this.getAccessToken();
    const url = `${GRAPH_URL}/v1.0/users/${this.userId}/messages/${mailId}/attachments`;
    try {
      const response = await fetch(url, {
        method: "GET",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
      });

      if (!response.ok) {
        throw new Error(`Failed to get email attachments: ${await response.text()}`);
      }

      const data = await response.json();
      return data;
    } catch (error) {
      console.error("Error getting email attachments:", error);
      throw error;
    }
  }

  /**
   * Moves a mail to a given folder.
   * @param {string} messageId The ID of the mail to move.
   * @param {string} folderName The name of the folder to move the mail to.
   * @returns {Promise<Object>} The moved mail object.
   * @throws {Error} If the request fails or the response is not OK.
   */
  async moveMailToFolder(messageId, folderName) {
    const accessToken = await this.getAccessToken();
  
    try {
      const folderUrl = `${GRAPH_URL}/v1.0/me/mailFolders`;
      const folderResponse = await fetch(folderUrl, {
        method: "GET",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
      });
  
      if (!folderResponse.ok) {
        throw new Error(`Failed to fetch mail folders: ${await folderResponse.text()}`);
      }
  
      const { value: folders } = await folderResponse.json();
      let folder = folders.find((folder) => folder.displayName === folderName);
  
      if (!folder) {
        console.log(`Folder "${folderName}" not found. Creating it...`);
        const createFolderResponse = await fetch(folderUrl, {
          method: "POST",
          headers: {
            Authorization: `Bearer ${accessToken}`,
            "Content-Type": "application/json",
          },
          body: JSON.stringify({
            displayName: folderName,
          }),
        });
  
        if (!createFolderResponse.ok) {
          throw new Error(`Failed to create folder "${folderName}": ${await createFolderResponse.text()}`);
        }
  
        folder = await createFolderResponse.json();
        console.log(`Folder "${folderName}" created successfully.`);
      }
  
      const folderId = folder.id;
  
      const moveUrl = `https://graph.microsoft.com/v1.0/me/messages/${messageId}/move`;
      const moveResponse = await fetch(moveUrl, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          destinationId: folderId,
        }),
      });
  
      if (!moveResponse.ok) {
        throw new Error(`Failed to move email: ${await moveResponse.text()}`);
      }
  
      const movedMessage = await moveResponse.json();
      console.log(`Email moved successfully to folder "${folderName}".`, movedMessage);
  
      return movedMessage;
    } catch (error) {
      console.error("Error moving email:", error);
      throw error;
    }
  }
  
}

module.exports = GraphEmailClient;
