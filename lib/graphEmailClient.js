const fetch = require("node-fetch");

class GraphEmailClient {
  constructor(clientId, clientSecret, tenantId, userId) {
    this.clientId = clientId;
    this.clientSecret = clientSecret;
    this.tenantId = tenantId;
    this.userId = userId;
  }

  // Método para obter o token de acesso
  async getAccessToken() {
    const url = `https://login.microsoftonline.com/${this.tenantId}/oauth2/v2.0/token`;
    const params = new URLSearchParams();

    params.append("client_id", this.clientId);
    params.append("client_secret", this.clientSecret);
    params.append("scope", "https://graph.microsoft.com/.default");
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
          `Failed to obtain access token: ${response.statusText}`
        );
      }

      const data = await response.json();
      return data.access_token;
    } catch (error) {
      console.error("Error getting access token:", error);
      throw error;
    }
  }

  // Método para enviar e-mail
  async sendEmail(toEmailAddress, subject, bodyContent, attachments, bccAddresses) {
    const accessToken = await this.getAccessToken();
    const url = `https://graph.microsoft.com/v1.0/users/${this.userId}/sendMail`;

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
}

module.exports = GraphEmailClient;
