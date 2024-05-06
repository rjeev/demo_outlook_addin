<template>
  <div id="app">
    <!-- <div class="loader" v-if="accessToken === null"></div> -->

    <div class="content">
      <div class="content-header">
        <div class="padding">
          <h1>Welcome</h1>
          <!-- <p>{{ account.name }}</p> -->
        </div>
      </div>
      <div class="content-main">
        <div class="email-content" v-if="subject">
          <div>
            <p class="title"><b>Subject:</b></p>
            {{ subject }}
          </div>
          <div>
            <p class="title"><b>Sender Email:</b></p>
            {{ senderEmail }}
          </div>
          <div>
            <p class="title"><b>Sender Name:</b></p>
            {{ senderName }}
          </div>
          <div v-if="ccRecipients.length > 0">
            <p class="title"><b>CC Recipients:</b></p>
            <span v-for="(recipient, index) in ccRecipients" :key="index"
              @click="fetchPersonalData(recipient.emailAddress)" class="clickable">
              {{ recipient.emailAddress }}
            </span>
          </div>
          <div v-if="bccRecipients.length > 0">
            <p class="title"><b>BCC Recipients:</b></p>
            <span v-for="(recipient, index) in bccRecipients" :key="index"
              @click="fetchPersonalData(recipient.emailAddress)" class="clickable">
              {{ recipient.emailAddress }}
            </span>
          </div>
          <div v-if="attachments.length > 0">
            <p class="title"><b>Attachments:</b></p>
            <div v-for="(attachment, index) in attachments" :key="index">
              <a :href="attachment.url" :download="attachment.name">{{
                attachment.name
                }}</a>
            </div>
          </div>
          <div>
            <b>Body:</b>
            <div class="spacer">&nbsp;</div>
            <div v-html="body" class="email-body"></div>
          </div>
          <div>
            <button class="myBtn" @click="handleLogEmail">Log mail</button>
          </div>
        </div>
        <button class="myBtn" v-else @click="fetchEmailData" :disabled="fetching">
          {{ fetching ? "Fetching..." : "Run" }}
        </button>
        <div>
          <button @click="sendEmail" class="myBtn">send</button>
          <button class="myBtn" @click="assignEvent">Assign Event</button>
        </div>

        <button class="myBtn" @click="summarizeContent" :disabled="fetchingOpenAi">
          {{ fetchingOpenAi ? "summarizing ..." : "Summarize using AI " }}
        </button>

        <div v-if="summarizedContent"><b>Summarized content:</b>
          <div class="summarizedContent">{{ summarizedContent }}</div>
        </div>

        <div v-if="error" class="error">{{ error }}</div>
      </div>
    </div>

    <!-- Popup to display personal data -->
    <div v-if="popupVisible" class="popup">
      <div class="popup-content">
        <span class="close" @click="hidePopup">&times;</span>
        <h2 v-if="personalData">{{ personalData.name }}</h2>
        <div v-if="personalData">
          <img :src="personalData.photo" alt="Profile Photo" class="profile-image" />
          <p><b>Email:</b> {{ personalData.email }}</p>
          <p><b>Designation:</b> {{ personalData.designation }}</p>
        </div>
        <div v-else>
          <p>No data available.</p>
        </div>
      </div>
    </div>
  </div>
</template>

<script>
import {
  PublicClientApplication,
  InteractionRequiredAuthError,
} from "@azure/msal-browser";
import { Configuration, OpenAIApi } from "openai";

// import LoadingSVG from "./assets/loading.svg";

export default {
  name: "App",
  // components: { LoadingSVG },
  data() {
    return {
      subject: "",
      senderEmail: "",
      senderName: "",
      body: "",
      ccRecipients: [],
      bccRecipients: [],
      attachments: [],
      error: null,
      fetching: false,
      emailItem: null,
      popupVisible: false,
      personalData: null,
      accessToken: null,
      msalInstance: null,
      isMsalInitialized: false,
      account: null,
      fetchingOpenAi: false,
      summarizedContent: ""
    };
  },
  created() {
    // this.initializeMsal();
  },
  methods: {
    async initializeMsal() {
      try {
        const msalConfig = {
          auth: {
            clientId: "6821c268-c82f-46be-a889-dc170861f0d8",
            authority:
              "https://login.microsoftonline.com/8cd7b528-f691-4489-b951-fe0d110d54a6",
            redirectUri: "https://localhost:3000",
          },
          system: {
            allowNativeBroker: true,
          },
        };

        const msalInstance = new PublicClientApplication(msalConfig);
        await msalInstance.initialize();
        this.isMsalInitialized = true;
        this.msalInstance = msalInstance;
        this.msalInstance.handleRedirectPromise();
        const that = this;

        await this.msalInstance
          .loginPopup({
            scopes: ["User.ReadWrite"],
          })
          .then(function (loginResponse) {
            console.log(loginResponse, "LoginResponse");
            that.accessToken = loginResponse.accessToken;
            that.account = loginResponse.account;
            // accountId = loginResponse.account.homeAccountId;
            // Display signed-in user content, call API, etc.
            // that.signIn();
          })
          .catch(function (error) {
            //login failure
            console.log(error);
          });
      } catch (error) {
        console.error("MSAL initialization error:", error);
      }
    },
    async summarizeContent() {
      this.summarizedContent = await this.getSelectedText();
    },
    async getSelectedText() {
      this.fetchingOpenAi = true;
      return new window.Office.Promise(function (resolve, reject) {
        try {
          window.Office.context.mailbox.item.body.getAsync(window.Office.CoercionType.Text, async function (asyncResult) {
            const configuration = new Configuration({
              apiKey: process.env.VUE_APP_OPENAI_KEY,
            });
            const openAI = new OpenAIApi(configuration);
            const response = await openAI.createChatCompletion({
              model: "gpt-3.5-turbo",
              messages: [
                {
                  role: "system",
                  content:
                    "You are a helpful assistant that can help users to better manage emails. The following prompt contains the whole mail thread. ",
                },
                {
                  role: "user",
                  content: "Summarize the following mail thread and extract the key points: " + asyncResult.value,
                },
              ],
            });

            resolve(response.data.choices[0].message.content);
            this.fetchingOpenAi = false
          }.bind(this));
        } catch (error) {

          reject(error);
        }
      }.bind(this));
    },
    async signIn() {
      // Check if MSAL instance is fully initialized before attempting sign-in
      if (!this.isMsalInitialized) {
        console.error("MSAL instance is not initialized");
        return;
      }
      const that = this;
      try {
        var request = {
          scopes: [
            "AuditLog.Read.All",
            "Calendars.Read",
            "Calendars.Read.Shared",
            "Calendars.ReadBasic",
            "Calendars.ReadWrite",
            "Calendars.ReadWrite.Shared",
            "Directory.Read.All",
            "email",
            "Mail.Read",
            "Mail.Read.Shared",
            "Mail.ReadBasic",
            "Mail.ReadBasic.Shared",
            "Mail.ReadWrite",
            "Mail.ReadWrite.Shared",
            "Mail.Send",
            "Mail.Send.Shared",
            "MailboxSettings.Read",
            "MailboxSettings.ReadWrite",
            "openid",
            "profile",
            "SecurityEvents.Read.All",
            "SecurityEvents.ReadWrite.All",
            "User.Read",
            "User.ReadWrite",
          ],
          account: that.account,
        };
        this.msalInstance
          .acquireTokenSilent(request)
          .then((tokenResponse) => {
            // Do something with the tokenResponse
            console.log(tokenResponse, "tokenResponse");
            this.accessToken = tokenResponse.accessToken;
          })
          .catch(async (error) => {
            console.log(error);
            if (error instanceof InteractionRequiredAuthError) {
              // Fallback to interactive token acquisition if silent call fails
              return this.msalInstance.acquireTokenSilent(request);
            } else if (error.message.includes("interaction_in_progress")) {
              console.error(
                "An authentication interaction is already in progress."
              );
              // Inform the user that an authentication interaction is in progress
              // Optionally, you can prevent additional sign-in attempts until the current interaction is completed
            } else {
              // Handle other errors
              console.error("Sign-in error:", error);
            }
          });
      } catch (error) {
        console.error("Sign-in error:", error);
      }
    },
    async sendEmail() {
      try {
        const accessToken = this.accessToken; // Assuming you have obtained the access token
        // console.log(accessToken, "this.accessToken");
        const apiUrl = "https://graph.microsoft.com/v1.0/me/sendMail";

        const emailData = {
          message: {
            subject: "Subject of the email",
            body: {
              contentType: "HTML",
              content: "Body of the email",
            },
            toRecipients: [
              {
                emailAddress: {
                  address: "9841pratik@gmail.com",
                },
              },
            ],
          },
        };

        const response = await fetch(apiUrl, {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            Authorization: `Bearer ${accessToken}`,
          },
          body: JSON.stringify(emailData),
        });

        if (response.ok) {
          console.log("Email sent successfully.");
        } else {
          console.error("Failed to send email:", response.statusText);
        }
      } catch (error) {
        console.error("Error sending email:", error);
      }
    },
    async assignEvent() {
      try {
        // Microsoft Graph API endpoint to create an event in the calendar
        const apiUrl = "https://graph.microsoft.com/v1.0/me/events";

        // Data for the event
        const eventData = {
          subject: "Meeting with Client 1",
          start: {
            dateTime: "2024-05-03T10:00:00",
            timeZone: "Pacific Standard Time",
          },
          end: {
            dateTime: "2024-05-06T11:00:00",
            timeZone: "Pacific Standard Time",
          },
          location: {
            displayName: "Conference Room",
          },
          body: {
            content: "Discuss project progress.",
            contentType: "text",
          },
        };

        // Make a POST request to create the event
        const response = await fetch(apiUrl, {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            Authorization: `Bearer ${this.accessToken}`, // Assuming you have obtained the access token
          },
          body: JSON.stringify(eventData),
        });

        if (response.ok) {
          console.log("Event added to Outlook calendar.");
        } else {
          console.error("Failed to add event:", response.statusText);
        }
      } catch (error) {
        console.error("Error assigning event:", error);
      }
    },
    async fetchEmailData() {
      this.error = null;
      this.fetching = true;
      try {
        const item = window.Office.context.mailbox.item;
        this.emailItem = item;
        this.subject = item.subject;
        this.senderEmail = item.from.emailAddress;
        this.senderName = item.from.displayName;
        this.ccRecipients = item.cc || [];
        this.bccRecipients = item.bcc || [];
        const that = this;
        window.Office.context.mailbox.getCallbackTokenAsync(
          { isRest: true },
          function (result) {
            if (result.status === window.Office.AsyncResultStatus.Succeeded) {
              that.accessToken = result.value;
              console.log(that.accessToken, "accessToken");
              // Use the token to authenticate with the remote service
            } else {
              console.error(
                "Failed to retrieve callback token:",
                result.error.message
              );
            }
          }
        );
        await this.fetchAttachments(item.attachments);
        await this.fetchEmailBody();
      } catch (error) {
        console.error("Error fetching email data:", error);
        this.error = "Error fetching email data. Please try again.";
      } finally {
        // console.log(this.emailItem, "item");
        this.fetching = false;
      }
    },
    async fetchAttachments(attachments) {
      await Promise.all(
        attachments.map(async (attachment) => {
          if (attachment.attachmentType === "file") {
            this.attachments.push({
              name: attachment.name,
              url: attachment.content,
            });
          } else {
            try {
              const blobUrl = await this.createBlobUrl(attachment);
              this.attachments.push({
                name: attachment.name,
                url: blobUrl,
              });
            } catch (error) {
              console.error("Error creating blob URL:", error);
              this.error = "Error fetching attachments. Please try again.";
            }
          }
        })
      );
    },
    async createBlobUrl(attachment) {
      const decodedContent = atob(attachment.content);
      const uint8Array = new Uint8Array(decodedContent.length);
      for (let i = 0; i < decodedContent.length; i++) {
        uint8Array[i] = decodedContent.charCodeAt(i);
      }
      const blob = new Blob([uint8Array], { type: attachment.contentType });
      return URL.createObjectURL(blob);
    },
    async fetchEmailBody() {
      return new Promise((resolve, reject) => {
        window.Office.context.mailbox.item.body.getAsync(
          window.Office.CoercionType.Html,
          (bodyResult) => {
            if (
              bodyResult.status === window.Office.AsyncResultStatus.Succeeded
            ) {
              this.body = bodyResult.value;
              resolve();
            } else {
              reject(
                new Error("Error fetching body: " + bodyResult.error.message)
              );
            }
          }
        );
      });
    },
    formatRecipients(recipients) {
      return recipients
        .map(
          (recipient) =>
            recipient.displayName || recipient.emailAddress || "Unknown"
        )
        .join(", ");
    },
    handleLogEmail() {
      const emailData = {
        subject: this.subject,
        senderEmail: this.senderEmail,
        senderName: this.senderName,
        ccRecipients: this.ccRecipients.map(
          (recipient) => recipient.emailAddress
        ),
        bccRecipients: this.bccRecipients.map(
          (recipient) => recipient.emailAddress
        ),
        body: this.body,
      };

      fetch("http://localhost:3009/emails", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(emailData),
      })
        .then((response) => {
          if (response.ok) {
            console.log("Email data successfully logged.");
            this.resetState();
          } else {
            throw new Error("Failed to log email data.");
          }
        })
        .catch((error) => {
          console.error("Error logging email data:", error);
        });
    },
    async fetchPersonalData(email) {
      // Fetch personal data method with error handling
      try {
        const response = await fetch(
          `http://localhost:3009/personal-data/${email}`
        );
        if (!response.ok) {
          throw new Error("Failed to fetch personal data.");
        }
        const data = await response.json();
        this.personalData = data;
        this.popupVisible = true;
      } catch (error) {
        console.error("Error fetching personal data:", error);
        this.personalData = null;
        this.popupVisible = false; // Hide the popup if an error occurs
      }
    },
    hidePopup() {
      // Hide popup method
      this.popupVisible = false;
    },
    resetState() {
      // Reset state method
      this.subject = "";
      this.senderEmail = "";
      this.senderName = "";
      this.body = "";
      this.ccRecipients = [];
      this.bccRecipients = [];
      this.attachments = [];
      this.error = null;
      this.fetching = false;
      this.emailItem = null;
      this.accessToken = null; // Clear the access token
      this.fetchingOpenAi = false;
      this.summarizedContent = "";
    },

    async bookAppointment() {
      try {
        const item = window.Office.context.calendar.getAppointmentForm();
        item.start.setHours(9, 0, 0, 0);
        item.end.setHours(10, 0, 0, 0);
        item.subject = "Appointment Subject";
        item.location = "Appointment Location";
        item.requiredAttendees.addEmailAddress("recipient@example.com");
        item.saveAsync((result) => {
          if (result.error) {
            console.error("Error booking appointment:", result.error);
          } else {
            console.log("Appointment booked successfully.");
          }
        });
      } catch (error) {
        console.error("Error booking appointment:", error);
      }
    },
    async sendEmailOrBookAppointment() {
      try {
        const isEmail = true; // Change to false for booking an appointment

        if (isEmail) {
          await this.sendEmail();
        } else {
          await this.bookAppointment();
        }
      } catch (error) {
        console.error("Error sending email or booking appointment:", error);
      }
    },
  },
};
</script>

<style>
:root {
  --primary-color: #2a8dd4;
  --secondary-color: #fff;
  --accent-color: #f00;
  --button-bg-color: var(--primary-color);
  --button-hover-color: #1d6fa5;
  --button-disabled-color: #ccc;
}

#app {
  font-family: Arial, sans-serif;
}

.content-header {
  background-color: var(--primary-color);
  color: var(--secondary-color);
  padding: 15px;
  text-align: center;
}

.content-main {
  padding-top: 20px;
}

.content {
  max-width: 800px;
  margin: 0 auto;
}

.email-content {
  background-color: var(--secondary-color);
  border: 1px solid #ddd;
  border-radius: 5px;
  padding: 15px;
  margin-bottom: 20px;
}

.email-content>div {
  margin-bottom: 20px;
}

.email-content div b {
  color: var(--primary-color);
}

.email-content span {
  display: block;
  margin-top: 10px;
}

.myBtn {
  padding: 10px 20px;
  background-color: var(--button-bg-color);
  color: var(--secondary-color);
  border: none;
  border-radius: 5px;
  margin: 20px auto;
  display: block;
  cursor: pointer;
  transition: background-color 0.3s;
}

.myBtn:disabled {
  background-color: var(--button-disabled-color);
  cursor: not-allowed;
}

.myBtn:hover {
  background-color: var(--button-hover-color);
}

.error {
  color: var(--accent-color);
  margin-top: 10px;
  text-align: center;
}

.email-body {
  width: 100%;
  overflow: auto;
}

.title {
  margin: 0 0 5px;
}

.popup {
  position: fixed;
  top: 50%;
  left: 50%;
  transform: translate(-50%, -50%);
  background-color: rgba(255, 255, 255, 0.9);
  border-radius: 10px;
  padding: 20px;
  box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
  z-index: 9999;
}

.popup-content {
  max-width: 300px;
  text-align: center;
  position: relative;
}

.close {
  position: absolute;
  top: 10px;
  right: 10px;
  cursor: pointer;
  transform: translate3d(16px, -34px, 10px);
}

.profile-image {
  width: 120px;
  height: 120px;
  border-radius: 50%;
  margin: 0 auto 10px;
}

.clickable {
  cursor: pointer;
  color: blue;
  text-decoration: underline;
}

.clickable:hover {
  color: darkblue;
}

.loader {
  height: 100vh;
  width: 100%;
  background: url('data:image/svg+xml,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 200"><circle fill="none" stroke-opacity="1" stroke="%2311EBFF" stroke-width=".5" cx="100" cy="100" r="0"><animate attributeName="r" calcMode="spline" dur="2" values="1;80" keyTimes="0;1" keySplines="0 .2 .5 1" repeatCount="indefinite"></animate><animate attributeName="stroke-width" calcMode="spline" dur="2" values="0;25" keyTimes="0;1" keySplines="0 .2 .5 1" repeatCount="indefinite"></animate><animate attributeName="stroke-opacity" calcMode="spline" dur="2" values="1;0" keyTimes="0;1" keySplines="0 .2 .5 1" repeatCount="indefinite"></animate></circle></svg>');
  background-position: center;
  background-repeat: no-repeat;
  background-size: 100px 100px;
}

.summarizedContent {
  white-space: pre-wrap;
}
</style>
