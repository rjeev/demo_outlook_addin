<template>
  <div id="app">
    <div class="loader" v-if="!isLoaded"></div>

    <div class="content" v-else>
      <div class="content-header" :class="{ blurred: createMode !== '' }">
        <div class="padding">
          <img
            :src="accountData?.photo"
            v-if="accountData !== null"
            class="avatar"
          />
          <p>{{ account.name }},&nbsp;{{ accountData?.designation }}</p>
          <p>Total email logged: {{ totalEmails }}</p>
          <div class="action-buttons-wrapper">
            <button
              class="myBtn action-buttons"
              title="Book Appointment"
              @click="() => (createMode = 'appointment')"
            >
              <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 448 512">
                <path
                  d="M128 0c17.7 0 32 14.3 32 32V64H288V32c0-17.7 14.3-32 32-32s32 14.3 32 32V64h48c26.5 0 48 21.5 48 48v48H0V112C0 85.5 21.5 64 48 64H96V32c0-17.7 14.3-32 32-32zM0 192H448V464c0 26.5-21.5 48-48 48H48c-26.5 0-48-21.5-48-48V192zM329 305c9.4-9.4 9.4-24.6 0-33.9s-24.6-9.4-33.9 0l-95 95-47-47c-9.4-9.4-24.6-9.4-33.9 0s-9.4 24.6 0 33.9l64 64c9.4 9.4 24.6 9.4 33.9 0L329 305z"
                />
              </svg>
            </button>

            <button
              class="myBtn action-buttons"
              title="Send Mail"
              @click="() => (createMode = 'email')"
            >
              <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512">
                <path
                  d="M48 64C21.5 64 0 85.5 0 112c0 15.1 7.1 29.3 19.2 38.4L236.8 313.6c11.4 8.5 27 8.5 38.4 0L492.8 150.4c12.1-9.1 19.2-23.3 19.2-38.4c0-26.5-21.5-48-48-48H48zM0 176V384c0 35.3 28.7 64 64 64H448c35.3 0 64-28.7 64-64V176L294.4 339.2c-22.8 17.1-54 17.1-76.8 0L0 176z"
                />
              </svg>
            </button>
          </div>
        </div>
      </div>
      <div class="content-main">
        <!-- Appointment form -->
        <div class="panel-inner" v-if="createMode === 'appointment'">
          <div class="panel-inner-header">
            <h4>Compose Appointment</h4>
            <button class="close-button" @click="handleCloseModal">X</button>
          </div>
          <!-- Appointment form -->
          <div>
            <label>Appointment Title</label>
            <input
              type="text"
              placeholder="Appointment title"
              @input="updateAppointmentTitle($event.target.value)"
              :value="appointmentTitle"
            />
          </div>
          <div>
            <label>Appointment Date</label>
            <input
              type="date"
              placeholder="Start time"
              @input="updateEventDay($event.target.value)"
              :value="eventDay"
            />
          </div>
          <div>
            <label>Appointment Stgart Time</label>
            <input
              type="time"
              name="Start time"
              id="startTime"
              placeholder="Start Time"
              @input="updateEventStartTime($event.target.value)"
              :value="eventStartTime"
            />
          </div>
          <div>
            <label>Appointment End Time</label>
            <input
              type="time"
              name="End time"
              placeholder="End Time"
              @input="updateEventEndTime($event.target.value)"
              :value="eventEndTime"
            />
          </div>
          <button class="myBtn" @click="assignEvent">Assign Event</button>
        </div>
        <!-- email form -->
        <div class="panel-inner" v-if="createMode === 'email'">
          <div class="panel-inner-header">
            <h4>Compose Email</h4>
            <button class="close-button" @click="handleCloseModal">X</button>
          </div>
          <!-- Send message form -->
          <div>
            <label>Subject of email</label>
            <input
              type="text"
              name="Message"
              placeholder="Message Title"
              @input="updateMessageTitle($event.target.value)"
              :value="messageTitle"
            />
          </div>
          <div>
            <label>Message of email</label>
            <textarea
              name="Message"
              placeholder="Message"
              @input="updateMessage($event.target.value)"
              :value="message"
            />
          </div>
          <button class="myBtn" @click="sendEmail">send</button>
        </div>
        <!-- parsed email contents -->
        <div class="email-content" :class="{ blurred: createMode !== '' }">
          <h2>Email</h2>
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
            <span
              v-for="(recipient, index) in ccRecipients"
              :key="index"
              @click="fetchPersonalData(recipient.emailAddress)"
              class="clickable"
            >
              {{ recipient.emailAddress }}
            </span>
          </div>
          <div v-if="bccRecipients.length > 0">
            <p class="title"><b>BCC Recipients:</b></p>
            <span
              v-for="(recipient, index) in bccRecipients"
              :key="index"
              @click="fetchPersonalData(recipient.emailAddress)"
              class="clickable"
            >
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
        </div>
        <div>
          <button class="myBtn" @click="handleLogEmail">Log mail</button>
        </div>
      </div>
    </div>

    <!-- Popup to display personal data -->
    <div v-if="popupVisible" class="popup">
      <div class="popup-content">
        <span class="close" @click="hidePopup">&times;</span>
        <h2 v-if="personalData">{{ personalData.name }}</h2>
        <div v-if="personalData">
          <img
            :src="personalData.photo"
            alt="Profile Photo"
            class="profile-image"
          />
          <p><b>Email:</b> {{ personalData.email }}</p>
          <p><b>Designation:</b> {{ personalData.designation }}</p>
        </div>
        <div v-else>
          <p>No data available.</p>
        </div>
      </div>
    </div>
    <!-- backdrop for modals -->
    <div
      v-if="createMode !== ''"
      class="backdrop"
      @click="handleCloseModal"
    ></div>
  </div>
</template>

<script>
import {
  PublicClientApplication,
  InteractionRequiredAuthError,
} from "@azure/msal-browser";
import { ref } from "vue";

const testTabs = ref(null);

export default {
  name: "App",

  data() {
    return {
      isLoaded: false,
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
      accountData: null,
      eventDay: null,
      eventStartTime: null,
      eventEndTime: null,
      isEventCreateMode: false,
      message: "",
      messageTitle: "",
      appointmentTitle: "",
      createMode: "",
      totalEmails: null,
    };
  },
  created() {
    this.initializeMsal();
    this.fetchEmailData();
    this.fetchTotalEmails();
  },
  methods: {
    createAppointment() {},
    async initializeMsal() {
      try {
        const msalConfig = {
          auth: {
            clientId: "6821c268-c82f-46be-a889-dc170861f0d8",
            authority: "https://login.microsoftonline.com/common",
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
            that.accessToken = loginResponse.accessToken;
            that.account = loginResponse.account;
            that.fetchPersonalAvatar(that.account.username);
            that.isLoaded = true;
          })
          .catch(function (error) {
            console.log(error);
          });
      } catch (error) {
        console.error("MSAL initialization error:", error);
      }
    },
    async signIn() {
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
    handleCloseModal() {
      this.createMode = "";
    },
    async sendEmail() {
      try {
        const accessToken = this.accessToken;
        const apiUrl = "https://graph.microsoft.com/v1.0/me/sendMail";

        const emailData = {
          message: {
            subject: this.messageTitle,
            body: {
              contentType: "HTML",
              content: this.message,
            },
            toRecipients: [
              {
                emailAddress: {
                  address: this.senderEmail,
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
          this.createMode = "";
          this.$toast.success("Email sent succesfully.");
        } else {
          console.error("Failed to send email:", response.statusText);
          this.$toast.error("Failed to send email:", response.statusText);
        }
      } catch (error) {
        console.error("Error sending email:", error);
        this.$toast.error("Error sending email:", error);
      }
    },
    async assignEvent() {
      const that = this;
      try {
        const apiUrl = "https://graph.microsoft.com/v1.0/me/events";

        const formatDate = (date) => {
          // Pad single digits with leading zero
          return date < 10 ? "0" + date : date;
        };

        const formatDateTime = (date, time) => {
          const [hour, minute] = time.split(":");
          const formattedHour = hour.padStart(2, "0"); // Ensure hour has two digits
          const formattedMinute = minute.padStart(2, "0"); // Ensure minute has two digits
          return `${date.getFullYear()}-${formatDate(
            date.getMonth() + 1
          )}-${formatDate(
            date.getDate()
          )}T${formattedHour}:${formattedMinute}:00`;
        };

        // Construct start and end dateTime
        const startDateTime = formatDateTime(
          new Date(this.eventDay),
          this.eventStartTime
        );
        const endDateTime = formatDateTime(
          new Date(this.eventDay),
          this.eventEndTime
        );

        const allAttendees = [
          { emailAddress: { address: that.senderEmail }, type: "required" },
          ...that.bccRecipients.map((recipient) => ({
            emailAddress: { address: recipient.emailAddress },
            type: "bcc",
          })),
          ...that.ccRecipients.map((recipient) => ({
            emailAddress: { address: recipient.emailAddress },
            type: "cc",
          })),
        ];
        const currentTimezone =
          Intl.DateTimeFormat().resolvedOptions().timeZone;
        console.log(allAttendees, "allAttendees");
        console.log(currentTimezone);
        const eventData = {
          subject: this.appointmentTitle,
          start: {
            dateTime: startDateTime,
            timeZone: currentTimezone,
          },
          end: {
            dateTime: endDateTime,
            timeZone: currentTimezone,
          },
          location: {
            displayName: "Conference Room",
          },
          body: {
            content: "Discuss project progress.",
            contentType: "text",
          },
          attendees: allAttendees.map((attendee) => ({
            emailAddress: {
              address: attendee.emailAddress.address,
            },
            type: "Required",
          })),
        };

        const response = await fetch(apiUrl, {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            Authorization: `Bearer ${this.accessToken}`,
          },
          body: JSON.stringify(eventData),
        });

        if (response.ok) {
          this.createMode = "";
          this.$toast.success("Appointment created succesfully.");
          console.log("Event added to Outlook calendar.");
        } else {
          console.error("Failed to add event:", response.statusText);
          this.$toast.error("Failed to add event:", response.statusText);
        }
      } catch (error) {
        console.error("Error assigning event:", error);
        this.$toast.error("Error assigning event:", error);
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

        await this.fetchAttachments(item.attachments);
        await this.fetchEmailBody();
      } catch (error) {
        console.error("Error fetching email data:", error);
        this.error = "Error fetching email data. Please try again.";
      } finally {
        this.fetching = false;
      }
    },
    updateEventDay(value) {
      this.eventDay = value;
    },
    updateAppointmentTitle(value) {
      this.appointmentTitle = value;
    },
    updateEventStartTime(value) {
      this.eventStartTime = value;
    },
    updateEventEndTime(value) {
      this.eventEndTime = value;
    },
    updateMessage(value) {
      this.message = value;
    },
    updateMessageTitle(value) {
      this.messageTitle = value;
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
            this.$toast.success("Email data successfully logged.");
            this.fetchTotalEmails();
            // this.resetState();
          } else {
            throw new Error("Failed to log email data.");
          }
        })
        .catch((error) => {
          console.error("Error logging email data:", error);
          this.$toast.error("Error logging email data:", error);
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
    async fetchPersonalAvatar(email) {
      // Fetch personal data method with error handling
      try {
        const response = await fetch(
          `http://localhost:3009/personal-data/${email}`
        );
        if (!response.ok) {
          throw new Error("Failed to fetch personal data.");
        }
        const data = await response.json();
        this.accountData = data;
      } catch (error) {
        console.error("Error fetching personal data:", error);
        this.accountData = null;
      }
    },
    async fetchTotalEmails() {
      try {
        const response = await fetch("http://localhost:3009/emails/count");
        if (!response.ok) {
          throw new Error("Failed to fetch total number of emails.");
        }
        const data = await response.json();
        this.totalEmails = data.totalRecords;
      } catch (error) {
        console.error("Error fetching total number of emails:", error);
        // Handle the error as per your application's requirements
      }
    },
    hidePopup() {
      this.popupVisible = false;
    },
    resetState() {
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
  border-radius: 5px;
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
  height: calc(100vh - 392px);
  overflow: auto;
}

.email-content > div {
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

.avatar {
  height: 100px;
  width: 100px;
  border: 2px solid #ffffff;
  border-radius: 50%;
}
.tabs-component-tabs {
  display: block;
  padding: 0;
}

.tabs-component-tab {
  list-style-type: none;
  margin-bottom: 20px;
}
.tabs-component-tab-a {
  display: block;
  text-align: center;
  height: 40px;
  line-height: 40px;
  background-color: var(--secondary-color);
  color: var(--primary-color);
  border: 2px solid var(--primary-color);
  border-radius: 5px;
  font-weight: 500;
  text-decoration: none;
}
.tabs-component-tab-a:hover {
  background: var(--primary-color);
  color: var(--secondary-color);
}
.tabs-component-tab-a.is-active {
  background: var(--primary-color);
  color: var(--secondary-color);
}
.tabs-component-tab:first-child {
  opacity: 0;
  height: 0;
  width: 0;
  margin: 0;
}
.tabs-component-panel {
}
#email-content-pane {
  position: static;
}
.panel-inner {
  background: white;
  margin: 20px;
  padding: 20px;
  border-radius: 4px;
  position: fixed;
  left: 0;
  top: 50%;
  right: 0;
  z-index: 2;
  transform: translateY(-50%);
  box-shadow: rgba(50, 50, 93, 0.25) 0px 50px 100px -20px,
    rgba(0, 0, 0, 0.3) 0px 30px 60px -30px;
}
.panel-inner-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  position: relative;
}
.panel-inner-header h4 {
  margin-top: 0;
}
.action-buttons-wrapper {
  display: flex;
  gap: 10px;
  justify-content: center;
  align-items: center;
}
.action-buttons {
  height: 35px;
  width: 35px;
  display: flex;
  justify-content: center;
  align-items: center;
  margin: 0;
  padding: 0;
  border-radius: 50%;
  background-color: white;
}
.action-buttons svg {
  height: 20px;
  width: 20px;
  fill: var(--primary-color);
}
.action-buttons:hover svg {
  fill: var(--secondary-color);
}
.backdrop {
  position: fixed;
  left: 0;
  top: 0;
  bottom: 0;
  right: 0;
  background-color: rgba(255, 255, 255, 0.746);
}
.blurred {
  filter: blur(5px);
}

/* form styles */
input,
textarea {
  margin-bottom: 20px;
  display: block;
  width: 100%;
}
label {
  font-size: 12px;
  margin-bottom: 5px;
  display: block;
  width: 100%;
  font-weight: 600;
  color: #888888;
}
.close-button {
  appearance: none;
  border: 0;
  padding: 0;
  background: transparent;
  position: absolute;
  right: -8px;
  top: -10px;
  cursor: pointer;
}
.close-button:hover {
  color: var(--primary-color);
}
</style>
