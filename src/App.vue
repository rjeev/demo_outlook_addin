<template>
  <div id="app">
    <div class="content">
      <div class="content-header">
        <div class="padding">
          <h1>Welcome</h1>
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
          <div>
            <button class="myBtn" @click="handleLogEmail">Log mail</button>
          </div>
        </div>
        <button
          class="myBtn"
          v-else
          @click="fetchEmailData"
          :disabled="fetching"
        >
          {{ fetching ? "Fetching..." : "Run" }}
        </button>
        <div v-if="error" class="error">{{ error }}</div>
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
  </div>
</template>

<script>
export default {
  name: "App",
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
    };
  },
  methods: {
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
        this.popupVisible = true; // Show the popup even if no data is available
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
</style>
