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
            {{ formatRecipients(ccRecipients) }}
          </div>
          <div v-if="bccRecipients.length > 0">
            <p class="title"><b>BCC Recipients:</b></p>
            {{ formatRecipients(bccRecipients) }}
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
        await this.fetchAttachments(item.attachments);
        await this.fetchEmailBody();
      } catch (error) {
        console.error("Error fetching email data:", error);
        this.error = "Error fetching email data. Please try again.";
      } finally {
        this.fetching = false;
      }
    },
    handleLogEmail() {
      console.log(this.emailItem);
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
</style>
