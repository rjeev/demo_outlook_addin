const express = require("express");
const { Pool } = require("pg");
const cors = require("cors"); // Import the cors middleware

const app = express();
const port = 3009;

// Create a pool to handle database connections
const pool = new Pool({
  user: "postgres",
  host: "localhost",
  database: "emaildump",
  password: "lfkoP@ssw0rd",
  port: 5432,
});

// Middleware to parse JSON request bodies
app.use(express.json());

// Enable CORS for all routes
app.use(cors());

app.post("/emails", async (req, res) => {
  const {
    subject,
    senderEmail,
    senderName,
    ccRecipients,
    bccRecipients,
    body,
  } = req.body;

  try {
    const client = await pool.connect();
    const queryText =
      "INSERT INTO emails (subject, sender_email, sender_name, cc_recipients, bcc_recipients, body) VALUES ($1, $2, $3, $4, $5, $6)";
    const ccRecipientsArray = ccRecipients
      .map((recipient) => `{${recipient}}`)
      .join(",");
    const bccRecipientsArray = bccRecipients
      .map((recipient) => `{${recipient}}`)
      .join(",");
    const values = [
      subject,
      senderEmail,
      senderName,
      `{${ccRecipientsArray}}`,
      `{${bccRecipientsArray}}`,
      body,
    ];
    await client.query(queryText, values);
    client.release();
    res.status(201).send("Email information stored successfully.");
  } catch (error) {
    console.error("Error storing email information:", error); // Log the error

    // Send a more detailed error message to the client
    res.status(500).send(`Internal Server Error: ${error.message}`);
  }
});

app.listen(port, () => {
  console.log(`Server is listening at http://localhost:${port}`);
});
