const { app } = require("@azure/functions");
const { MongoClient, ObjectId } = require("mongodb");
const { Client } = require("@microsoft/microsoft-graph-client");
const { ClientSecretCredential } = require("@azure/identity");

const connectionString = process.env["MongoDBConnectionString"];
const client = new MongoClient(connectionString);
const tenantId = process.env["TENANT_ID"];
const clientId = process.env["CLIENT_ID"];
const clientSecret = process.env["CLIENT_SECRET"];
const Azure_user_id = process.env["Azure_user_id"];

const htmlContent = `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Thank You for Visiting Henny Penny</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            background-color: #f4f4f4;
            margin: 0;
            padding: 20px;
        }
        .container {
            background-color: #ffffff;
            padding: 20px;
            border-radius: 5px;
            box-shadow: 0 2px 5px rgba(0, 0, 0, 0.1);
        }
        h1 {
            color: #333;
        }
        p {
            color: #555;
        }
        .footer {
            margin-top: 20px;
            font-size: 12px;
            color: #999;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>Thank You for Visiting Henny Penny! 🎉</h1>
        <p>We are so grateful for your visit and hope you had a wonderful experience! 😊</p>
        <p>Your support means the world to us, and we look forward to welcoming you back soon!</p>
        <p>Best Regards,<br>The Henny Penny Team</p>
    </div>
    <div class="footer">
        <p>&copy; 2023 Henny Penny. All rights reserved.</p>
    </div>
</body>
</html>`;

app.http("sendEmail", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "sendEmail",
  handler: async (request, context) => {
    try {
      await client.connect();
      const database = client.db("ics-cluster-1");
      const collection = database.collection("users");

      const { id, subject, html } = await request.json();

      if (!id) {
        return { status: 400, body: "User ID is required in the body." };
      }

      const userObj = await collection.findOne({ _id: new ObjectId(id) });
      const user = JSON.parse(JSON.stringify(userObj));

      if (!user || !user.email) {
        return { status: 404, body: "User not found or email is missing." };
      }

      const credential = new ClientSecretCredential(
        tenantId,
        clientId,
        clientSecret
      );
      const graphClient = Client.initWithMiddleware({
        authProvider: {
          getAccessToken: async () => {
            const tokenResponse = await credential.getToken(
              "https://graph.microsoft.com/.default"
            );
            return tokenResponse.token;
          },
        },
      });

      const email = {
        message: {
          subject: subject ? subject : "Test Email from Microsoft Graph",
          body: {
            contentType: "HTML",
            content: htmlContent
              ? htmlContent
              : "<p>This is a test email sent using Microsoft Graph API.</p>",
          },
          toRecipients: [
            {
              emailAddress: {
                address: user?.email,
              },
            },
          ],
          from: {
            emailAddress: {
              address: "Ics-reports@hennypenny.com",
            },
          },
        },
      };

      await graphClient.api(`/users/${Azure_user_id}/sendMail`).post(email);
      return {
        body: `Email sent to ${user?.email} successfully!`,
      };
    } catch (error) {
      context.error("Error sending email:", error);
      return { status: 500, body: "Error sending email." };
    } finally {
      await client.close();
    }
  },
});
