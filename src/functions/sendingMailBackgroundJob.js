const { app } = require("@azure/functions");
const { ServiceBusClient } = require("@azure/service-bus");
const { ClientSecretCredential } = require("@azure/identity");
const { Client } = require("@microsoft/microsoft-graph-client");
const { MongoClient } = require("mongodb");

const connectionString = process.env["MongoDBConnectionString"];
const client = new MongoClient(connectionString);
const project_tenant_id =
  process.env["PROJECT_TENANT_ID"] ||
  process.env["Azure_user_id"] ||
  "cd7a7f4a-01fc-4a13-9886-73eb7f26a8b9";

app.timer("sendingMailBackgroundJob", {
  schedule: "*/20 * * * * *", // Every 20 seconds
  handler: async (myTimer, context) => {
    await client.connect();
    const database = client.db("ics-cluster-1");
    const collection = database.collection("users");
    const queueName = process.env["QUEUE_NAME"] || "email-queue";
    const serviceBusConnectionString =
      process.env["ServiceBusConnectionString"];
    const serviceBusClient = new ServiceBusClient(serviceBusConnectionString);
    const receiver = serviceBusClient.createReceiver(queueName);

    const tenantId = process.env["TENANT_ID"];
    const clientId = process.env["CLIENT_ID"];
    const clientSecret = process.env["CLIENT_SECRET"];

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

    try {
      const messages = await receiver.receiveMessages(10, {
        maxWaitTimeInMs: 2000,
      });
      let processedCount = 0;
      context.log(`Received ${messages.length} messages from the queue.`);
      for (const message of messages) {
        try {
          const emailData = JSON.parse(message.body);
          context.log(`Processing email for: ${emailData.email}`);

          // Check if the email was sent recently

          // const lastSent = new Date(emailData.lastSent);

          // const now = new Date();

          // const timeDiff = (now - lastSent) / 1000; // time difference in seconds

          // if (timeDiff < 60) {
          //   // Assuming we want to prevent sending emails within 60 seconds

          //   context.log(
          //     `Skipping email to ${emailData.email} as it was sent recently.`
          //   );

          //   await receiver.completeMessage(message);

          //   continue;
          // }

          const email = {
            message: {
              subject: emailData.subject,
              body: {
                contentType: "HTML",
                content: emailData.html,
              },
              toRecipients: [
                {
                  emailAddress: {
                    address: emailData.email,
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

          await graphClient
            .api(`/users/${project_tenant_id}/sendMail`)
            .post(email);
          await receiver.completeMessage(message);
          processedCount++;
          await collection.updateOne(
            { _id: user._id },
            {
              $set: {
                mailSent: true,
                lastSent: new Date(),
                count: user?.count ? user.count + 1 : 1,
              },
            }
          );
          context.log(`Email sent successfully to ${emailData.email}  🎊🎉`);
        } catch (emailError) {
          context.log.error(
            `Error sending email to ${emailData.email}:`,
            emailError
          );
          await receiver.abandonMessage(message);
        }
      }

      return {
        body: `Processed ${processedCount} email messages successfully! 🎊🎉`,
      };
    } catch (error) {
      context.log.error("Error processing email queue:", error);
      return { status: 500, body: "Error processing email queue." };
    } finally {
      await receiver.close();
      await serviceBusClient.close();
      await client.close();
    }
  },
});
