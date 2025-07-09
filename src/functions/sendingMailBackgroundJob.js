const { app } = require("@azure/functions");
const { ServiceBusClient } = require("@azure/service-bus");
const { ClientSecretCredential } = require("@azure/identity");
const { Client } = require("@microsoft/microsoft-graph-client");

app.timer("sendingMailBackgroundJob", {
  schedule: "* */5 * * * *", // Every 5 minutes
  handler: async (myTimer, context) => {
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
        maxWaitTimeInMs: 5000,
      });
      let processedCount = 0;
      context.log(`Received ${messages.length} messages from the queue.`);
      for (const message of messages) {
        try {
          const emailData = JSON.parse(message.body);
          context.log(`Processing email for: ${emailData.email}`);

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
            .api(`/users/cd7a7f4a-01fc-4a13-9886-73eb7f26a8b9/sendMail`)
            .post(email);
          await receiver.completeMessage(message);
          processedCount++;
          context.log(`Email sent successfully to ${emailData.email}`);
        } catch (emailError) {
          context.log.error(
            `Error sending email to ${emailData.email}:`,
            emailError
          );
          await receiver.abandonMessage(message); // Consider implementing a retry mechanism
        }
      }

      return {
        body: `Processed ${processedCount} email messages successfully!`,
      };
    } catch (error) {
      context.log.error("Error processing email queue:", error);
      return { status: 500, body: "Error processing email queue." };
    } finally {
      await receiver.close();
      await serviceBusClient.close();
    }
  },
});
