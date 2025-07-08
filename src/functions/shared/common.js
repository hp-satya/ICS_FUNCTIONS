const { ServiceBusClient } = require("@azure/service-bus");
const { ClientSecretCredential } = require("@azure/identity");
const { Client } = require("@microsoft/microsoft-graph-client");

// Add this shared function at the top of your file (outside of app.http)
const processEmailMessages = async (context) => {
  context.log("Processing email messages from Service Bus queue...");
  const queueName = process.env["QUEUE_NAME"] || "email-queue";
  const serviceBusConnectionString = process.env["ServiceBusConnectionString"];
  const serviceBusClient = new ServiceBusClient(serviceBusConnectionString);
  const receiver = serviceBusClient.createReceiver(queueName);

  const tenantId = process.env["TENANT_ID"];
  const clientId = process.env["CLIENT_ID"];
  const clientSecret = process.env["CLIENT_SECRET"];
  const Azure_user_id = process.env["Azure_user_id"]; // Replace with your Azure AD user ID

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

  const messages = await receiver.receiveMessages(2, {
    maxWaitTimeInMs: 1000,
  });
  let processedCount = 0;

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

      await graphClient.api(`/users/${Azure_user_id}/sendMail`).post(email);

      await receiver.completeMessage(message);
      processedCount++;
      context.log(`Email sent successfully to ${emailData.email}`);
    } catch (emailError) {
      context.log.error("Error sending email:", emailError);
      await receiver.abandonMessage(message);
    }
  }

  await receiver.close();
  await serviceBusClient.close();

  return processedCount;
};

module.exports = { processEmailMessages };
