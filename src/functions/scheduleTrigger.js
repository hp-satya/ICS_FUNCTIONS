const { app } = require("@azure/functions");
const { ServiceBusClient } = require("@azure/service-bus");
const { processEmailMessages } = require("./shared/common");
const { MongoClient } = require("mongodb");

const connectionString = process.env["MongoDBConnectionString"];
const client = new MongoClient(connectionString);
const queueName = process.env["QUEUE_NAME"] || "email-queue";
const serviceBusConnectionString = process.env["ServiceBusConnectionString"];
const html = `<!DOCTYPE html>
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

app.timer("scheduleTrigger", {
  schedule: "0 */5 * * * *",
  handler: async (myTimer, context) => {
    await client.connect();
    const database = client.db("ics-cluster-1");

    const currentTime = new Date();
    const currentHours = currentTime.getUTCHours().toString().padStart(2, "0");
    const currentMinutes = currentTime
      .getUTCMinutes()
      .toString()
      .padStart(2, "0");
    const currentFormattedTime = `${currentHours}:${currentMinutes}`;

    try {
      const collection = database.collection("users");
      const serviceBusClient = new ServiceBusClient(serviceBusConnectionString);

      const sender = serviceBusClient.createSender(queueName);

      collection.forEach(async (user) => {
        if (user?.scheduled) {
          const scheduledTime = new Date(user?.scheduled);
          const scheduledHours = scheduledTime
            .getUTCHours()
            .toString()
            .padStart(2, "0");
          const scheduledMinutes = scheduledTime
            .getUTCMinutes()
            .toString()
            .padStart(2, "0");
          const scheduledFormattedTime = `${scheduledHours}:${scheduledMinutes}`;

          if (scheduledFormattedTime === currentFormattedTime) {
            context.log(
              `Triggering action for user: ${user.id} at scheduled time: ${scheduledFormattedTime}`
            );

            const message = {
              body: JSON.stringify({
                userId: id,
                email: user.email,
                subject: "Test Email from Microsoft Graph",
                html: html,
                timestamp: new Date().toISOString(),
                messageId: `email-${Date.now()}-${Math.random()
                  .toString(36)
                  .substr(2, 9)}`,
              }),
              contentType: "application/json",
            };
            await sender.sendMessages(message);
            const processedCount = await processEmailMessages(context);
            context.log(`Processed ${processedCount} emails immediately`);
          }
        }
      });
      await sender.close();
      await serviceBusClient.close();
      return {
        body: `Email queued and emails processed successfully!`,
      };
    } catch (error) {
      context.log.error("Error querying Cosmos DB:", error);
    }
  },
});
