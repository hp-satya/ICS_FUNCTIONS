const { app } = require("@azure/functions");
const { MongoClient } = require("mongodb");

const connectionString = process.env["MongoDBConnectionString"];
const client = new MongoClient(connectionString);

app.http("getUsers", {
  methods: ["GET"],
  authLevel: "anonymous",
  handler: async (request, context) => {
    try {
      await client.connect();
      const database = client.db("ics-cluster-1");
      const collection = database.collection("users");

      const result = await collection.find({}).toArray();
      // Return the result as JSON
      return { body: JSON.stringify(result) };
    } catch (error) {
      context.error("Error fetching users:", error);
      return { status: 500, body: "Error fetching users" };
    } finally {
      await client.close();
    }
  },
});
