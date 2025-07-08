const { app } = require("@azure/functions");
const { MongoClient, ObjectId } = require("mongodb");

const connectionString = process.env["MongoDBConnectionString"];
const client = new MongoClient(connectionString);

app.http("getUser", {
  methods: ["GET"],
  authLevel: "anonymous",
  route: "getUser/{id}", // Define route with {id} as a parameter
  handler: async (request, context) => {
    try {
      await client.connect();
      const database = client.db("ics-cluster-1");
      const collection = database.collection("users");

      // Extract `id` from route parameters using `context.bindingData`
      const { id } = request.params;
      if (!id) {
        return { status: 400, body: "ID is required" };
      }

      // Convert `id` to ObjectId and fetch the user
      const user = await collection.findOne({ _id: new ObjectId(id) });

      if (!user) {
        return { status: 404, body: "User not found" };
      }

      return { body: JSON.stringify(user) };
    } catch (error) {
      context.error("Error finding user by ID:", error);
      return { status: 500, body: "Error finding user by ID" };
    } finally {
      await client.close();
    }
  },
});
