const { app } = require("@azure/functions");
const { MongoClient, ObjectId } = require("mongodb");

const connectionString = process.env["MongoDBConnectionString"];
const client = new MongoClient(connectionString);

app.http("deleteUser", {
  methods: ["DELETE"],
  authLevel: "anonymous",
  route: "deleteUser/{id}", // Define route with {id} as a parameter
  handler: async (request, context) => {
    try {
      await client.connect();
      const database = client.db("ics-cluster-1");
      const collection = database.collection("users");

      const { id } = request.params;

      if (!id) {
        return { status: 400, body: "User ID is required in the URL." };
      }

      // Perform the delete operation
      const result = await collection.deleteOne({
        _id: new ObjectId(id),
      });

      if (result.deletedCount === 0) {
        return { status: 404, body: "User not found." };
      }

      return { body: `Deleted ${result.deletedCount} document(s)` };
    } catch (error) {
      context.error("Error deleting user:", error);
      return { status: 500, body: "Error deleting user." };
    } finally {
      await client.close();
    }
  },
});
