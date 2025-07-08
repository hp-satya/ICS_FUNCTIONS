const { app } = require("@azure/functions");
const { MongoClient, ObjectId } = require("mongodb");

const connectionString = process.env["MongoDBConnectionString"];
const client = new MongoClient(connectionString);

app.http("updateUser", {
  methods: ["PUT"],
  authLevel: "anonymous",
  route: "updateUser/{id}", // Define route with {id} as a parameter
  handler: async (request, context) => {
    try {
      await client.connect();
      const database = client.db("ics-cluster-1");
      const collection = database.collection("users");

      // Extract `id` from route parameters using `context.bindingData`
      const { id } = request.params;

      if (!id) {
        return { status: 400, body: "User ID is required in the URL." };
      }

      const update = await request.json(); // Assuming the update fields are sent in the request body

      // Validation logic for update fields
      const errors = [];
      if (
        update.name &&
        (typeof update.name !== "string" || update.name.trim().length < 3)
      ) {
        errors.push("Name must be a string with at least 3 characters.");
      }
      if (
        update.email &&
        (typeof update.email !== "string" ||
          !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(update.email))
      ) {
        errors.push("Email must be a valid email address.");
      }
      if (
        update.mobile &&
        (typeof update.mobile !== "string" || !/^\d{10}$/.test(update.mobile))
      ) {
        errors.push("Mobile must be a valid 10-digit number.");
      }
      if (
        update.age &&
        (typeof update.age !== "number" || update.age < 18 || update.age > 100)
      ) {
        errors.push("Age must be a number between 18 and 100.");
      }

      // If there are validation errors, return them
      if (errors.length > 0) {
        return { status: 400, body: { errors } };
      }

      // Check for duplicate name or email (excluding the current user being updated)
      const duplicateUser = await collection.findOne({
        $or: [{ name: update.name }, { email: update.email }],
        _id: { $ne: new ObjectId(id) }, // Exclude the current user
      });

      if (duplicateUser) {
        return {
          status: 400,
          body: "Another user with the same name or email already exists.",
        };
      }

      // Perform the update
      const result = await collection.updateOne(
        { _id: new ObjectId(id) },
        { $set: update }
      );

      if (result.matchedCount === 0) {
        return { status: 404, body: "User not found." };
      }

      return {
        body: `Matched ${result.matchedCount} document(s) and modified ${result.modifiedCount} document(s).`,
      };
    } catch (error) {
      context.error("Error updating user:", error);
      return { status: 500, body: "Error updating user." };
    } finally {
      await client.close();
    }
  },
});
