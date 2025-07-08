const { app } = require("@azure/functions");
const { MongoClient } = require("mongodb");

const connectionString = process.env["MongoDBConnectionString"];
const client = new MongoClient(connectionString);

app.http("createUser", {
  methods: ["POST"],
  authLevel: "anonymous",
  handler: async (request, context) => {
    try {
      await client.connect();
      const database = client.db("ics-cluster-1");
      const collection = database.collection("users");

      const user = await request.json(); // Assuming the user is sent in the request body

      // Validation logic
      const errors = [];
      if (
        !user.name ||
        typeof user.name !== "string" ||
        user.name.trim().length < 3
      ) {
        errors.push("Name must be a string with at least 3 characters.");
      }
      if (
        !user.email ||
        typeof user.email !== "string" ||
        !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(user.email)
      ) {
        errors.push("Email must be a valid email address.");
      }
      if (
        !user.mobile ||
        typeof user.mobile !== "string" ||
        !/^\d{10}$/.test(user.mobile)
      ) {
        errors.push("Mobile must be a valid 10-digit number.");
      }
      if (
        !user.age ||
        typeof user.age !== "number" ||
        user.age < 18 ||
        user.age > 100
      ) {
        errors.push("Age must be a number between 18 and 100.");
      }

      // If there are validation errors, return them
      if (errors.length > 0) {
        return { status: 400, body: { errors } };
      }

      // Check for duplicate name or email
      const duplicateUser = await collection.findOne({
        $or: [{ name: user.name }, { email: user.email }],
      });

      if (duplicateUser) {
        return {
          status: 400,
          body: "A user with the same name or email already exists.",
        };
      }

      // Insert the validated user into the database
      const result = await collection.insertOne(user);

      return { body: `User created with ID: ${result.insertedId}` };
    } catch (error) {
      context.error("Error creating user:", error);
      return { status: 500, body: "Error creating user" };
    } finally {
      await client.close();
    }
  },
});
