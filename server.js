// Imported necessary packages
import express from "express";
import dotenv from "dotenv";
import cors from "cors";
import bodyParser from "body-parser";
import process from "process";

// Import custom route we have added to generate the excel file
import ExcelRoute from "./routes/excel.route.js";

// Create an instance of the express server
const app = express();

// Use middleware to parse incoming requests in JSON format
app.use(express.json());

// Use body-parser middleware to handle request data in specific formats
// Set a limit of 30 megabytes for incoming request bodies to avoid crashes or attacks
app.use(bodyParser.json({ limit: "30mb", extended: true }));
app.use(bodyParser.urlencoded({ limit: "30mb", extended: true }));

// Enable CORS to allow cross-origin requests
app.use(cors());

// Load environment variables from a .env file
dotenv.config();

// Configured the port to listen to incoming requests using env file
const PORT = process.env.PORT;

// Configured the base url as /api
app.use("/api", ExcelRoute);

// Start the server and listen for incoming requests
app.listen(PORT, () => {
	console.log(`Listening at ${PORT}`);
});