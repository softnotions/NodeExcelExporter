import express from "express";
import dotenv from "dotenv";
import cors from "cors";
import bodyParser from "body-parser";
import process from "process";

import ExcelRoute from "./routes/excel.route.js";

const app = express();

app.use(express.json());
app.use(bodyParser.json({ limit: "30mb", extended: true }));
app.use(bodyParser.urlencoded({ limit: "30mb", extended: true }));
app.use(cors());

dotenv.config();

const PORT = process.env.PORT;

app.use("/api", ExcelRoute);

app.listen(PORT, () => {
	console.log(`Listening at ${PORT}`);
});