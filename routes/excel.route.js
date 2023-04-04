// Imported necessary packages
import express from "express";

// Imported the excel controller which includes generateExcel funcion
import { generateExcel } from "../controllers/excel.controller.js";

// Initialized router from excel
const router = express.Router();

// Set the excel route as /excel. So the request url will be /api/excel
router.get("/excel", generateExcel);

export default router;