import express from "express";
import { generateExcel } from "../controllers/excel.controller.js";

const router = express.Router();

router.get("/excel", generateExcel);

export default router;