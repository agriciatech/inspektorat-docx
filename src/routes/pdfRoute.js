const express = require("express");
const { setPDF } = require("../controller/lkeUtama/pdf");
const { setPDFLKE } = require("../controller/lke");
const { setExcelLKE } = require("../controller/lkeExcel/LKEutama");
const { setLKEExcel } = require("../controller/lkeExcel/lke");
const { setExcelLKEUtama } = require("../controller/lkeExcel/utama");
const router = express.Router();

router.get("/inspkesisgeneratepdf/:id", setPDF);
router.get("/generateinspeksiutama/:id", setPDFLKE);
router.get("/generateinspeksiexcel", setExcelLKE);
router.get("/generateexcellke/:id", setExcelLKEUtama);
router.get("/generateexcelutama/:id", setLKEExcel);

module.exports = router;
