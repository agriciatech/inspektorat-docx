const express = require("express");
const { setPDF } = require("../controller/lkeUtama/pdf");
const { setPDFLKE } = require("../controller/lke");
const { setExcelLKE } = require("../controller/lkeExcel");
const router = express.Router();

router.get("/inspkesisgeneratepdf/:id", setPDF);
router.get("/generateinspeksiutama/:id", setPDFLKE);
router.get("/generateinspeksiexcel", setExcelLKE);

module.exports = router;
