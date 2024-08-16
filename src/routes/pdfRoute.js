const express = require("express");
const { setPDF } = require("../controller/lkeUtama/pdf");
const { setPDFLKE } = require("../controller/lke");
const { setExcelLKE } = require("../controller/lkeExcel/LKEutama");
const { setLKEExcel } = require("../controller/lkeExcel/lke");
const { setExcelLKEUtama } = require("../controller/lkeExcel/utama");
const { setForm9Result } = require("../controller/form9ByTahun/index");
const router = express.Router();

router.get("/inspkesisgeneratepdf/:id", setPDF);
router.get("/generateinspeksiutama/:id", setPDFLKE);
router.get("/generateinspeksiexcel", setExcelLKE);
router.get("/generateexcellke/:id", setLKEExcel);
router.get("/generateexcelutama/:id", setExcelLKEUtama);
// router.get("/form9/:tahun", setForm9);
router.get("/form9/:tahun", setForm9Result);

module.exports = router;
