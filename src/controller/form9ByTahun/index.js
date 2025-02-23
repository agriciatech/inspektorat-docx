const { enpoint } = require("../../config/url");
const { setForm9 } = require("./doc");
const { setForm9page2 } = require("./doc2");
const { default: axios } = require("axios");

exports.setForm9Result = async (req, res) => {
  const { tahun } = req.params;
  const { idOPD } = req.query;

  await axios
    .get(`${enpoint}/api/v1/formulir9/view?tahun=${tahun}&opd=${idOPD}`)
    .then(async function (response) {
      console.log(response.data);

      let page1 = await setForm9(tahun, response.data);
      let page2 = await setForm9page2(tahun, response.data);

      res.status(200).json({
        status: 200,
        message: "Success",
        data: [page1, page2],
      });
    })
    .catch(function (error) {
      res.status(500).json({
        status: 500,
        message: error,
        data: null,
      });
    });
};
