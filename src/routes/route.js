const pdfRoute = require("./pdfRoute");

const mainRoute = (app) => {
  app.use("/api/v1", pdfRoute);
  app.use("*/api/v1", function (req, res) {
    res.status(404).json({ message: "api not found" });
  });
};

module.exports = { mainRoute };
