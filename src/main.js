const express = require("express");
const app = express();
var cors = require("cors");
const swaggerUi = require("swagger-ui-express");
const swaggerDocument = require("../swagger-output.json");
const { mainRoute } = require("./routes/route");

const run = async () => {
  try {
    app.use(express.json());
    app.use(
      cors({
        origin: "*",
        allowedHeaders: [
          "Accept-Version",
          "Authorization",
          "Credentials",
          "Content-Type",
        ],
      })
    );
    app.use("/api-docs", swaggerUi.serve, swaggerUi.setup(swaggerDocument));

    mainRoute(app);

    app.use("/uploads", express.static("uploads"));

    const PORT = process.env.PORT || 3000;

    app.listen(PORT, () => {
      console.log(`Server berjalan di port ${PORT}`);
    });
  } catch (err) {
    console.log(err);
  }
};

run();
