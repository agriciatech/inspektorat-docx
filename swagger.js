const swaggerAutogen = require("swagger-autogen")();

const doc = {
  info: {
    title: "Pest Control APi",
    description: "Description",
  },
  servers: [
    {
      url: "http://localhost:3000/",
      description: "main server",
    },
    {
      url: "https://apps.sucofindo.co.id/pest-control-dev/be",
      description: "the other server",
    },
  ],
  schemes: ["http"],
};

const outputFile = "./swagger-output.json";
const routes = ["./src/main.js", "./src/routes/route.js"];

swaggerAutogen(outputFile, routes, doc).then(async () => {
  await require("./src/main.js");
});
