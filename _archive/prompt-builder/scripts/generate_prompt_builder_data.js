"use strict";

const fs = require("fs");
const { getRootPath, buildOutputString } = require("./prompt_builder_data_lib");

fs.writeFileSync(getRootPath("prompt_builder_data.js"), buildOutputString(), "utf8");

console.log("Wrote prompt_builder_data.js");
