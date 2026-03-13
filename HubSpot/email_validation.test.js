const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

function loadHubSpotCode_() {
  const source = fs.readFileSync(path.join(__dirname, "Code.gs"), "utf8");
  const context = {
    console,
    URLSearchParams
  };

  vm.createContext(context);
  vm.runInContext(source, context, { filename: "Code.gs" });

  return context;
}

const { getHubSpotPrimaryEmailValue_ } = loadHubSpotCode_();

test("keeps a single email unchanged", () => {
  assert.equal(
    getHubSpotPrimaryEmailValue_("creator@example.com"),
    "creator@example.com"
  );
});

test("returns the first email when multiple are comma separated", () => {
  assert.equal(
    getHubSpotPrimaryEmailValue_("creator@example.com, manager@example.com"),
    "creator@example.com"
  );
});

test("trims whitespace and skips blank comma-separated segments", () => {
  assert.equal(
    getHubSpotPrimaryEmailValue_("  , creator@example.com , manager@example.com "),
    "creator@example.com"
  );
});
