const fs = require("fs");
const path = require("path");
const vm = require("vm");
const assert = require("assert");

function loadGlobals(...files) {
  const context = vm.createContext({ console, URL });
  for (const file of files) {
    const code = fs.readFileSync(file, "utf8");
    vm.runInContext(code, context, { filename: file });
  }
  return context;
}

const root = path.resolve(__dirname, "..");
const constantsPath = path.join(root, "shared", "constants.js");
const validationPath = path.join(root, "shared", "validation.js");
const optionsPath = path.join(root, "options", "options.js");

const ctx = loadGlobals(constantsPath, validationPath);
const get = expression => vm.runInContext(expression, ctx);

assert.strictEqual(get("DEFAULT_APPLICATION_ID"), "REPLACE_WITH_APPLICATION_ID");
assert.strictEqual(get("isPlaceholder")("REPLACE_WITH_APPLICATION_ID"), true);
assert.strictEqual(get("isPlaceholder")("not-a-placeholder"), false);

const okAuthority = get("normalizeAuthorityHost")("https://login.microsoftonline.com");
assert.strictEqual(okAuthority.ok, true);
assert.strictEqual(okAuthority.normalized, "https://login.microsoftonline.com");

const badAuthority = get("normalizeAuthorityHost")("http://login.microsoftonline.com");
assert.strictEqual(badAuthority.ok, false);

const pathAuthority = get("normalizeAuthorityHost")("https://login.microsoftonline.com/tenant");
assert.strictEqual(pathAuthority.ok, false);

const okTenantGuid = get("normalizeTenant")("9188040d-6c67-4c5b-b112-36a304b66dad");
assert.strictEqual(okTenantGuid.ok, true);

const okTenantDomain = get("normalizeTenant")("example.onmicrosoft.com");
assert.strictEqual(okTenantDomain.ok, true);

const okTenantSpecial = get("normalizeTenant")("organizations");
assert.strictEqual(okTenantSpecial.ok, true);

const badTenant = get("normalizeTenant")("not a tenant");
assert.strictEqual(badTenant.ok, false);

const validation = get("validateSettings")({
  authorityHost: "https://login.microsoftonline.com",
  tenant: "organizations"
});
assert.strictEqual(validation.ok, true);

const warningSuppressed = get("validateSettings")({
  authorityHost: "https://login.contoso.example",
  tenant: "organizations",
  allowCustomAuthorityHost: true
});
assert.strictEqual(warningSuppressed.ok, true);
assert.strictEqual(warningSuppressed.warnings.length, 0);

console.log("All tests passed.");

const optionsSource = fs.readFileSync(optionsPath, "utf8");
assert.ok(
  /browser\.storage\.local\.get\(\{[^}]*tenant:\s*DEFAULT_TENANT[^}]*\}\)/s.test(optionsSource),
  "options.js should load tenant default before validating test connection"
);

console.log("UI defaults test passed.");
