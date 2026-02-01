/* global DEFAULT_AUTHORITY_HOST */

const KNOWN_AUTHORITY_HOSTS = [
  "login.microsoftonline.com",
  "login.microsoftonline.us",
  "login.chinacloudapi.cn",
  "login.microsoftonline.de",
  "login.partner.microsoftonline.cn"
];

function normalizeAuthorityHost(value) {
  const raw = String(value || "").trim();
  if (!raw) {
    return { ok: false, error: "Authority host is required." };
  }
  let url = null;
  try {
    url = new URL(raw);
  } catch (err) {
    return { ok: false, error: "Authority host must be a valid URL (https://...)." };
  }
  if (url.protocol !== "https:") {
    return { ok: false, error: "Authority host must use https://." };
  }
  if (url.username || url.password) {
    return { ok: false, error: "Authority host must not include credentials." };
  }
  if (url.search || url.hash) {
    return { ok: false, error: "Authority host must not include query or hash." };
  }
  if (url.pathname && url.pathname !== "/") {
    return { ok: false, error: "Authority host must not include a path." };
  }
  const normalized = `https://${url.host}`;
  const warning = KNOWN_AUTHORITY_HOSTS.includes(url.hostname)
    ? ""
    : "Authority host is not a known Microsoft login endpoint.";
  return { ok: true, normalized, warning };
}

function normalizeTenant(value) {
  let tenant = String(value || "").trim();
  if (!tenant) {
    return { ok: false, error: "Tenant is required." };
  }
  const aliases = {
    organisation: "organizations",
    organisations: "organizations",
    consumer: "consumers"
  };
  if (Object.prototype.hasOwnProperty.call(aliases, tenant)) {
    tenant = aliases[tenant];
  }
  const normalized = tenant;
  const special = ["organizations", "common", "consumers"];
  if (special.includes(tenant)) {
    return { ok: true, normalized };
  }
  const guid = /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/;
  if (guid.test(tenant)) {
    return { ok: true, normalized };
  }
  const domain = /^[a-zA-Z0-9-]+(\.[a-zA-Z0-9-]+)+$/;
  if (domain.test(tenant)) {
    return { ok: true, normalized };
  }
  return { ok: false, error: "Tenant must be a GUID, domain, or one of: organizations, common, consumers." };
}

function validateSettings(settings) {
  const errors = [];
  const warnings = [];

  const authority = normalizeAuthorityHost(settings.authorityHost || DEFAULT_AUTHORITY_HOST);
  if (!authority.ok) {
    errors.push(authority.error);
  } else if (authority.warning && !settings.allowCustomAuthorityHost) {
    warnings.push(authority.warning);
  }

  const tenant = normalizeTenant(settings.tenant);
  if (!tenant.ok) {
    errors.push(tenant.error);
  }

  return {
    ok: errors.length === 0,
    errors,
    warnings,
    normalized: {
      authorityHost: authority.normalized || "",
      tenant: tenant.normalized || ""
    }
  };
}
