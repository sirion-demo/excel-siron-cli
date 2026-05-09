// api.js

const Bottleneck = require("bottleneck");

const limiter = new Bottleneck({
  reservoir: 49,
  reservoirRefreshAmount: 49,
  reservoirRefreshInterval: 60 * 1000,
});

function createApi(baseUrl) {
  const normalizedBaseUrl = String(baseUrl ?? "").trim().replace(/\/$/, "");
  const apiBaseUrl = /^https?:\/\//i.test(normalizedBaseUrl)
    ? normalizedBaseUrl
    : `https://${normalizedBaseUrl}`;

  async function request(path, options = {}) {
    return limiter.schedule(async () => {
      const { headers, body, auth = true, token, ...fetchOptions } = options;
      if (auth && !token) {
        throw new Error("API token is required for this request");
      }

      const authHeader = auth ? { Authorization: `Bearer ${token}` } : {};

      const res = await fetch(`${apiBaseUrl}${path}`, {
        ...fetchOptions,
        headers: {
          ...authHeader,
          "Content-Type": "application/json",
          ...headers
        },
        body: body === undefined ? undefined : JSON.stringify(body)
      });

      if (!res.ok) {
        throw new Error(`${res.status} ${await res.text()}`);
      }

      if (res.status === 204) {
        return null;
      }

      const text = await res.text();
      if (!text) {
        return null;
      }

      const contentType = res.headers.get("content-type") || "";
      return contentType.includes("application/json") ? JSON.parse(text) : text;
    });
  }

  const encodePath = (value) => encodeURIComponent(String(value));

  const clientTokenResponse = (oauthClientId, oauthClientSecret, loginId) => request("/oauth/client/token", {
    method: "POST",
    auth: false,
    body: {
      oauthClientId,
      oauthClientSecret,
      loginId
    }
  });

  const clientToken = async (oauthClientId, oauthClientSecret, loginId) => {
    const tokenResponse = await clientTokenResponse(oauthClientId, oauthClientSecret, loginId);
    const token = tokenResponse?.data?.access_token;

    if (!token) {
      throw new Error("Token response did not include an access token");
    }

    return token;
  };

  const auth = {
    clientToken,
    clientTokenResponse
  };

  const createContractLineItems = (token, entityName, entityId, templateId, data) => request(
    `/b2bapi/v2/${encodePath(entityName)}/${encodePath(entityId)}/contract-line-item-templates/${encodePath(templateId)}/contract-line-items?async=false`,
    {
      method: "POST",
      token,
      body: data
    }
  );

  return {
    auth,
    createContractLineItems
  };
}

module.exports = {
  createApi,
  createAPI: createApi
};
