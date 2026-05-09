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

  const createContractLineItemsV3 = (token, entityName, entityUuid, templateUuid, data) => request(
    `/b2bapi/v3/${encodePath(entityName)}/${encodePath(entityUuid)}/contract-line-item-templates/${encodePath(templateUuid)}/contract-line-items`,
    {
      method: "POST",
      token,
      body: data
    }
  );

  const getContractLineItemTemplateUuid = async (token, entityName, cltId) => {
    const response = await request(`/b2bapi/v3/${encodePath(entityName)}/contract-line-item-templates/search`, {
      method: "POST",
      token,
      body: {
        orderDirection: "desc",
        limit: 200,
        offset: 0,
        fields: [
          "id"
        ]
      }
    });

    const records = Array.isArray(response) ? response : response?.data;
    const template = records?.find((record) => record?.id === cltId);

    if (!template?.s_uuid) {
      throw new Error(`Contract line item template not found for id ${cltId}`);
    }

    return template.s_uuid;
  };

  const getEntityUuid = async (token, entityName, entityId) => {
    const response = await request(`/b2bapi/v3/${encodePath(entityName)}/search`, {
      method: "POST",
      token,
      body: {
        filters: [
          {
            apiName: "id",
            negate: false,
            empty: false,
            value: [
              entityId
            ]
          }
        ],
        orderDirection: "desc",
        offset: 0,
        limit: 1,
        fields: [
          "id"
        ]
      }
    });

    const records = Array.isArray(response) ? response : response?.data;
    const entity = records?.find((record) => record?.id === entityId) ?? records?.[0];

    if (!entity?.s_uuid) {
      throw new Error(`Entity not found for id ${entityId}`);
    }

    return entity.s_uuid;
  };

  return {
    auth,
    createContractLineItems,
    createContractLineItemsV3,
    getContractLineItemTemplateUuid,
    getEntityUuid
  };
}

module.exports = {
  createApi,
  createAPI: createApi
};
