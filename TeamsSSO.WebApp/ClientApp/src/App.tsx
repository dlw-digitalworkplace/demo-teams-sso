import * as microsoftTeams from "@microsoft/teams-js";
import React, { useEffect, useState } from "react";

const App = () => {
  const [ssoToken, setSsoToken] = useState<string | null>(null);
  const [apiResult, setApiResult] = useState<string | null>(null);

  useEffect(() => {
    (async () => {
      const token = await new Promise<string>((resolve, reject) =>
        microsoftTeams.authentication.getAuthToken({
          successCallback: resolve,
          failureCallback: reject
        })
      );

      setSsoToken(token);
    })();
  }, []);

  const getOBOToken = async (resource: string) => {
    const oboToken = await fetch(
      `/oauth/AcquireOBOToken?resource=${resource}`,
      {
        method: "GET",
        headers: {
          Authorization: `Bearer ${ssoToken}`
        }
      }
    ).then((r) => r.text());

    return oboToken;
  };

  const callWebApi = async () => {
    const result = await fetch("/api/weatherforecast", {
      method: "GET",
      headers: {
        Authorization: `Bearer ${ssoToken}`
      }
    }).then((r) => r.json());

    setApiResult(JSON.stringify(result, null, 2));
  };

  const callGraphApi = async () => {
    const oboToken = await getOBOToken("https://graph.microsoft.com");
    const result = await fetch("https://graph.microsoft.com/v1.0/me", {
      method: "GET",
      headers: {
        Authorization: `Bearer ${oboToken}`
      }
    }).then((r) => r.json());

    setApiResult(JSON.stringify(result, null, 2));
  };

  const callSharePointApi = async () => {
    const sharePointUrl = "https://<SHAREPOINT_URL>";
    const oboToken = await getOBOToken(sharePointUrl);
    const result = await fetch(`${sharePointUrl}/_api/lists`, {
      method: "GET",
      headers: {
        Authorization: `Bearer ${oboToken}`,
        Accept: "application/json"
      }
    }).then((r) => r.json());

    setApiResult(JSON.stringify(result, null, 2));
  };

  return (
    <div>
      {ssoToken && (
        <div>
          <button onClick={callWebApi}>Call Web API</button>
          <button onClick={callGraphApi}>Call Graph API</button>
          <button onClick={callSharePointApi}>Call SharePoint API</button>
        </div>
      )}

      {apiResult && (
        <div>
          <code>{apiResult}</code>
        </div>
      )}
    </div>
  );
};

export default App;
