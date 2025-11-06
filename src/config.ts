export const msalConfig = {
  auth: {
    clientId: "0be831ea-013d-488c-b174-b3c85bf4bc61",
    authority: "https://login.microsoftonline.com/e77be010-7134-4688-b3a1-63ebbdaafe4c",
    redirectUri: "https://localhost:3000/taskpane.html"
  },
  cache: { cacheLocation: "localStorage" }
};

export const loginRequest = {
  scopes: ["User.Read", "Calendars.ReadWrite", "offline_access"]
};
