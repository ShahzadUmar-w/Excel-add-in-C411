import { createNestablePublicClientApplication } from "@azure/msal-browser";
import { updateOneDriveExcelCell } from "./UpdateCell"; // Assuming the function below is in this file

let pca: any = undefined; // Using 'any' to simplify typing for this example

Office.onReady(async (info) => {
  if (info.host) {
    // Initialize the public client application
    pca = await createNestablePublicClientApplication({
      auth: {
        clientId: "c265bfe4-51ce-4b0f-b74c-d7cb02ac4936",
        authority: "https://login.microsoftonline.com/common",
      },
    });
  }
});

export async function Login_Function() {
  // Specify scopes needed for the access token to read AND write.
  const tokenRequest = {
    scopes: ["Files.ReadWrite", "User.Read", "openid", "profile"],
  };
  let accessToken = null;

  try {
    console.log("Trying to acquire token silently...");
    const userAccount = await pca.acquireTokenSilent(tokenRequest);
    console.log("Acquired token silently.");
    accessToken = userAccount.accessToken;
  } catch (error) {
    console.log(`Silent token acquisition failed (this is expected when new scopes are added): ${error}`);
  }

  if (accessToken === null) {
    // Acquire token silent failure. Send an interactive request via popup.
    try {
      console.log("Trying to acquire token interactively...");
      const userAccount = await pca.acquireTokenPopup(tokenRequest);
      console.log("Acquired token interactively.");
      accessToken = userAccount.accessToken;
    } catch (popupError) {
      // Acquire token interactive failure.
      console.log(`Interactive token acquisition failed: ${popupError}`);
    }
  }

  // Log error if both silent and popup requests failed.
  if (accessToken === null) {
    console.error(`Unable to acquire access token.`);
    return;
  }

  // Call the Microsoft Graph API with the access token.
  runUpdate(accessToken).catch(console.error);
}

async function runUpdate(accessTokenA: any) {
  console.log("Running update function...");

  const accessToken = accessTokenA;
  const filePath = "/Book.xlsx"; // Path in your OneDrive root
  const sheetName = "Sheet1";

  // Using a specific, finite range
  const rangeAddress = "A1";

  // Formatting the value as a 2D array (array of rows)
  const newValue:any = [["test again"]]; // 2D array for the cell value

  await updateOneDriveExcelCell(
    accessToken,
    filePath,
    sheetName,
    rangeAddress,
    newValue
  );
}