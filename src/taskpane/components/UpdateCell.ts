// // updateOneDriveExcelCell.ts

// export async function updateOneDriveExcelCell(
//   accessToken: string,
//   filePath: string, // Example: "/Documents/MyExcelFile.xlsx"
//   sheetName: string,
//   rangeAddress: string,
//   newValue: string
// ): Promise<void> {
//   try {
//     const endpoint = `https://graph.microsoft.com/v1.0/me/drive/root:${encodeURIComponent(filePath)}:/workbook/worksheets('${encodeURIComponent(sheetName)}')/range(address='${encodeURIComponent(rangeAddress)}')`;

//     const response = await fetch(endpoint, {
//       method: "PATCH",
//       headers: {
//         "Authorization": `Bearer ${accessToken}`,
//         "Content-Type": "application/json"
//       },
//       body: JSON.stringify({
//         values: [[newValue]]
//       })
//     });

//     if (!response.ok) {
//       const errorText = await response.text();
//       throw new Error(`Graph API Error: ${response.status} - ${errorText}`);
//     }

//     console.log(`Cell ${rangeAddress} updated successfully to "${newValue}" in ${filePath}`);
//   } catch (error) {
//     console.error("Error updating Excel cell:", error);
//     throw error;
//   }
// }






// UpdateCell.ts

/**
 * Updates a specified range in an Excel workbook in OneDrive using the Microsoft Graph API.
 * @param accessToken The Microsoft Graph API access token.
 * @param filePath The path to the workbook in OneDrive (e.g., "/Documents/report.xlsx").
 * @param sheetName The name of the worksheet to update.
 *param rangeAddress The range to update (e.g., "A1", "C2:D4").
 * @param values A 2D array of values to write to the range.
 */
export async function updateOneDriveExcelCell(
  accessToken: string,
  filePath: string,
  sheetName: string,
  rangeAddress: string,
  values: any[][] // Accepts a 2D array
) {
  // Construct the Graph API endpoint URL. Note the encoding for file path and sheet name.
  const encodedFilePath = encodeURIComponent(filePath);
  const encodedSheetName = encodeURIComponent(sheetName);
  const encodedRangeAddress = encodeURIComponent(rangeAddress);

  const url = `https://graph.microsoft.com/v1.0/me/drive/root:${encodedFilePath}:/workbook/worksheets/${encodedSheetName}/range(address='${encodedRangeAddress}')`;

  // The request body must be a JSON object with a "values" key.
  const body = {
    values: values,
  };

  console.log(`Sending PATCH request to: ${url}`);
  console.log(`With body: ${JSON.stringify(body)}`);

  try {
    const response = await fetch(url, {
      method: "PATCH",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(body),
    });

    // Check if the request was successful
    if (response.ok) {
      const data = await response.json();
      console.log("Successfully updated Excel cell. Response:", data);
      // You can add a success notification for the user here
    } else {
      // If the response is not OK, it's an error from the Graph API
      const errorData = await response.json();
      console.error(
        "Failed to update Excel cell. Status:",
        response.status,
        "Error:",
        JSON.stringify(errorData, null, 2) // Pretty-print the JSON error
      );
    }
  } catch (error) {
    // This catches network errors or other issues with the fetch call itself
    console.error("An unexpected error occurred during the fetch operation:", error);
  }
}