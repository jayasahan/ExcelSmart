/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.addEventListener("DOMContentLoaded", setupUI);
  }
});

function setupUI() {
  // Add event listeners to both buttons now
  document.getElementById("generate-button").addEventListener("click", handleGenerateFormula);
  document.getElementById("apply-button").addEventListener("click", applyFormula); // New listener
}

async function handleGenerateFormula() {
  const promptText = document.getElementById("prompt-input").value;
  const formulaOutput = document.getElementById("formula-output");
  const statusText = document.getElementById("status-text");

  statusText.textContent = ""; // Clear previous status
  if (!promptText) {
    formulaOutput.textContent = "Please enter a description first.";
    return;
  }
  formulaOutput.textContent = "Generating...";

  try {
    //const serverUrl = "http://localhost:3001/api/get-formula";   //For local run
    const serverUrl = "https://excel-smart.vercel.app/api/get-formula";
    const response = await fetch(serverUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ prompt: promptText }),
    });
    if (!response.ok) {
      throw new Error(`Server error: ${response.statusText}`);
    }
    const data = await response.json();
    formulaOutput.textContent = data.formula;
  } catch (error) {
    console.error("Failed to fetch formula:", error);
    formulaOutput.textContent = "Error: Could not retrieve formula. Is the backend server running?";
  }
}

// --- NEW FUNCTION TO WRITE TO EXCEL ---
async function applyFormula() {
  const formula = document.getElementById("formula-output").textContent;
  const statusText = document.getElementById("status-text");

  // 1. Check if there's a valid formula to apply
  if (!formula || formula === "..." || formula.startsWith("Error") || formula.startsWith("Generating")) {
    statusText.textContent = "Please generate a valid formula first.";
    return;
  }

  try {
    // 2. Use Excel.run to execute a batch of commands against the workbook
    await Excel.run(async (context) => {
      // 3. Get the currently selected cell (the "active cell")
      const cell = context.workbook.getActiveCell();

      // 4. Set the formula of that cell
      cell.formulas = [[formula]]; // Formulas are set as a 2D array

      // 5. Load the cell's address to provide feedback
      cell.load("address");

      // 6. context.sync() executes all the commands we've queued up (like a "commit")
      await context.sync();

      // 7. Update the status text with the cell address
      statusText.textContent = `Formula applied to cell ${cell.address}!`;
    });
  } catch (error) {
    // Handle any errors from the Excel.run operation
    console.error(error);
    statusText.textContent = "Error: Could not apply formula to the cell.";
  }
}