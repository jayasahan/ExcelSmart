/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// --- NEW: Define UI elements at the top for easy access ---
let generateButton, applyButton, loader, formulaOutput, statusText;

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.addEventListener("DOMContentLoaded", setupUI);
  }
});

function setupUI() {
  // Get references to all our interactive elements
  generateButton = document.getElementById("generate-button");
  applyButton = document.getElementById("apply-button");
  loader = document.getElementById("loader");
  formulaOutput = document.getElementById("formula-output");
  statusText = document.getElementById("status-text");

  // Add event listeners
  generateButton.addEventListener("click", handleGenerateFormula);
  applyButton.addEventListener("click", applyFormula);
}

// --- NEW: Helper function to manage the UI state ---
function setLoadingState(isLoading) {
  if (isLoading) {
    generateButton.disabled = true;
    applyButton.disabled = true;
    loader.style.display = "inline-block"; // Show spinner
    formulaOutput.textContent = "Generating...";
    statusText.textContent = "";
  } else {
    generateButton.disabled = false;
    applyButton.disabled = false;
    loader.style.display = "none"; // Hide spinner
  }
}

async function handleGenerateFormula() {
  const promptText = document.getElementById("prompt-input").value;

  if (!promptText) {
    formulaOutput.textContent = "Please enter a description first.";
    return;
  }
  
  setLoadingState(true); // --- UX Change: Enter loading state ---

  try {
    const serverUrl = "https://excel-smart.vercel.app/api/get-formula"; // Your live Vercel URL
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
    formulaOutput.textContent = "Error: Could not retrieve formula.";
  } finally {
    setLoadingState(false); // --- UX Change: Exit loading state, no matter what happens ---
  }
}

async function applyFormula() {
  const formula = formulaOutput.textContent;

  if (!formula || formula === "..." || formula.startsWith("Error") || formula.startsWith("Generating")) {
    statusText.textContent = "Generate a valid formula before applying.";
    return;
  }

  try {
    await Excel.run(async (context) => {
      const cell = context.workbook.getActiveCell();
      cell.formulas = [[formula]];
      cell.load("address");
      await context.sync();
      statusText.textContent = `Formula applied to cell ${cell.address}!`;
    });
  } catch (error) {
    console.error(error);
    statusText.textContent = "Error: Could not apply formula to the selected cell.";
  }
}