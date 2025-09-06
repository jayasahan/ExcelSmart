// 1. Import necessary libraries
const express = require('express');
const { GoogleGenerativeAI } = require('@google/generative-ai');
const dotenv = require('dotenv');
const cors = require('cors');

// 2. Load environment variables from the .env file
dotenv.config();

// 3. Initialize the Express app and enable CORS
const app = express();
app.use(cors()); // Allows our Excel add-in to talk to this server
app.use(express.json()); // Allows the server to understand JSON requests

// 4. Initialize the Google Generative AI client
const genAI = new GoogleGenerativeAI(process.env.GOOGLE_API_KEY);
const model = genAI.getGenerativeModel({ model: "gemini-2.5-flash" });

// 5. Define the API endpoint that the Excel Add-in will call
app.post('/api/get-formula', async (req, res) => {
  try {
    // Get the user's query from the request body
    const userQuery = req.body.prompt;
    if (!userQuery) {
      return res.status(400).json({ error: 'Prompt is required' });
    }

    // Create a specific, detailed prompt for the AI
    const prompt = `
      You are an expert in Microsoft Excel formulas.
      Based on the following user request, provide only the Excel formula as a raw string.
      Do not explain it, do not wrap it in quotes, and do not add any extra text.
      User request: "${userQuery}"
      Formula:
    `;

    // 6. Call the Gemini API
    const result = await model.generateContent(prompt);
    const response = await result.response;
    const formula = response.text().trim();

    // 7. Send the formula back to the Excel Add-in
    res.json({ formula: formula });

  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Failed to get formula from AI' });
  }
});

// 8. Start the server
const PORT = 3001; // We'll use a different port from the main add-in server
app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});