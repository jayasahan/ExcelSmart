// Import the Google AI library
const { GoogleGenerativeAI } = require('@google/generative-ai');

// This is the main handler function for the serverless endpoint
module.exports = async (req, res) => {
  // --- NEW: SET CORS HEADERS ---
  // This allows your Vercel domain to be accessed by the add-in.
  res.setHeader('Access-Control-Allow-Origin', 'https://excel-smart.vercel.app');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  // Handle pre-flight requests (sent by browsers to check permissions)
  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }
  // --- END OF NEW SECTION ---

  // 1. Check if the request is a POST request.
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method Not Allowed' });
  }

  try {
    // 2. Initialize the Google AI client using the environment variable
    const genAI = new GoogleGenerativeAI(process.env.GOOGLE_API_KEY);
    const model = genAI.getGenerativeModel({ model: "gemini-pro" });

    // 3. Get the user's query from the request body
    const userQuery = req.body.prompt;
    if (!userQuery) {
      return res.status(400).json({ error: 'Prompt is required' });
    }

    // 4. Create the specific prompt for the AI
    const prompt = `
      You are an expert in Microsoft Excel formulas.
      Based on the following user request, provide only the Excel formula as a raw string.
      Do not explain it, do not wrap it in quotes, and do not add any extra text.
      User request: "${userQuery}"
      Formula:
    `;

    // 5. Call the Gemini API
    const result = await model.generateContent(prompt);
    const response = await result.response;
    const formula = response.text().trim();

    // 6. Send the formula back as a successful JSON response
    res.status(200).json({ formula: formula });

  } catch (error) {
    console.error('Error calling Gemini API:', error);
    res.status(500).json({ error: 'Failed to get formula from AI' });
  }
};