exports.handler = async function(event) {
  if (event.httpMethod !== "POST") {
    return {
      statusCode: 405,
      body: JSON.stringify({ error: "Method not allowed" })
    };
  }

  const apiKey = process.env.GEMINI_API_KEY;

  if (!apiKey) {
    return {
      statusCode: 500,
      body: JSON.stringify({ error: "GEMINI_API_KEY is not configured in Netlify." })
    };
  }

  let body;

  try {
    body = JSON.parse(event.body || "{}");
  } catch {
    return {
      statusCode: 400,
      body: JSON.stringify({ error: "Invalid JSON body." })
    };
  }

  const endpoint =
    `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${apiKey}`;

  let prompt = "";

  if (body.mode === "criteria") {
    prompt = `You are a procurement evaluation expert. From the scope below, create JSON only with an array named criteria. Each item must have id, type, parameter, maxMark, guidance. Type must be Technical or Functional.

Scope:
${body.scope}`;
  } else if (body.mode === "score") {
    prompt = `You assist a procurement evaluation committee. Score the vendor response against each criterion.

Criteria:
${JSON.stringify(body.criteria)}

Vendor:
${body.vendorName}

Vendor Response:
${body.vendorText}

Return JSON only with an array named scores.`;
  } else {
    return {
      statusCode: 400,
      body: JSON.stringify({ error: "Unsupported mode." })
    };
  }

  try {
    const response = await fetch(endpoint, {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        contents: [
          {
            parts: [
              {
                text: prompt
              }
            ]
          }
        ]
      })
    });

    const data = await response.json();

    if (!response.ok) {
      return {
        statusCode: response.status,
        body: JSON.stringify(data)
      };
    }

    return {
      statusCode: 200,
      body: JSON.stringify(data)
    };

  } catch (err) {
    return {
      statusCode: 500,
      body: JSON.stringify({ error: err.message })
    };
  }
};
