// api/ocr.js
// Endpoint para OCR de imagens usando GPT-4 Vision.
// Recebe uma imagem em base64 e retorna o texto extraído.

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    return res.status(500).json({
      error: 'OPENAI_API_KEY não configurada.'
    });
  }

  const { image, mimeType } = req.body;

  if (!image) {
    return res.status(400).json({ error: 'Nenhuma imagem fornecida.' });
  }

  try {
    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`,
      },
      body: JSON.stringify({
        model: 'gpt-4.1-mini',
        max_tokens: 4096,
        messages: [
          {
            role: 'system',
            content: 'Você é um especialista em OCR. Extraia TODO o texto visível na imagem, preservando a estrutura (parágrafos, tabelas, listas). Se for um documento (laudo, BO, nota fiscal, carta), transcreva fielmente. Retorne apenas o texto extraído, sem comentários adicionais.'
          },
          {
            role: 'user',
            content: [
              {
                type: 'text',
                text: 'Extraia todo o texto desta imagem. Se for um documento, transcreva fielmente preservando a formatação.'
              },
              {
                type: 'image_url',
                image_url: {
                  url: `data:${mimeType || 'image/jpeg'};base64,${image}`,
                  detail: 'high'
                }
              }
            ]
          }
        ],
      }),
    });

    if (!response.ok) {
      const err = await response.text();
      return res.status(response.status).json({ error: err });
    }

    const data = await response.json();
    const text = data.choices?.[0]?.message?.content || '';
    return res.status(200).json({ text });
  } catch (error) {
    return res.status(500).json({ error: 'Erro OCR: ' + error.message });
  }
}
