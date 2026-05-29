// api/analyze.js
// Proxy seguro: a chave da API fica no servidor (Vercel), nunca exposta ao cliente.
// Configure a variável de ambiente OPENAI_API_KEY no painel do Vercel.

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    return res.status(500).json({
      error: 'OPENAI_API_KEY não configurada. Configure a variável de ambiente no painel do Vercel.'
    });
  }

  const { messages, system } = req.body;

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
            content: system || 'Você é um especialista em regulação de sinistros de seguros no Brasil, com profundo conhecimento da Lei 15.040/2024, Circular SUSEP 621/2021, Circular SUSEP 667/2022, Resolução CNSP 407/2021 e Resolução CNSP 445/2022.'
          },
          ...(messages || []),
        ],
      }),
    });

    if (!response.ok) {
      const err = await response.text();
      return res.status(response.status).json({ error: err });
    }

    const data = await response.json();
    const text = data.choices?.[0]?.message?.content || '';
    // Retorna no mesmo formato esperado pelo frontend (compatível com Anthropic)
    return res.status(200).json({ content: [{ text }] });
  } catch (error) {
    return res.status(500).json({ error: 'Erro ao conectar com a API: ' + error.message });
  }
}
