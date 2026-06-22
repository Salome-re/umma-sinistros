// api/parse-msg.js
// Endpoint serverless para parsear arquivos .msg do Outlook.
// Aceita upload via FormData (multipart) para suportar arquivos grandes (até 100MB).
// Também aceita JSON com fileBase64 para retrocompatibilidade.

export const config = {
  api: {
    bodyParser: false, // Desabilitar bodyParser para lidar com multipart manualmente
  },
};

async function parseMultipart(req) {
  // Ler o body inteiro como buffer
  const chunks = [];
  for await (const chunk of req) {
    chunks.push(chunk);
  }
  const body = Buffer.concat(chunks);
  
  const contentType = req.headers['content-type'] || '';
  
  // Se for JSON (retrocompatibilidade)
  if (contentType.includes('application/json')) {
    const json = JSON.parse(body.toString('utf-8'));
    if (json.fileBase64) {
      return Buffer.from(json.fileBase64, 'base64');
    }
    return null;
  }
  
  // Se for multipart/form-data
  if (contentType.includes('multipart/form-data')) {
    const boundary = contentType.split('boundary=')[1];
    if (!boundary) return null;
    
    const boundaryBuffer = Buffer.from(`--${boundary}`);
    const parts = [];
    let start = body.indexOf(boundaryBuffer) + boundaryBuffer.length + 2; // +2 for \r\n
    
    while (true) {
      const nextBoundary = body.indexOf(boundaryBuffer, start);
      if (nextBoundary === -1) break;
      
      const part = body.slice(start, nextBoundary - 2); // -2 for \r\n before boundary
      const headerEnd = part.indexOf('\r\n\r\n');
      if (headerEnd !== -1) {
        const headers = part.slice(0, headerEnd).toString('utf-8');
        const content = part.slice(headerEnd + 4);
        if (headers.includes('filename')) {
          parts.push(content);
        }
      }
      start = nextBoundary + boundaryBuffer.length + 2;
    }
    
    return parts.length > 0 ? parts[0] : null;
  }
  
  // Se for application/octet-stream (raw binary)
  if (contentType.includes('application/octet-stream')) {
    return body;
  }
  
  return null;
}

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const fileBuffer = await parseMultipart(req);
    
    if (!fileBuffer || fileBuffer.length === 0) {
      return res.status(400).json({ error: 'Nenhum arquivo fornecido ou formato inválido.' });
    }

    // Dynamic import para usar no serverless (Node.js)
    const { default: MsgReader } = await import('@kenjiuno/msgreader');
    
    const reader = new MsgReader(fileBuffer);
    const msgData = reader.getFileData();

    const body = msgData.body || "";
    const subject = msgData.subject || "";
    const from = msgData.senderEmail || msgData.senderName || "";
    
    // Extrair anexos com conteúdo em base64
    const attachments = (msgData.attachments || []).map((a, idx) => {
      try {
        const attData = reader.getAttachment(idx);
        return {
          filename: a.fileName || a.name || `attachment_${idx}`,
          contentType: a.mimeType || "",
          contentBase64: attData?.content ? Buffer.from(attData.content).toString('base64') : null
        };
      } catch(e) {
        return {
          filename: a.fileName || a.name || `attachment_${idx}`,
          contentType: "",
          contentBase64: null
        };
      }
    });

    return res.status(200).json({ body, subject, from, attachments });
  } catch (error) {
    return res.status(500).json({ error: 'Erro ao processar .msg: ' + error.message });
  }
}
