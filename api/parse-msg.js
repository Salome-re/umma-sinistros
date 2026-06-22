// api/parse-msg.js
// Endpoint serverless para parsear arquivos .msg do Outlook.
// Recebe o arquivo em base64 e retorna body, subject, from e attachments.

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const { fileBase64 } = req.body;
  if (!fileBase64) {
    return res.status(400).json({ error: 'Nenhum arquivo fornecido.' });
  }

  try {
    // Dynamic import para usar no serverless (Node.js)
    const { default: MsgReader } = await import('@kenjiuno/msgreader');
    
    const buffer = Buffer.from(fileBase64, 'base64');
    const reader = new MsgReader(buffer);
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
