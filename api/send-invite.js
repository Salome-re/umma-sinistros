// api/send-invite.js
// Envia convite por e-mail via Resend API
// Configure RESEND_API_KEY no painel do Vercel
export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const apiKey = process.env.RESEND_API_KEY;
  if (!apiKey) {
    return res.status(500).json({
      error: 'RESEND_API_KEY não configurada. Configure a variável de ambiente no painel do Vercel.'
    });
  }

  const { email, nome, role, convidadoPor } = req.body;

  if (!email || !nome) {
    return res.status(400).json({ error: 'Email e nome são obrigatórios.' });
  }

  const appUrl = process.env.VERCEL_URL 
    ? `https://${process.env.VERCEL_URL}` 
    : 'https://umma-sinistros.vercel.app';

  const linkAcesso = `${appUrl}?invite=true`;

  const htmlEmail = `
    <div style="font-family: 'IBM Plex Sans', Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 40px 20px;">
      <div style="text-align: center; margin-bottom: 30px;">
        <img src="${appUrl}/assets/umma_logo_dark.png" alt="UMMA" style="height: 60px;" />
      </div>
      <div style="background: #F8FAFC; border-radius: 12px; padding: 32px; border: 1px solid #E2E8F0;">
        <h2 style="color: #02124C; margin: 0 0 16px 0; font-size: 22px;">
          Olá, ${nome}! 👋
        </h2>
        <p style="color: #334155; font-size: 15px; line-height: 1.6; margin: 0 0 20px 0;">
          Você foi convidado(a) por <strong>${convidadoPor || 'um administrador'}</strong> para acessar a plataforma 
          <strong>UMMA Sinistros</strong> como <strong>${role === 'admin' ? 'Administrador(a)' : 'Visualizador(a)'}</strong>.
        </p>
        <p style="color: #334155; font-size: 15px; line-height: 1.6; margin: 0 0 24px 0;">
          Para acessar, basta clicar no botão abaixo e inserir seu e-mail (<strong>${email}</strong>):
        </p>
        <div style="text-align: center; margin: 28px 0;">
          <a href="${linkAcesso}" 
             style="background: #3E75FF; color: white; padding: 14px 32px; border-radius: 8px; 
                    text-decoration: none; font-weight: 700; font-size: 15px; display: inline-block;">
            Acessar UMMA Sinistros →
          </a>
        </div>
        <p style="color: #64748B; font-size: 13px; line-height: 1.5; margin: 20px 0 0 0; border-top: 1px solid #E2E8F0; padding-top: 16px;">
          Se você não esperava este convite, pode ignorar este e-mail com segurança.
        </p>
      </div>
      <p style="color: #94A3B8; font-size: 11px; text-align: center; margin-top: 24px;">
        UMMA Corretora de Seguros · Plataforma de Gestão de Sinistros
      </p>
    </div>
  `;

  try {
    const response = await fetch('https://api.resend.com/emails', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${apiKey}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        from: 'UMMA Sinistros <noreply@ummaseguros.com.br>',
        to: [email],
        subject: `${convidadoPor || 'UMMA'} convidou você para a plataforma UMMA Sinistros`,
        html: htmlEmail,
      }),
    });

    const data = await response.json();

    if (!response.ok) {
      console.error('Resend error:', data);
      return res.status(response.status).json({ 
        error: data.message || 'Erro ao enviar e-mail',
        details: data 
      });
    }

    return res.status(200).json({ success: true, id: data.id });
  } catch (error) {
    console.error('Send invite error:', error);
    return res.status(500).json({ error: 'Erro interno ao enviar convite.' });
  }
}
