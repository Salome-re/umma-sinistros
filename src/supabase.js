import { createClient } from '@supabase/supabase-js';

const SUPABASE_URL = 'https://fwsqwgflajanhtpuagey.supabase.co';
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImZ3c3F3Z2ZsYWphbmh0cHVhZ2V5Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODIxNTE3MDgsImV4cCI6MjA5NzcyNzcwOH0.SbuaYSWSx_ZeTD0W0QCMsftDVry3K-ugfleaEc_u-vo';

export const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

// Verificar se email está autorizado
export async function checkEmail(email) {
  const { data, error } = await supabase
    .from('usuarios_autorizados')
    .select('*')
    .eq('email', email.toLowerCase().trim())
    .single();
  
  if (error || !data) return null;
  
  // Atualizar último acesso
  await supabase
    .from('usuarios_autorizados')
    .update({ ultimo_acesso: new Date().toISOString() })
    .eq('id', data.id);
  
  return data;
}

// Listar todos os usuários (para admin)
export async function listUsers() {
  const { data, error } = await supabase
    .from('usuarios_autorizados')
    .select('*')
    .order('criado_em', { ascending: false });
  
  return error ? [] : data;
}

// Adicionar novo usuário
export async function addUser(email, nome, role, convidadoPor) {
  const { data, error } = await supabase
    .from('usuarios_autorizados')
    .insert([{ 
      email: email.toLowerCase().trim(), 
      nome, 
      role, 
      convidado_por: convidadoPor 
    }])
    .select()
    .single();
  
  if (error) throw new Error(error.message);
  return data;
}

// Remover usuário
export async function removeUser(id) {
  const { error } = await supabase
    .from('usuarios_autorizados')
    .delete()
    .eq('id', id);
  
  if (error) throw new Error(error.message);
}

// Atualizar role do usuário
export async function updateUserRole(id, role) {
  const { error } = await supabase
    .from('usuarios_autorizados')
    .update({ role })
    .eq('id', id);
  
  if (error) throw new Error(error.message);
}
