function getCookie(request, name) {
  const header = request.headers.get('cookie') || ''
  for (const part of header.split(';')) {
    const [k, ...v] = part.trim().split('=')
    if (k === name) return v.join('=')
  }
  return null
}

async function verifyToken(token) {
  const secret = process.env.SESSION_SECRET
  const passwordHash = process.env.PASSWORD_HASH
  if (!secret || !passwordHash || !token) return false

  const key = await crypto.subtle.importKey(
    'raw',
    new TextEncoder().encode(secret),
    { name: 'HMAC', hash: 'SHA-256' },
    false,
    ['sign']
  )
  const sig = await crypto.subtle.sign('HMAC', key, new TextEncoder().encode(passwordHash))
  const expected = [...new Uint8Array(sig)].map(b => b.toString(16).padStart(2, '0')).join('')

  if (expected.length !== token.length) return false
  let diff = 0
  for (let i = 0; i < expected.length; i++) {
    diff |= expected.charCodeAt(i) ^ token.charCodeAt(i)
  }
  return diff === 0
}

export default async function middleware(request) {
  const { pathname } = new URL(request.url)
  if (!pathname.startsWith('/ces-2026-analisi')) return
  if (pathname === '/ces-2026-analisi/login') return

  const token = getCookie(request, 'ces_auth')
  const valid = await verifyToken(token)

  if (!valid) {
    return Response.redirect(new URL('/ces-2026-analisi/login', request.url), 302)
  }
}

export const config = {
  matcher: ['/ces-2026-analisi', '/ces-2026-analisi/(.*)'],
}
