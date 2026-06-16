const crypto = require('crypto')

module.exports = async (req, res) => {
  if (req.method !== 'POST') return res.status(405).end()

  let body = ''
  for await (const chunk of req) body += chunk
  const params = new URLSearchParams(body)
  const password = params.get('password') || ''

  const submitted = crypto.createHash('sha256').update(password).digest('hex')
  const stored = process.env.PASSWORD_HASH || ''

  const lengthsMatch = submitted.length === stored.length
  const safe = Buffer.alloc(submitted.length, submitted)
  const ref = Buffer.alloc(stored.length || submitted.length, stored)
  const match = lengthsMatch && crypto.timingSafeEqual(safe, ref)

  if (!match) {
    return res.redirect(302, '/login?error=1')
  }

  const token = crypto
    .createHmac('sha256', process.env.SESSION_SECRET)
    .update(stored)
    .digest('hex')

  res.setHeader('Set-Cookie',
    `ces_auth=${token}; HttpOnly; Secure; SameSite=Lax; Path=/ces-2026-analisi; Max-Age=28800`
  )
  res.redirect(302, '/ces-2026-analisi')
}
