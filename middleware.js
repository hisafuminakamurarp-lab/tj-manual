// Vercel Edge Middleware — サイト全体にBasic認証をかける
// 認証情報を変えたいときはこのファイルの USER / PASS を書き換えて push する
const USER = 'info-tokyo@ztjas.com';
const PASS = 'tjadmin';

export default function middleware(request) {
  const auth = request.headers.get('authorization');
  if (auth) {
    const [scheme, encoded] = auth.split(' ');
    if (scheme === 'Basic' && encoded) {
      try {
        if (atob(encoded) === `${USER}:${PASS}`) {
          return; // 認証OK → そのまま静的ファイルを配信
        }
      } catch (e) { /* 不正なBase64は未認証扱い */ }
    }
  }
  return new Response('認証が必要です / Authentication required', {
    status: 401,
    headers: {
      'WWW-Authenticate': 'Basic realm="tj-manual"',
      'Content-Type': 'text/plain; charset=utf-8',
    },
  });
}

export const config = {
  matcher: '/(.*)',
};
