type LocalR2Bucket = {
  put(
    key: string,
    value: ReadableStream | ArrayBuffer | ArrayBufferView | string | null,
    options?: { httpMetadata?: { contentType?: string } },
  ): Promise<unknown>;
};

type LocalSeederEnv = {
  ASSETS: LocalR2Bucket;
  MATRIX_SEED_TOKEN: string;
};

const KEY_HEADER = 'x-matrix-r2-key';
const TOKEN_HEADER = 'x-matrix-seed-token';

export default {
  async fetch(request: Request, env: LocalSeederEnv): Promise<Response> {
    if (request.headers.get(TOKEN_HEADER) !== env.MATRIX_SEED_TOKEN) {
      return new Response('Unauthorized', { status: 401 });
    }
    const url = new URL(request.url);
    if (request.method === 'GET' && url.pathname === '/__matrix__/ready') {
      return Response.json({ ready: true });
    }
    if (request.method !== 'PUT' || url.pathname !== '/__matrix__/object') {
      return new Response('Not found', { status: 404 });
    }
    const key = request.headers.get(KEY_HEADER);
    if (!key || key.includes('\0') || key.split('/').some((part) => part === '..')) {
      return new Response('Invalid object key', { status: 400 });
    }
    await env.ASSETS.put(key, request.body, {
      httpMetadata: {
        contentType: request.headers.get('content-type') || 'application/octet-stream',
      },
    });
    return new Response(null, { status: 204 });
  },
};
