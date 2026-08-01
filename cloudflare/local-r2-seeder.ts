type LocalR2Bucket = {
  head(key: string): Promise<{ size: number } | null>;
  put(
    key: string,
    value: ReadableStream | ArrayBuffer | ArrayBufferView | string | null,
    options?: { httpMetadata?: { contentType?: string } },
  ): Promise<unknown>;
  createMultipartUpload(
    key: string,
    options?: { httpMetadata?: { contentType?: string } },
  ): Promise<LocalR2MultipartUpload>;
  resumeMultipartUpload(key: string, uploadId: string): LocalR2MultipartUpload;
};

type LocalR2MultipartUpload = {
  uploadId: string;
  uploadPart(
    partNumber: number,
    value: ReadableStream | ArrayBuffer | ArrayBufferView | string | null,
  ): Promise<{ partNumber: number; etag: string }>;
  complete(parts: Array<{ partNumber: number; etag: string }>): Promise<unknown>;
  abort(): Promise<void>;
};

type LocalSeederEnv = {
  ASSETS: LocalR2Bucket;
  MATRIX_SEED_TOKEN: string;
};

const KEY_HEADER = 'x-matrix-r2-key';
const TOKEN_HEADER = 'x-matrix-seed-token';
const UPLOAD_ID_HEADER = 'x-matrix-r2-upload-id';
const PART_NUMBER_HEADER = 'x-matrix-r2-part-number';

function objectKey(request: Request): string | null {
  const key = request.headers.get(KEY_HEADER);
  return key && !key.includes('\0') && !key.split('/').some((part) => part === '..') ? key : null;
}

function multipartUpload(request: Request, env: LocalSeederEnv, key: string): LocalR2MultipartUpload | null {
  const uploadId = request.headers.get(UPLOAD_ID_HEADER);
  if (!uploadId || uploadId.length > 1_024 || uploadId.includes('\0') || /[\r\n]/.test(uploadId)) return null;
  try {
    return env.ASSETS.resumeMultipartUpload(key, uploadId);
  } catch {
    return null;
  }
}

export default {
  async fetch(request: Request, env: LocalSeederEnv): Promise<Response> {
    if (request.headers.get(TOKEN_HEADER) !== env.MATRIX_SEED_TOKEN) {
      return new Response('Unauthorized', { status: 401 });
    }
    const url = new URL(request.url);
    if (request.method === 'GET' && url.pathname === '/__matrix__/ready') {
      return Response.json({ ready: true });
    }
    const key = objectKey(request);
    if (!key) {
      return new Response('Invalid object key', { status: 400 });
    }
    if (request.method === 'HEAD' && url.pathname === '/__matrix__/object') {
      const object = await env.ASSETS.head(key);
      return object
        ? new Response(null, { status: 200, headers: { 'Content-Length': String(object.size) } })
        : new Response(null, { status: 404 });
    }
    if (request.method === 'PUT' && url.pathname === '/__matrix__/object') {
      await env.ASSETS.put(key, request.body, {
        httpMetadata: {
          contentType: request.headers.get('content-type') || 'application/octet-stream',
        },
      });
      return new Response(null, { status: 204 });
    }
    if (request.method === 'POST' && url.pathname === '/__matrix__/multipart/start') {
      const upload = await env.ASSETS.createMultipartUpload(key, {
        httpMetadata: {
          contentType: request.headers.get('content-type') || 'application/octet-stream',
        },
      });
      return Response.json({ uploadId: upload.uploadId });
    }
    const upload = multipartUpload(request, env, key);
    if (!upload) return new Response('Invalid multipart upload', { status: 400 });
    if (request.method === 'PUT' && url.pathname === '/__matrix__/multipart/part') {
      const partNumber = Number(request.headers.get(PART_NUMBER_HEADER) || '0');
      if (!Number.isSafeInteger(partNumber) || partNumber < 1 || partNumber > 10_000) {
        return new Response('Invalid multipart part number', { status: 400 });
      }
      const uploaded = await upload.uploadPart(partNumber, request.body);
      return Response.json({ partNumber: uploaded.partNumber, etag: uploaded.etag });
    }
    if (request.method === 'POST' && url.pathname === '/__matrix__/multipart/complete') {
      let parts: Array<{ partNumber: number; etag: string }>;
      try {
        const value = (await request.json()) as { parts?: Array<{ partNumber?: number; etag?: string }> };
        if (
          !Array.isArray(value.parts) ||
          value.parts.length === 0 ||
          value.parts.length > 10_000 ||
          value.parts.some(
            (part, index) =>
              part.partNumber !== index + 1 ||
              typeof part.etag !== 'string' ||
              !part.etag ||
              part.etag.length > 512 ||
              part.etag.includes('\0') ||
              /[\r\n]/.test(part.etag),
          )
        ) {
          return new Response('Invalid multipart completion', { status: 400 });
        }
        parts = value.parts as Array<{ partNumber: number; etag: string }>;
      } catch {
        return new Response('Invalid multipart completion', { status: 400 });
      }
      await upload.complete(parts);
      return new Response(null, { status: 204 });
    }
    if (request.method === 'DELETE' && url.pathname === '/__matrix__/multipart') {
      await upload.abort();
      return new Response(null, { status: 204 });
    }
    return new Response('Not found', { status: 404 });
  },
};
