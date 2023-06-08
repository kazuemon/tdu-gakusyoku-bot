export const getPdfData = async (url: string) => {
  const res = await fetch(url);
  if (res.status !== 200) {
    console.error(`Failed to fetch PDF data`);
    return null;
  }
  if (res.headers.get('content-type') !== 'application/pdf') {
    console.error(`Invalid content type: ${res.headers.get('content-type')}`);
    return null;
  }
  const lastModifiedDate = res.headers.get('last-modified');
  const etag = res.headers.get('etag');
  const buffer = Buffer.from(await res.arrayBuffer());
  return { buffer, etag, lastModifiedDate };
}

export const pdf2Excel = async () => {
  // WIP
}