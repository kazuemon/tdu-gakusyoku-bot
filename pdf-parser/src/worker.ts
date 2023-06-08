import { Buffer } from 'node:buffer';
import { getPdfData } from './pdf';

const handlers: ExportedHandler<Env> = {
	async scheduled(controller, env, ctx) {
		const data = await getPdfData(env.PDF_URL);
		if (data === null) return;
		const { buffer, etag, lastModifiedDate } = data;
	}
};

export default handlers;