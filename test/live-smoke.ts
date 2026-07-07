import { GraphClient } from '../src/graph/client.js';

async function main(): Promise<void> {
  if ((!process.env.MS_TENANT_ID || !process.env.MS_CLIENT_ID || !process.env.MS_CLIENT_SECRET) && !process.env.MS_ACCESS_TOKEN) {
    console.error('Skipping live smoke test: set MS_TENANT_ID/MS_CLIENT_ID/MS_CLIENT_SECRET or MS_ACCESS_TOKEN.');
    return;
  }

  const client = new GraphClient();
  const skus = await client.get<{ value?: unknown[] }>('/subscribedSkus');
  console.log(JSON.stringify({ ok: true, subscribedSkuCount: skus.data?.value?.length ?? 0 }, null, 2));
}

main().catch(error => {
  console.error(error);
  process.exit(1);
});
