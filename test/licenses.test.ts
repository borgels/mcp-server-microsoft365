import { describe, expect, it } from 'vitest';
import { isGuid, resolveSkuId, type SubscribedSku } from '../src/graph/licenses.js';

const skus: SubscribedSku[] = [
  {
    skuId: '05e9a617-0261-4cee-bb44-138d3ef5d965',
    skuPartNumber: 'SPE_E3',
    prepaidUnits: { enabled: 25 },
    consumedUnits: 20,
  },
  {
    skuId: 'c7df2760-2c81-4ef7-b578-5b5392b571df',
    skuPartNumber: 'ENTERPRISEPREMIUM',
    prepaidUnits: { enabled: 10 },
    consumedUnits: 10,
  },
  {
    skuId: '3b555118-da6a-4418-894f-7df1e2096870',
    skuPartNumber: 'O365_BUSINESS_ESSENTIALS',
    prepaidUnits: { enabled: 5 },
    consumedUnits: 1,
  },
];

describe('resolveSkuId', () => {
  it('detects GUIDs', () => {
    expect(isGuid('05e9a617-0261-4cee-bb44-138d3ef5d965')).toBe(true);
    expect(isGuid('E3')).toBe(false);
  });

  it('matches a SKU GUID directly', () => {
    const result = resolveSkuId(skus, '05e9a617-0261-4cee-bb44-138d3ef5d965');
    expect(result).toMatchObject({ skuPartNumber: 'SPE_E3', matchedBy: 'guid', available: 5 });
  });

  it('returns undefined for a GUID the tenant does not own', () => {
    expect(resolveSkuId(skus, '00000000-0000-0000-0000-000000000000')).toBeUndefined();
  });

  it('matches a skuPartNumber case-insensitively', () => {
    const result = resolveSkuId(skus, 'enterprisepremium');
    expect(result).toMatchObject({ skuPartNumber: 'ENTERPRISEPREMIUM', matchedBy: 'partNumber', available: 0 });
  });

  it('resolves friendly names to an owned SKU part number', () => {
    // E3 -> ['ENTERPRISEPACK', 'SPE_E3']; only SPE_E3 is owned.
    const result = resolveSkuId(skus, 'E3');
    expect(result).toMatchObject({ skuPartNumber: 'SPE_E3', matchedBy: 'friendlyName' });
  });

  it('resolves friendly names with punctuation/spacing', () => {
    const result = resolveSkuId(skus, 'Business Basic');
    expect(result).toMatchObject({ skuPartNumber: 'O365_BUSINESS_ESSENTIALS', available: 4 });
  });

  it('computes availability from prepaid enabled minus consumed', () => {
    const result = resolveSkuId(skus, 'SPE_E3');
    expect(result).toMatchObject({ enabled: 25, consumed: 20, available: 5 });
  });

  it('returns undefined when nothing matches', () => {
    expect(resolveSkuId(skus, 'NOT_A_LICENSE')).toBeUndefined();
  });
});
