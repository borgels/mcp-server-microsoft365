export interface SubscribedSku {
  skuId: string;
  skuPartNumber: string;
  prepaidUnits?: {
    enabled?: number;
    suspended?: number;
    warning?: number;
    lockedOut?: number;
  };
  consumedUnits?: number;
}

export interface SkuResolution {
  skuId: string;
  skuPartNumber: string;
  enabled: number;
  consumed: number;
  available: number;
  matchedBy: 'guid' | 'partNumber' | 'friendlyName';
}

/**
 * Friendly license aliases mapped to one or more Microsoft `skuPartNumber`
 * candidates, tried in order. Many friendly names cover both the Office 365 and
 * the Microsoft 365 packaging (for example E3 maps to both `ENTERPRISEPACK` and
 * `SPE_E3`), so the first candidate that a tenant actually owns wins.
 */
export const FRIENDLY_LICENSE_NAMES: Record<string, string[]> = {
  E1: ['STANDARDPACK', 'SPE_E1'],
  E3: ['ENTERPRISEPACK', 'SPE_E3'],
  E5: ['ENTERPRISEPREMIUM', 'SPE_E5'],
  F1: ['DESKLESSPACK', 'SPE_F1'],
  F3: ['SPE_F1'],
  OFFICE_365_E1: ['STANDARDPACK'],
  OFFICE_365_E3: ['ENTERPRISEPACK'],
  OFFICE_365_E5: ['ENTERPRISEPREMIUM'],
  MICROSOFT_365_E3: ['SPE_E3'],
  MICROSOFT_365_E5: ['SPE_E5'],
  MICROSOFT_365_F3: ['SPE_F1'],
  BUSINESS_BASIC: ['O365_BUSINESS_ESSENTIALS'],
  BUSINESS_STANDARD: ['O365_BUSINESS_PREMIUM'],
  BUSINESS_PREMIUM: ['SPB'],
  EXCHANGE_ONLINE_PLAN_1: ['EXCHANGESTANDARD'],
  EXCHANGE_ONLINE_PLAN_2: ['EXCHANGEENTERPRISE'],
  EXCHANGE_ONLINE_KIOSK: ['EXCHANGEDESKLESS'],
  POWER_BI_PRO: ['POWER_BI_PRO'],
  TEAMS_ESSENTIALS: ['Teams_Ess'],
  ENTRA_ID_P1: ['AAD_PREMIUM'],
  ENTRA_ID_P2: ['AAD_PREMIUM_P2'],
};

const GUID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

export function isGuid(value: string): boolean {
  return GUID_PATTERN.test(value.trim());
}

/**
 * Resolve a friendly name, `skuPartNumber`, or SKU GUID against the tenant's
 * `subscribedSkus`. Returns the resolved SKU id plus availability
 * (`prepaidUnits.enabled - consumedUnits`), or `undefined` when nothing in the
 * tenant matches.
 */
export function resolveSkuId(
  subscribedSkus: readonly SubscribedSku[],
  nameOrPartNumberOrGuid: string,
): SkuResolution | undefined {
  const input = nameOrPartNumberOrGuid.trim();
  if (!input) {
    return undefined;
  }

  if (isGuid(input)) {
    const byGuid = subscribedSkus.find(sku => sku.skuId.toLowerCase() === input.toLowerCase());
    return byGuid ? toResolution(byGuid, 'guid') : undefined;
  }

  const byPartNumber = subscribedSkus.find(
    sku => sku.skuPartNumber.toLowerCase() === input.toLowerCase(),
  );
  if (byPartNumber) {
    return toResolution(byPartNumber, 'partNumber');
  }

  const candidates = FRIENDLY_LICENSE_NAMES[normalizeKey(input)];
  if (candidates) {
    for (const partNumber of candidates) {
      const match = subscribedSkus.find(
        sku => sku.skuPartNumber.toLowerCase() === partNumber.toLowerCase(),
      );
      if (match) {
        return toResolution(match, 'friendlyName');
      }
    }
  }

  return undefined;
}

function toResolution(sku: SubscribedSku, matchedBy: SkuResolution['matchedBy']): SkuResolution {
  const enabled = sku.prepaidUnits?.enabled ?? 0;
  const consumed = sku.consumedUnits ?? 0;
  return {
    skuId: sku.skuId,
    skuPartNumber: sku.skuPartNumber,
    enabled,
    consumed,
    available: enabled - consumed,
    matchedBy,
  };
}

function normalizeKey(value: string): string {
  return value
    .trim()
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
}
