function normalizarEstruturaModelo(input = {}) {
  const normalized = normalizeJsonValue(input, 0);
  if (!normalized || Array.isArray(normalized) || typeof normalized !== 'object') return {};
  delete normalized._historico_id;
  return normalized;
}

function normalizeJsonValue(value, depth) {
  if (depth > 16 || value == null) return value == null ? null : undefined;
  if (typeof value === 'string') return value.slice(0, 100000);
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
  if (typeof value === 'boolean') return value;
  if (Array.isArray(value)) {
    return value.slice(0, 10000)
      .map((item) => normalizeJsonValue(item, depth + 1))
      .filter((item) => item !== undefined);
  }
  if (typeof value !== 'object') return undefined;

  const normalized = {};
  Object.entries(value).slice(0, 2000).forEach(([key, item]) => {
    if (['__proto__', 'prototype', 'constructor'].includes(key)) return;
    const normalizedItem = normalizeJsonValue(item, depth + 1);
    if (normalizedItem !== undefined) normalized[key] = normalizedItem;
  });
  return normalized;
}

module.exports = { normalizarEstruturaModelo };
