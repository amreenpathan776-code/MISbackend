/**
 * Format number with Indian commas + 2 decimals.
 */
function formatIndianDecimal(num) {
  if (num === null || num === undefined) return num;

  if (typeof num !== "number") num = Number(num);
  if (isNaN(num)) return num;

  const isNegative = num < 0;
  num = Math.abs(num);

  const numStr = num.toFixed(2);
  const [integer, decimals] = numStr.split(".");

  let formattedInt = integer;
  if (integer.length > 3) {
    const lastThree = integer.slice(-3);
    const rest = integer.slice(0, -3);
    formattedInt =
      rest.replace(/\B(?=(\d{2})+(?!\d))/g, ",") + "," + lastThree;
  }

  return (isNegative ? "-" : "") + formattedInt + "." + decimals;
}

/**
 * Decide if a meaning SHOULD NOT receive .00 formatting.
 */
function shouldSkipMeaning(meaning) {
  if (!meaning) return false;
  const m = meaning.toLowerCase();

  if (m.replace(/[\s-]/g, "_") === "branch_code") return true;
  if (m.endsWith("accounts") || m.endsWith("a/cs")) return true;
  if (m.includes("budget")) return true;

  return false;
}

/**
 * Meaning-based formatting for 1 row or entire array.
 */
function formatWithMeanings(data, meaningMap = {}) {
  if (!Array.isArray(data)) return data;

  return data.map((row) => {
    const formatted = {};

    for (const key in row) {
      const value = row[key];
      const meaning = meaningMap[key];

      if (shouldSkipMeaning(meaning)) {
        formatted[key] = value;
        continue;
      }

      if (value !== null && value !== "" && !isNaN(Number(value))) {
        formatted[key] = formatIndianDecimal(Number(value));
      } else {
        formatted[key] = value;
      }
    }

    return formatted;
  });
}

/**
 * Deep formatter (without meanings) — used for generic responses.
 */
function formatNumbersDeep(obj) {
  if (Array.isArray(obj)) {
    return obj.map(formatNumbersDeep);
  }

  if (typeof obj === "object" && obj !== null) {
    const newObj = {};
    for (const key in obj) {
      const value = obj[key];

      // Skip ACs / page / ID-like patterns
      if (typeof value === "string" && /a\/cs|page|\d+\/\d+/i.test(value)) {
        newObj[key] = value;
        continue;
      }

      // Pure numeric → format
      if (typeof value === "number" || (!isNaN(value) && value !== "")) {
        newObj[key] = formatIndianDecimal(Number(value));
      } else {
        newObj[key] = formatNumbersDeep(value);
      }
    }
    return newObj;
  }

  return obj;
}

module.exports = {
  formatIndianDecimal,
  formatWithMeanings,
  shouldSkipMeaning,
  formatNumbersDeep
};

