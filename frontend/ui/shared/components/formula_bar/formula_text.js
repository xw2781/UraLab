/*
===============================================================================
Formula Text
How ArcRho reads a cell formula. Every surface that shows one — the DFM Ratios
formula bar, the Dataset Viewer linked-cell editor — tokenises it here, so a
reference is recognised the same way wherever it is displayed. Rendering is left
to each surface, which knows which token kinds it can offer actions for.
===============================================================================
*/

/**
 * Tokenise a formula string into typed segments.
 * Recognises Excel refs, quoted row references, bracketed references,
 * operators, and plain text.
 */
export function tokenizeFormula(rawText) {
  const text = String(rawText || "").trim();
  if (!text) return [];

  // Ensure leading '='
  const normalizedText = text.startsWith("=") ? text : "=" + text;
  let remaining = normalizedText;
  let offset = 0;
  const tokens = [];

  const pushToken = (type, tokenText) => {
    tokens.push({ type, text: tokenText, start: offset, end: offset + tokenText.length });
    offset += tokenText.length;
    remaining = remaining.slice(tokenText.length);
  };

  while (remaining.length > 0) {
    // Excel ref: 'dir\[file.xlsx]Sheet'!A1 or a range such as ...!A1:C3
    const xlMatch = /^'([^[]*)\[([^\]]+)\]([^'!]+)'!\$?[A-Z]+\$?[0-9]+(?::\$?[A-Z]+\$?[0-9]+)?/i.exec(remaining);
    if (xlMatch) {
      pushToken("excel", xlMatch[0]);
      continue;
    }
    // Quoted row reference: "Some Label" or 'Some Label'
    const quotedMatch = /^(["'])(.+?)\1/.exec(remaining);
    if (quotedMatch) {
      pushToken("ref", quotedMatch[0]);
      continue;
    }
    // Dataset names and coordinates: preserve everything inside each complete
    // bracket pair verbatim so negative indices and operator-like characters
    // remain part of the reference rather than formula operators.
    const bracketMatch = /^\[[^\]]*\]/.exec(remaining);
    if (bracketMatch) {
      pushToken("bracket", bracketMatch[0]);
      continue;
    }
    // Operator
    const opMatch = /^[+\-*/]/.exec(remaining);
    if (opMatch) {
      pushToken("op", opMatch[0]);
      continue;
    }
    // Plain text (one char at a time)
    pushToken("plain", remaining[0]);
  }

  // Merge consecutive plain tokens
  const merged = [];
  for (const tok of tokens) {
    if (tok.type === "plain" && merged.length > 0 && merged[merged.length - 1].type === "plain") {
      merged[merged.length - 1].text += tok.text;
      merged[merged.length - 1].end = tok.end;
    } else {
      merged.push({ ...tok });
    }
  }
  for (let index = 0; index < merged.length; index += 1) {
    const token = merged[index];
    if (token.type !== "bracket") continue;
    let nextIndex = index + 1;
    while (
      merged[nextIndex]?.type === "plain"
      && !String(merged[nextIndex].text || "").trim()
    ) nextIndex += 1;
    if (merged[nextIndex]?.type !== "bracket") continue;
    const datasetName = token.text.slice(1, -1).trim();
    const coordinateLabel = merged[nextIndex].text.slice(1, -1).trim();
    if (!datasetName || !coordinateLabel) continue;
    token.datasetName = datasetName;
    token.datasetCoordinateLabel = coordinateLabel;
    merged[nextIndex].datasetCoordinate = true;
    index = nextIndex;
  }
  return merged;
}

/**
 * Format a raw formula string with proper spacing around operators
 * and ensure leading '='. Does not alter content inside Excel refs
 * or bracketed/quoted references.
 */
export function formatFormulaText(rawText) {
  const tokens = tokenizeFormula(rawText);
  if (!tokens.length) return String(rawText || "").trim();
  let out = "";
  for (const tok of tokens) {
    if (tok.type === "op") {
      out = out.replace(/\s+$/, "");
      out += " " + tok.text + " ";
    } else if (tok.type === "plain") {
      out += tok.text.trim();
    } else {
      out += tok.text;
    }
  }
  const formatted = out.replace(/\s+$/, "");
  if (formatted.startsWith("=")) return `= ${formatted.slice(1).trimStart()}`;
  return formatted;
}

/**
 * Drop the ROUND(...) a formula wraps a term in, keeping the term itself.
 *
 * A reserve review's User Entry formula names its own precision - an average
 * row is taken at four decimals, and the formula says so - but a reader
 * scanning the Ratios tab wants the factor, not the bookkeeping. The stored
 * text is untouched: this is only what the rendered formula shows, so editing
 * the cell still reveals the ROUND that is really there.
 *
 * Only a complete ROUND(term, digits) is unwrapped. Anything else - a ROUND
 * still being typed, or one whose parentheses do not balance - is left alone,
 * so a half-written formula reads back exactly as it was entered.
 */
export function stripRoundWrappers(rawText) {
  const text = String(rawText || "");
  let out = "";
  let i = 0;
  while (i < text.length) {
    const match = /^round\s*\(/i.exec(text.slice(i));
    if (!match) {
      out += text[i];
      i += 1;
      continue;
    }
    const open = i + match[0].length;
    let depth = 1;
    let j = open;
    let lastComma = -1;
    for (; j < text.length && depth > 0; j += 1) {
      const ch = text[j];
      if (ch === "(") depth += 1;
      else if (ch === ")") depth -= 1;
      else if (ch === "," && depth === 1) lastComma = j;
    }
    // Unbalanced, or no digits argument: leave the call as the author wrote it.
    if (depth !== 0 || lastComma < 0 || !/^\s*\d+\s*$/.test(text.slice(lastComma + 1, j - 1))) {
      out += text[i];
      i += 1;
      continue;
    }
    out += stripRoundWrappers(text.slice(open, lastComma)).trim();
    i = j;
  }
  return out;
}
