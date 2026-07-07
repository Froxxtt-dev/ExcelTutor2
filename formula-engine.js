/* ============================================================
   Tiny formula engine for the practice sandbox.
   Supports: + - * / ^, comparisons, parentheses, cell refs (A1),
   ranges (A1:A3), and SUM/AVERAGE/COUNT/MIN/MAX/IF/AND/OR/ROUND/
   CONCAT/ABS. Deliberately small & dependency-free so it stays
   fast on low-end mobile devices.
   ============================================================ */

const AGGREGATE_FNS = new Set(['SUM', 'AVERAGE', 'COUNT', 'MIN', 'MAX']);

function feTokenize(expr) {
  const re = /\s*(\$?[A-Fa-f]\$?\d+(?::\$?[A-Fa-f]\$?\d+)?|\d+\.?\d*|"[^"]*"|<=|>=|<>|[+\-*/^(),<>=]|[A-Za-z_][A-Za-z0-9_]*)\s*/g;
  const tokens = [];
  let m;
  let lastIndex = 0;
  while ((m = re.exec(expr)) !== null) {
    if (m.index !== lastIndex) {
      // Unrecognized character sequence — stop rather than silently skipping.
      throw new Error('Unrecognized syntax near "' + expr.slice(lastIndex, m.index) + '"');
    }
    tokens.push(m[1]);
    lastIndex = re.lastIndex;
  }
  if (lastIndex !== expr.length) {
    throw new Error('Unrecognized syntax near "' + expr.slice(lastIndex) + '"');
  }
  return tokens;
}

function isCellRef(tok) {
  return /^\$?[A-Fa-f]\$?\d+$/.test(tok);
}
function isRange(tok) {
  return /^\$?[A-Fa-f]\$?\d+:\$?[A-Fa-f]\$?\d+$/.test(tok);
}
function cleanRef(tok) {
  return tok.replace(/\$/g, '').toUpperCase();
}

class FormulaEngine {
  constructor(sheetGetter) {
    // sheetGetter(ref) -> raw string/number stored in that cell (or undefined)
    this.getRaw = sheetGetter;
    this.cache = new Map();
    this.visiting = new Set();
  }

  valueAt(ref) {
    ref = cleanRef(ref);
    if (this.cache.has(ref)) return this.cache.get(ref);
    if (this.visiting.has(ref)) throw new Error('#CYCLE!');
    const raw = this.getRaw(ref);
    let val;
    if (raw === undefined || raw === null || raw === '') {
      val = 0;
    } else if (typeof raw === 'string' && raw.trim().startsWith('=')) {
      this.visiting.add(ref);
      val = this.evaluate(raw.trim().slice(1));
      this.visiting.delete(ref);
    } else {
      const n = parseFloat(raw);
      val = Number.isNaN(n) ? String(raw) : n;
    }
    this.cache.set(ref, val);
    return val;
  }

  rangeValues(rangeTok) {
    const [a, b] = cleanRef(rangeTok).split(':');
    const colOf = (r) => r.charCodeAt(0);
    const rowOf = (r) => parseInt(r.slice(1), 10);
    const c1 = colOf(a), c2 = colOf(b), r1 = rowOf(a), r2 = rowOf(b);
    const out = [];
    for (let c = Math.min(c1, c2); c <= Math.max(c1, c2); c++) {
      for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r++) {
        out.push(this.valueAt(String.fromCharCode(c) + r));
      }
    }
    return out;
  }

  evaluate(expr) {
    this.tokens = feTokenize(expr);
    this.pos = 0;
    const result = this.parseComparison();
    if (this.pos < this.tokens.length) {
      throw new Error('Unexpected token "' + this.tokens[this.pos] + '"');
    }
    return result;
  }

  peek() { return this.tokens[this.pos]; }
  next() { return this.tokens[this.pos++]; }

  parseComparison() {
    let left = this.parseAddSub();
    while (['=', '<>', '<', '>', '<=', '>='].includes(this.peek())) {
      const op = this.next();
      const right = this.parseAddSub();
      const l = Array.isArray(left) ? left[0] : left;
      const r = Array.isArray(right) ? right[0] : right;
      switch (op) {
        case '=': left = l === r; break;
        case '<>': left = l !== r; break;
        case '<': left = l < r; break;
        case '>': left = l > r; break;
        case '<=': left = l <= r; break;
        case '>=': left = l >= r; break;
      }
    }
    return left;
  }

  parseAddSub() {
    let left = this.parseTerm();
    while (this.peek() === '+' || this.peek() === '-') {
      const op = this.next();
      const right = this.parseTerm();
      const l = Array.isArray(left) ? left[0] : left;
      const r = Array.isArray(right) ? right[0] : right;
      left = op === '+' ? Number(l) + Number(r) : Number(l) - Number(r);
    }
    return left;
  }

  parseTerm() {
    let left = this.parsePower();
    while (this.peek() === '*' || this.peek() === '/') {
      const op = this.next();
      const right = this.parsePower();
      const l = Array.isArray(left) ? left[0] : left;
      const r = Array.isArray(right) ? right[0] : right;
      left = op === '*' ? Number(l) * Number(r) : Number(l) / Number(r);
    }
    return left;
  }

  parsePower() {
    let left = this.parseUnary();
    while (this.peek() === '^') {
      this.next();
      const right = this.parseUnary();
      left = Math.pow(Number(left), Number(right));
    }
    return left;
  }

  parseUnary() {
    if (this.peek() === '-') {
      this.next();
      return -Number(this.parseUnary());
    }
    return this.parsePrimary();
  }

  parsePrimary() {
    const tok = this.peek();
    if (tok === undefined) throw new Error('Unexpected end of formula');

    if (tok === '(') {
      this.next();
      const val = this.parseComparison();
      if (this.next() !== ')') throw new Error('Expected ")"');
      return val;
    }

    if (/^\d/.test(tok)) {
      this.next();
      return parseFloat(tok);
    }

    if (tok.startsWith('"')) {
      this.next();
      return tok.slice(1, -1);
    }

    if (isRange(tok)) {
      this.next();
      return this.rangeValues(tok);
    }

    if (isCellRef(tok)) {
      this.next();
      return this.valueAt(tok);
    }

    // Function call: NAME(args...)
    if (/^[A-Za-z_]/.test(tok)) {
      this.next();
      const name = tok.toUpperCase();
      if (this.peek() !== '(') throw new Error('Expected "(" after ' + name);
      this.next();
      const args = [];
      if (this.peek() !== ')') {
        args.push(this.parseComparison());
        while (this.peek() === ',') {
          this.next();
          args.push(this.parseComparison());
        }
      }
      if (this.next() !== ')') throw new Error('Expected ")" to close ' + name);
      return this.callFn(name, args);
    }

    throw new Error('Unexpected token "' + tok + '"');
  }

  callFn(name, args) {
    if (AGGREGATE_FNS.has(name)) {
      const pool = [];
      args.forEach((a) => {
        if (Array.isArray(a)) pool.push(...a);
        else pool.push(a);
      });
      const nums = pool.map(Number).filter((n) => !Number.isNaN(n));
      switch (name) {
        case 'SUM': return nums.reduce((s, n) => s + n, 0);
        case 'AVERAGE': return nums.length ? nums.reduce((s, n) => s + n, 0) / nums.length : 0;
        case 'COUNT': return nums.length;
        case 'MIN': return nums.length ? Math.min(...nums) : 0;
        case 'MAX': return nums.length ? Math.max(...nums) : 0;
      }
    }
    const scalar = (v) => (Array.isArray(v) ? v[0] : v);
    switch (name) {
      case 'IF': {
        const cond = scalar(args[0]);
        const truthy = cond === true || (typeof cond === 'number' && cond !== 0) || (typeof cond === 'string' && cond.length > 0 && cond !== '0');
        return truthy ? args[1] : args[2];
      }
      case 'AND': return args.every((a) => !!scalar(a));
      case 'OR': return args.some((a) => !!scalar(a));
      case 'NOT': return !scalar(args[0]);
      case 'ROUND': {
        const n = Number(scalar(args[0]));
        const d = Number(scalar(args[1] ?? 0));
        const f = Math.pow(10, d);
        return Math.round(n * f) / f;
      }
      case 'ABS': return Math.abs(Number(scalar(args[0])));
      case 'CONCAT':
      case 'CONCATENATE':
        return args.map((a) => String(scalar(a))).join('');
      default:
        throw new Error('Unsupported function ' + name);
    }
  }
}

function formatFormulaResult(val) {
  if (val === true) return 'TRUE';
  if (val === false) return 'FALSE';
  if (typeof val === 'number') {
    if (!Number.isFinite(val)) return '#ERROR!';
    return Math.abs(val - Math.round(val)) < 1e-9 ? String(Math.round(val)) : String(Math.round(val * 10000) / 10000);
  }
  return String(val);
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = { FormulaEngine, formatFormulaResult };
}
