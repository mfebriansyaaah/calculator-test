/**
 * Class Panel Properties is used to calculate the properties of panel CLT Layup.
 * Panel properties can calculate
 *  - Shear Analogy Method
 *  - Gamma Method
 * 
 * How to use : 
 * calculate(CLTLayup) => PanelProperties
 */

import { unzipSync } from 'fflate';
import { PanelPropertiesType } from '../type/panel-properties-type.js';

class PanelProperties {
    static #excelModel = null;

    static setExcelModel(excelModel) {
        PanelProperties.#excelModel = excelModel;
    }

    static async loadFromXlsxBuffer(xlsxBuffer) {
        const zip = unzipSync(new Uint8Array(xlsxBuffer));
        const sheetXmlBytes = zip['xl/worksheets/sheet1.xml'];
        const sharedXmlBytes = zip['xl/sharedStrings.xml'];
        if (!sheetXmlBytes || !sharedXmlBytes) {
            throw new Error('Excel model tidak lengkap (sheet1.xml/sharedStrings.xml tidak ditemukan).');
        }

        const decoder = new TextDecoder('utf-8');
        const sheetXml = decoder.decode(sheetXmlBytes);
        const sharedXml = decoder.decode(sharedXmlBytes);
        const model = ExcelModel.fromXml({ sheetXml, sharedXml });
        PanelProperties.setExcelModel(model);
        return model;
    }

    static calculate(cltLayup) {
        return new PanelProperties().calculate(cltLayup);
    }

    calculate(cltLayup) {
        const validation = cltLayup.validate();
        if (!validation.ok) {
            throw new Error(validation.message);
        }

        if (!PanelProperties.#excelModel) {
            throw new Error('Excel model belum dimuat.');
        }

        if (cltLayup.method === 'gamma') {
            return new GammaMethod(PanelProperties.#excelModel).calculate(cltLayup);
        }
        return new ShearAnalogyMethod(PanelProperties.#excelModel).calculate(cltLayup);
    }
}

class ShearAnalogyMethod extends PanelProperties {
    constructor(excelModel) {
        super();
        this.excel = excelModel;
    }

    calculate(cltLayup) {
        this.excel.resetCache();
        this.excel.setValue('B9', cltLayup.getLayers()[0].grade.name);
        this.excel.setValue('D9', cltLayup.getLayerCount());
        this.excel.setValue('F9', cltLayup.getLayers()[0].thicknessMm);
        this.excel.setValue('H9', cltLayup.lengthM);
        this.excel.setValue('D18', cltLayup.beffMm);

        const eiEff = this.excel.evaluate('H53');

        const cells = {
            ...this.excel.pickRect('Z36', 'AF46'),
            ...this.excel.pickRect('Z49', 'AF59'),
            ...this.excel.pickRect('B52', 'K53')
        };

        return new PanelPropertiesType({
            method: 'shear-analogy',
            beffMm: cltLayup.beffMm,
            lengthM: cltLayup.lengthM,
            totalThicknessMm: cltLayup.getTotalThicknessMm(),
            neutralAxisYmm: 0,
            eiEff: eiEff || 0
        }).setExcel({ cells, range: null });
    }
}

class GammaMethod extends PanelProperties {
    constructor(excelModel) {
        super();
        this.excel = excelModel;
    }

    calculate(cltLayup) {
        this.excel.resetCache();
        this.excel.setValue('B9', cltLayup.getLayers()[0].grade.name);
        this.excel.setValue('D9', cltLayup.getLayerCount());
        this.excel.setValue('F9', cltLayup.getLayers()[0].thicknessMm);
        this.excel.setValue('H9', cltLayup.lengthM);
        this.excel.setValue('D18', cltLayup.beffMm);

        const eiEff = this.excel.evaluate('H53');
        const cells = {
            ...this.excel.pickRect('Z36', 'AF46'),
            ...this.excel.pickRect('Z67', 'AM80'),
            ...this.excel.pickRect('B52', 'K53')
        };

        return new PanelPropertiesType({
            method: 'gamma',
            beffMm: cltLayup.beffMm,
            lengthM: cltLayup.lengthM,
            totalThicknessMm: cltLayup.getTotalThicknessMm(),
            neutralAxisYmm: 0,
            eiEff: eiEff || 0
        }).setExcel({ cells, range: null });
    }
}

class ExcelModel {
    constructor({ cells, sharedStrings }) {
        this.cells = cells;
        this.sharedStrings = sharedStrings;
        this.overrides = {};
        this.cache = {};
        this.textToCell = null;
    }

    static fromXml({ sheetXml, sharedXml }) {
        const sharedStrings = ExcelModel.#parseSharedStrings(sharedXml);
        const cells = ExcelModel.#parseSheetCells(sheetXml);
        return new ExcelModel({ cells, sharedStrings });
    }

    resetCache() {
        this.cache = {};
        this.textToCell = null;
    }

    setValue(ref, value) {
        this.overrides[ExcelModel.#normalizeRef(ref)] = value;
        this.cache = {};
        this.textToCell = null;
    }

    evaluate(ref) {
        const key = ExcelModel.#normalizeRef(ref);
        return this.#evalCell(key, new Set());
    }

    evaluateRange(range) {
        if (!range) return {};
        const out = {};
        for (const ref of this.#iterateRange(range)) {
            out[ref] = this.evaluate(ref);
        }
        return out;
    }

    pickCells(refs) {
        const out = {};
        for (const ref of refs) {
            out[ExcelModel.#normalizeRef(ref)] = this.evaluate(ref);
        }
        return out;
    }

    pickRect(startRef, endRef) {
        const a = ExcelModel.#splitRef(ExcelModel.#normalizeRef(startRef));
        const b = ExcelModel.#splitRef(ExcelModel.#normalizeRef(endRef));
        const r1 = Math.min(a.row, b.row);
        const r2 = Math.max(a.row, b.row);
        const c1 = Math.min(ExcelModel.#colToNum(a.col), ExcelModel.#colToNum(b.col));
        const c2 = Math.max(ExcelModel.#colToNum(a.col), ExcelModel.#colToNum(b.col));

        const out = {};
        for (let r = r1; r <= r2; r++) {
            for (let c = c1; c <= c2; c++) {
                const ref = ExcelModel.#numToCol(c) + r;
                out[ref] = this.evaluate(ref);
            }
        }
        return out;
    }

    rangeFromAnchors({ startText, endText, startCol, endCol }) {
        const startRef = startText ? this.findCellByText(startText) : null;
        const endRef = endText ? this.findCellByText(endText) : null;
        const startRow = startRef ? ExcelModel.#splitRef(startRef).row : 1;
        const endRow = endRef ? ExcelModel.#splitRef(endRef).row - 1 : this.#maxRow();
        return {
            startRow,
            endRow: Math.max(startRow, endRow),
            startCol: startCol || 'B',
            endCol: endCol || 'AM'
        };
    }

    findCellByText(text) {
        if (!this.textToCell) {
            this.textToCell = new Map();
            for (const ref of Object.keys(this.cells)) {
                const cell = this.cells[ref];
                if (cell.t === 's' && cell.v !== '') {
                    const idx = Number(cell.v);
                    const s = this.sharedStrings[idx] || '';
                    if (s) this.textToCell.set(s, ref);
                }
            }
        }
        return this.textToCell.get(text) || null;
    }

    #maxRow() {
        let max = 1;
        for (const ref of Object.keys(this.cells)) {
            const { row } = ExcelModel.#splitRef(ref);
            if (row > max) max = row;
        }
        return max;
    }

    *#iterateRange({ startRow, endRow, startCol, endCol }) {
        const sc = ExcelModel.#colToNum(startCol);
        const ec = ExcelModel.#colToNum(endCol);
        for (let r = startRow; r <= endRow; r++) {
            for (let c = sc; c <= ec; c++) {
                yield ExcelModel.#numToCol(c) + r;
            }
        }
    }

    #evalCell(ref, visiting) {
        if (Object.prototype.hasOwnProperty.call(this.overrides, ref)) {
            return this.overrides[ref];
        }

        if (Object.prototype.hasOwnProperty.call(this.cache, ref)) {
            return this.cache[ref];
        }

        const cell = this.cells[ref];
        if (!cell) {
            this.cache[ref] = 0;
            return 0;
        }

        if (!cell.f) {
            const val = this.#decodeCellValue(cell);
            this.cache[ref] = val;
            return val;
        }

        if (cell.f.includes('!') || cell.f.includes('[')) {
            const val = this.#decodeCellValue(cell);
            this.cache[ref] = val;
            return val;
        }

        if (visiting.has(ref)) {
            this.cache[ref] = 0;
            return 0;
        }
        visiting.add(ref);

        const val = this.#evalFormula(cell.f, visiting);
        visiting.delete(ref);
        this.cache[ref] = val;
        return val;
    }

    #decodeCellValue(cell) {
        if (cell.v === '') return 0;
        if (cell.t === 's') {
            const idx = Number(cell.v);
            return this.sharedStrings[idx] || '';
        }
        const n = Number(cell.v);
        if (Number.isFinite(n)) return n;
        return cell.v;
    }

    #evalFormula(formula, visiting) {
        const expr = formula.startsWith('=') ? formula.slice(1) : formula;
        const tokens = ExcelTokenizer.tokenize(expr);
        const parser = new ExcelParser(tokens);
        const ast = parser.parseExpression();
        return ExcelEvaluator.evaluate(ast, {
            getCell: (ref) => this.#evalCell(ExcelModel.#normalizeRef(ref), visiting),
            getRange: (a, b) => this.#getRangeValues(a, b, visiting)
        });
    }

    #getRangeValues(aRef, bRef, visiting) {
        const a = ExcelModel.#splitRef(ExcelModel.#normalizeRef(aRef));
        const b = ExcelModel.#splitRef(ExcelModel.#normalizeRef(bRef));
        const r1 = Math.min(a.row, b.row);
        const r2 = Math.max(a.row, b.row);
        const c1 = Math.min(ExcelModel.#colToNum(a.col), ExcelModel.#colToNum(b.col));
        const c2 = Math.max(ExcelModel.#colToNum(a.col), ExcelModel.#colToNum(b.col));

        const out = [];
        for (let r = r1; r <= r2; r++) {
            const row = [];
            for (let c = c1; c <= c2; c++) {
                const ref = ExcelModel.#numToCol(c) + r;
                row.push(this.#evalCell(ref, visiting));
            }
            out.push(row);
        }
        return out;
    }

    static #normalizeRef(ref) {
        return String(ref || '').replace(/\$/g, '').toUpperCase();
    }

    static #splitRef(ref) {
        const m = /^([A-Z]+)(\d+)$/.exec(ref);
        if (!m) return { col: 'A', row: 1 };
        return { col: m[1], row: Number(m[2]) };
    }

    static #colToNum(col) {
        let n = 0;
        for (const ch of col) n = n * 26 + (ch.charCodeAt(0) - 64);
        return n;
    }

    static #numToCol(n) {
        let s = '';
        let x = n;
        while (x > 0) {
            const r = (x - 1) % 26;
            s = String.fromCharCode(65 + r) + s;
            x = Math.floor((x - 1) / 26);
        }
        return s;
    }

    static #parseSharedStrings(xml) {
        const siRe = /<si[\s\S]*?<\/si>/g;
        const tRe = /<t[^>]*>([\s\S]*?)<\/t>/g;
        const shared = [];
        let sm;
        while ((sm = siRe.exec(xml))) {
            const si = sm[0];
            let text = '';
            let tm;
            while ((tm = tRe.exec(si))) text += ExcelModel.#unescapeXml(tm[1]);
            shared.push(text);
        }
        return shared;
    }

    static #parseSheetCells(xml) {
        const cellRe = /<c\b([^>]*?)r="([^"]+)"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
        const cells = {};
        let m;
        while ((m = cellRe.exec(xml))) {
            const fullTag = `<c${m[1]}r="${m[2]}"${m[3]}>`;
            const ref = ExcelModel.#normalizeRef(m[2]);
            const body = m[4] || '';
            const tMatch = /t="([^"]+)"/.exec(fullTag);
            const t = tMatch ? tMatch[1] : '';
            const vMatch = /<v>([\s\S]*?)<\/v>/.exec(body);
            const fMatch = /<f[^>]*>([\s\S]*?)<\/f>/.exec(body);
            const v = vMatch ? ExcelModel.#unescapeXml(vMatch[1]) : '';
            const f = fMatch ? ExcelModel.#unescapeXml(fMatch[1]) : '';
            cells[ref] = { t, v, f };
        }
        return cells;
    }

    static #unescapeXml(s) {
        return String(s || '')
            .replace(/&lt;/g, '<')
            .replace(/&gt;/g, '>')
            .replace(/&amp;/g, '&')
            .replace(/&quot;/g, '"')
            .replace(/&#39;/g, "'");
    }
}

class ExcelTokenizer {
    static tokenize(input) {
        const s = String(input || '');
        const tokens = [];
        let i = 0;

        const isDigit = (ch) => ch >= '0' && ch <= '9';
        const isAlpha = (ch) => (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z') || ch === '_' || ch === '$';

        while (i < s.length) {
            const ch = s[i];
            if (ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r') {
                i++;
                continue;
            }

            if (ch === '"') {
                let j = i + 1;
                let out = '';
                while (j < s.length) {
                    const c = s[j];
                    if (c === '"') break;
                    out += c;
                    j++;
                }
                tokens.push({ type: 'string', value: out });
                i = j + 1;
                continue;
            }

            if (isDigit(ch) || (ch === '.' && isDigit(s[i + 1] || ''))) {
                let j = i;
                let num = '';
                while (j < s.length && (isDigit(s[j]) || s[j] === '.')) {
                    num += s[j];
                    j++;
                }
                tokens.push({ type: 'number', value: Number(num) });
                i = j;
                continue;
            }

            const two = s.slice(i, i + 2);
            if (two === '>=' || two === '<=' || two === '<>') {
                tokens.push({ type: 'op', value: two });
                i += 2;
                continue;
            }

            if ('+-*/^&=><'.includes(ch)) {
                tokens.push({ type: 'op', value: ch });
                i++;
                continue;
            }

            if (ch === '(' || ch === ')') {
                tokens.push({ type: 'paren', value: ch });
                i++;
                continue;
            }

            if (ch === ',') {
                tokens.push({ type: 'comma', value: ch });
                i++;
                continue;
            }

            if (ch === ':') {
                tokens.push({ type: 'colon', value: ch });
                i++;
                continue;
            }

            if (isAlpha(ch)) {
                let j = i;
                let ident = '';
                while (j < s.length && (isAlpha(s[j]) || isDigit(s[j]) || s[j] === '.' || s[j] === '!' || s[j] === '\'')) {
                    ident += s[j];
                    j++;
                }
                tokens.push({ type: 'ident', value: ident });
                i = j;
                continue;
            }

            i++;
        }

        tokens.push({ type: 'eof', value: '' });
        return tokens;
    }
}

class ExcelParser {
    constructor(tokens) {
        this.tokens = tokens;
        this.pos = 0;
    }

    peek() {
        return this.tokens[this.pos] || { type: 'eof', value: '' };
    }

    consume() {
        const t = this.peek();
        this.pos++;
        return t;
    }

    match(type, value) {
        const t = this.peek();
        if (t.type !== type) return false;
        if (value != null && t.value !== value) return false;
        this.pos++;
        return true;
    }

    parseExpression() {
        return this.parseComparison();
    }

    parseComparison() {
        let node = this.parseAdditive();
        while (this.peek().type === 'op' && ['=', '<>', '<', '>', '<=', '>='].includes(this.peek().value)) {
            const op = this.consume().value;
            const right = this.parseAdditive();
            node = { type: 'bin', op, left: node, right };
        }
        return node;
    }

    parseAdditive() {
        let node = this.parseMultiplicative();
        while (this.peek().type === 'op' && ['+', '-', '&'].includes(this.peek().value)) {
            const op = this.consume().value;
            const right = this.parseMultiplicative();
            node = { type: 'bin', op, left: node, right };
        }
        return node;
    }

    parseMultiplicative() {
        let node = this.parsePower();
        while (this.peek().type === 'op' && ['*', '/'].includes(this.peek().value)) {
            const op = this.consume().value;
            const right = this.parsePower();
            node = { type: 'bin', op, left: node, right };
        }
        return node;
    }

    parsePower() {
        let node = this.parseUnary();
        while (this.peek().type === 'op' && this.peek().value === '^') {
            this.consume();
            const right = this.parseUnary();
            node = { type: 'bin', op: '^', left: node, right };
        }
        return node;
    }

    parseUnary() {
        if (this.peek().type === 'op' && (this.peek().value === '+' || this.peek().value === '-')) {
            const op = this.consume().value;
            const expr = this.parseUnary();
            return { type: 'un', op, expr };
        }
        return this.parsePrimary();
    }

    parsePrimary() {
        const t = this.peek();
        if (t.type === 'number') {
            this.consume();
            return { type: 'num', value: t.value };
        }
        if (t.type === 'string') {
            this.consume();
            return { type: 'str', value: t.value };
        }
        if (t.type === 'ident') {
            const ident = this.consume().value;
            const next = this.peek();
            if (next.type === 'paren' && next.value === '(') {
                this.consume();
                const args = [];
                if (!(this.peek().type === 'paren' && this.peek().value === ')')) {
                    args.push(this.parseExpression());
                    while (this.match('comma')) args.push(this.parseExpression());
                }
                this.match('paren', ')');
                return { type: 'call', name: ident, args };
            }

            const norm = ident.replace(/\$/g, '');
            const refMatch = /([A-Z]+)(\d+)/i.exec(norm.split('!').pop() || '');
            if (refMatch) {
                const ref = refMatch[1].toUpperCase() + refMatch[2];
                if (this.match('colon')) {
                    const t2 = this.consume();
                    const norm2 = String(t2.value || '').replace(/\$/g, '');
                    const refMatch2 = /([A-Z]+)(\d+)/i.exec(norm2.split('!').pop() || '');
                    const ref2 = refMatch2 ? refMatch2[1].toUpperCase() + refMatch2[2] : ref;
                    return { type: 'range', a: ref, b: ref2 };
                }
                return { type: 'ref', ref };
            }

            return { type: 'name', value: ident };
        }
        if (this.match('paren', '(')) {
            const expr = this.parseExpression();
            this.match('paren', ')');
            return expr;
        }
        this.consume();
        return { type: 'num', value: 0 };
    }
}

class ExcelEvaluator {
    static evaluate(node, ctx) {
        switch (node.type) {
            case 'num':
                return node.value;
            case 'str':
                return node.value;
            case 'ref':
                return ctx.getCell(node.ref);
            case 'range':
                return ctx.getRange(node.a, node.b);
            case 'un': {
                const v = ExcelEvaluator.toNumber(ExcelEvaluator.evaluate(node.expr, ctx));
                return node.op === '-' ? -v : v;
            }
            case 'bin': {
                const left = ExcelEvaluator.evaluate(node.left, ctx);
                const right = ExcelEvaluator.evaluate(node.right, ctx);
                return ExcelEvaluator.evalBin(node.op, left, right);
            }
            case 'call':
                return ExcelFunctions.call(node.name, node.args.map((a) => ExcelEvaluator.evaluate(a, ctx)), ctx);
            default:
                return 0;
        }
    }

    static evalBin(op, left, right) {
        if (op === '&') return String(left ?? '') + String(right ?? '');
        if (op === '^') return Math.pow(ExcelEvaluator.toNumber(left), ExcelEvaluator.toNumber(right));
        if (op === '+') return ExcelEvaluator.toNumber(left) + ExcelEvaluator.toNumber(right);
        if (op === '-') return ExcelEvaluator.toNumber(left) - ExcelEvaluator.toNumber(right);
        if (op === '*') return ExcelEvaluator.toNumber(left) * ExcelEvaluator.toNumber(right);
        if (op === '/') {
            const d = ExcelEvaluator.toNumber(right);
            if (d === 0) return 0;
            return ExcelEvaluator.toNumber(left) / d;
        }

        if (['=', '<>', '<', '>', '<=', '>='].includes(op)) {
            const a = ExcelEvaluator.normalizeComparable(left);
            const b = ExcelEvaluator.normalizeComparable(right);
            if (op === '=') return ExcelEvaluator.bool(a === b);
            if (op === '<>') return ExcelEvaluator.bool(a !== b);
            if (typeof a === 'number' && typeof b === 'number') {
                if (op === '<') return ExcelEvaluator.bool(a < b);
                if (op === '>') return ExcelEvaluator.bool(a > b);
                if (op === '<=') return ExcelEvaluator.bool(a <= b);
                if (op === '>=') return ExcelEvaluator.bool(a >= b);
            }
            const sa = String(a);
            const sb = String(b);
            if (op === '<') return ExcelEvaluator.bool(sa < sb);
            if (op === '>') return ExcelEvaluator.bool(sa > sb);
            if (op === '<=') return ExcelEvaluator.bool(sa <= sb);
            if (op === '>=') return ExcelEvaluator.bool(sa >= sb);
        }

        return 0;
    }

    static normalizeComparable(v) {
        if (v == null) return 0;
        if (typeof v === 'string') {
            const trimmed = v.trim();
            if (trimmed === '') return '';
            const n = Number(trimmed);
            return Number.isFinite(n) ? n : trimmed;
        }
        return v;
    }

    static toNumber(v) {
        if (v == null) return 0;
        if (typeof v === 'number') return Number.isFinite(v) ? v : 0;
        if (typeof v === 'boolean') return v ? 1 : 0;
        const s = String(v).trim();
        if (s === '') return 0;
        const n = Number(s);
        return Number.isFinite(n) ? n : 0;
    }

    static bool(b) {
        return b ? 1 : 0;
    }
}

class ExcelFunctions {
    static call(name, args, ctx) {
        const fn = String(name || '').toUpperCase();
        if (fn === 'IF') return ExcelFunctions.IF(args);
        if (fn === 'OR') return ExcelFunctions.OR(args);
        if (fn === 'AND') return ExcelFunctions.AND(args);
        if (fn === 'SUM') return ExcelFunctions.SUM(args);
        if (fn === 'MAX') return ExcelFunctions.MAX(args);
        if (fn === 'ABS') return Math.abs(ExcelEvaluator.toNumber(args[0]));
        if (fn === 'ROUND') return ExcelFunctions.ROUND(args);
        if (fn === 'PI') return Math.PI;
        if (fn === 'VLOOKUP') return ExcelFunctions.VLOOKUP(args);
        if (fn === 'COUNTIF') return ExcelFunctions.COUNTIF(args);
        return 0;
    }

    static IF(args) {
        const cond = ExcelEvaluator.toNumber(args[0]) !== 0;
        return cond ? args[1] : args[2];
    }

    static OR(args) {
        return args.some((v) => ExcelEvaluator.toNumber(v) !== 0) ? 1 : 0;
    }

    static AND(args) {
        return args.every((v) => ExcelEvaluator.toNumber(v) !== 0) ? 1 : 0;
    }

    static SUM(args) {
        const flat = ExcelFunctions.#flattenArgs(args);
        return flat.reduce((sum, v) => sum + ExcelEvaluator.toNumber(v), 0);
    }

    static MAX(args) {
        const flat = ExcelFunctions.#flattenArgs(args).map((v) => ExcelEvaluator.toNumber(v));
        if (!flat.length) return 0;
        return Math.max(...flat);
    }

    static ROUND(args) {
        const n = ExcelEvaluator.toNumber(args[0]);
        const digits = Math.trunc(ExcelEvaluator.toNumber(args[1]));
        const factor = Math.pow(10, Math.abs(digits));
        let rounded;
        if (digits >= 0) {
            rounded = (Math.sign(n) || 1) * Math.round(Math.abs(n) * factor) / factor;
        } else {
            rounded = (Math.sign(n) || 1) * Math.round(Math.abs(n) / factor) * factor;
        }
        return rounded;
    }

    static VLOOKUP(args) {
        const lookupValue = args[0];
        const table = args[1];
        const colIndex = Math.trunc(ExcelEvaluator.toNumber(args[2]));
        const rangeLookup = args[3];
        const exact = rangeLookup === 0 || rangeLookup === false || String(rangeLookup).toUpperCase() === 'FALSE';

        if (!Array.isArray(table) || !table.length || colIndex < 1) return 0;
        if (!exact) return 0;

        for (const row of table) {
            const key = row[0];
            if (ExcelFunctions.#equalsExcel(key, lookupValue)) {
                return row[colIndex - 1] ?? 0;
            }
        }
        return 0;
    }

    static COUNTIF(args) {
        const range = args[0];
        const criteria = args[1];
        const items = ExcelFunctions.#flattenArgs([range]);
        return items.filter((v) => ExcelFunctions.#matchCriteria(v, criteria)).length;
    }

    static #flattenArgs(args) {
        const out = [];
        for (const a of args) {
            if (Array.isArray(a)) {
                for (const row of a) {
                    if (Array.isArray(row)) out.push(...row);
                    else out.push(row);
                }
            } else {
                out.push(a);
            }
        }
        return out;
    }

    static #equalsExcel(a, b) {
        const an = ExcelEvaluator.normalizeComparable(a);
        const bn = ExcelEvaluator.normalizeComparable(b);
        return an === bn;
    }

    static #matchCriteria(value, criteria) {
        const c = criteria == null ? '' : criteria;
        if (typeof c === 'number') return ExcelFunctions.#equalsExcel(value, c);
        const s = String(c);
        const ops = ['>=', '<=', '<>', '>', '<', '='];
        for (const op of ops) {
            if (s.startsWith(op)) {
                const rest = s.slice(op.length);
                const target = rest === '' ? '' : rest;
                const left = ExcelEvaluator.normalizeComparable(value);
                const right = ExcelEvaluator.normalizeComparable(target);
                if (op === '=') return left === right;
                if (op === '<>') return left !== right;
                const ln = ExcelEvaluator.toNumber(left);
                const rn = ExcelEvaluator.toNumber(right);
                if (op === '>') return ln > rn;
                if (op === '<') return ln < rn;
                if (op === '>=') return ln >= rn;
                if (op === '<=') return ln <= rn;
            }
        }
        return ExcelFunctions.#equalsExcel(value, s);
    }
}

export { PanelProperties, ExcelModel };
