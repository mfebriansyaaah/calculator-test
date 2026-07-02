import { MaterialGrade } from './type/material-grade-type.js';
import { CLTLayerType } from './type/clt-layer-type.js';
import { CLTLayupType } from './type/clt-layup-type.js';
import { PanelProperties } from './calculation/panel-properties.js';

function buildLayerCountOptions(method) {
    if (method === 'gamma') return [3, 5];
    return [3, 4, 5, 6, 7, 8, 9];
}

function getSymmetricOrientationDeg(layerIndex, totalLayers) {
    const i = layerIndex - 1;
    const mirrorIndex = i < totalLayers - 1 - i ? i : totalLayers - 1 - i;
    return mirrorIndex % 2 === 0 ? 0 : 90;
}

function formatNumber(value, decimals = 3) {
    const n = Number(value);
    if (!Number.isFinite(n)) return '-';
    return n.toLocaleString('en-US', { maximumFractionDigits: decimals, minimumFractionDigits: 0 });
}

function renderLayerCountSelect(selectEl, method) {
    const options = buildLayerCountOptions(method);
    const current = Number(selectEl.value);

    selectEl.innerHTML = options.map((v) => `<option value="${v}">${v}</option>`).join('');
    if (options.includes(current)) {
        selectEl.value = String(current);
    } else {
        selectEl.value = String(options[0]);
    }
}

function setOutputVisibility(method) {
    const shear = document.getElementById('output-shear');
    const gamma = document.getElementById('output-gamma');
    if (!shear || !gamma) return;

    if (method === 'gamma') {
        shear.classList.add('d-none');
        gamma.classList.remove('d-none');
    } else {
        gamma.classList.add('d-none');
        shear.classList.remove('d-none');
    }
}

function findGradeListStartRow(excel) {
    for (let r = 1; r <= 60; r++) {
        const v = excel.evaluate(`P${r}`);
        if (typeof v === 'string' && v.trim() === 'Grade List') return r + 1;
    }
    return 8;
}

function getGradeList(excel) {
    const startRow = findGradeListStartRow(excel);
    const out = [];
    for (let r = startRow; r <= startRow + 30; r++) {
        const name = excel.evaluate(`P${r}`);
        const e = excel.evaluate(`Q${r}`);
        if (!name || typeof name !== 'string') break;
        const trimmed = name.trim();
        if (!trimmed) break;
        const eNum = Number(e);
        if (!Number.isFinite(eNum) || eNum === 0) break;
        out.push(trimmed);
    }
    return Array.from(new Set(out));
}

function buildMaterialGrade(excel, gradeName) {
    const startRow = findGradeListStartRow(excel);
    for (let r = startRow; r <= startRow + 30; r++) {
        const name = excel.evaluate(`P${r}`);
        const eProbe = excel.evaluate(`Q${r}`);
        const eNum = Number(eProbe);
        if (!name || typeof name !== 'string') break;
        if (!Number.isFinite(eNum) || eNum === 0) break;
        if (String(name).trim() !== gradeName) continue;
        const e = excel.evaluate(`Q${r}`);
        const e90 = excel.evaluate(`R${r}`);
        const g = excel.evaluate(`S${r}`);
        const g90 = excel.evaluate(`T${r}`);
        return new MaterialGrade({ name: gradeName, e, e90, g, g90 });
    }
    return new MaterialGrade({ name: gradeName, e: 0, e90: 0, g: 0, g90: 0 });
}

function buildLayupFromForm(excel, form) {
    const method = form.method.value;
    const gradeName = form.grade.value;
    const grade = buildMaterialGrade(excel, gradeName);
    const layerCount = Number(form.layerCount.value);
    const thicknessMm = Number(form.thicknessMm.value);
    const beffMm = Number(form.beffMm.value);
    const lengthM = Number(form.lengthM.value);

    const layup = new CLTLayupType({ method, beffMm, lengthM });
    const layers = [];

    for (let i = 1; i <= layerCount; i++) {
        layers.push(
            new CLTLayerType({
                index: i,
                thicknessMm,
                orientationDeg: getSymmetricOrientationDeg(i, layerCount),
                grade
            })
        );
    }

    layup.setLayers(layers);
    return layup;
}

function formatCell(value) {
    if (value == null) return '';
    if (typeof value === 'string') return value;
    const n = Number(value);
    if (!Number.isFinite(n)) return String(value);
    return formatNumber(n, 6);
}

function get(cells, ref) {
    return cells[String(ref).toUpperCase()] ?? '';
}

function formatFixed(value, decimals) {
    const n = Number(value);
    if (!Number.isFinite(n)) return typeof value === 'string' ? value : '';
    return n.toFixed(decimals);
}

function formatSci(value, decimals = 2) {
    const n = Number(value);
    if (!Number.isFinite(n)) return typeof value === 'string' ? value : '';
    if (n === 0) return `${(0).toFixed(decimals)}E+00`;
    const exp = n.toExponential(decimals);
    const [m, eRaw] = exp.split('e');
    const e = Number(eRaw);
    const sign = e >= 0 ? '+' : '-';
    const abs = String(Math.abs(e)).padStart(2, '0');
    return `${m}E${sign}${abs}`;
}

function formatByStyle(value, style) {
    if (value == null) return '';
    if (typeof value === 'string') return value;
    const n = Number(value);
    if (!Number.isFinite(n)) return String(value);

    if (style === 'sci2') return formatSci(n, 2);
    if (style === 'sci3') return formatSci(n, 3);
    if (style === 'fixed0') return Number.isInteger(n) ? String(n) : formatFixed(n, 0);
    if (style === 'fixed1') return formatFixed(n, 1);
    if (style === 'fixed2') return formatFixed(n, 2);
    if (style === 'fixed3') return formatFixed(n, 3);
    return formatCell(n);
}

function isNumericValue(value) {
    if (value == null) return false;
    if (typeof value === 'number') return Number.isFinite(value);
    if (typeof value !== 'string') return false;
    const s = value.trim();
    if (s === '') return false;
    const n = Number(s.replace(/,/g, ''));
    return Number.isFinite(n);
}

function td(value, { key = false, style = null } = {}) {
    const raw = value;
    const text = style ? formatByStyle(raw, style) : formatCell(raw);
    const cls = isNumericValue(raw) ? 'excel-num' : 'excel-text';
    const keyCls = key ? ' excel-key' : '';
    return `<td class="${cls}${keyCls}">${text}</td>`;
}

function renderSectionProperties(cells) {
    const cols = [
        { col: 'Z', header: 'Z38' },
        { col: 'AA', header: 'AA38', unit: 'AA39', style: 'fixed1' },
        { col: 'AB', header: 'AB38', unit: 'AB39', style: 'fixed1' },
        { col: 'AC', header: 'AC38', unit: null, style: 'fixed1' },
        { col: 'AD', header: 'AD38', unit: 'AD39', style: 'fixed1' },
        { col: 'AE', header: 'AE38', unit: 'AE39', style: 'fixed1' },
        { col: 'AF', header: 'AF38', unit: 'AF39', style: 'fixed1' }
    ];

    const rows = [];
    for (let r = 40; r <= 46; r++) {
        rows.push(cols.map((c) => get(cells, `${c.col}${r}`)));
    }

    return `
        <div class="excel-block mb-3">
            <div class="excel-title">${formatCell(get(cells, 'Z36') || 'SECTION PROPERTIES')}</div>
            <div class="p-3">
                <div class="table-responsive">
                    <table class="excel-table excel-green">
                        <thead>
                            <tr>
                                ${cols.map((c) => `<th>${formatCell(get(cells, c.header))}</th>`).join('')}
                            </tr>
                            <tr>
                                ${cols
                                    .map((c) => `<th class="excel-units-row">${c.unit ? formatCell(get(cells, c.unit)) : ''}</th>`)
                                    .join('')}
                            </tr>
                        </thead>
                        <tbody>
                            ${rows
                                .map(
                                    (row) => `
                                <tr>
                                    ${td(row[0])}
                                    ${td(row[1], { style: cols[1].style })}
                                    ${td(row[2], { style: cols[2].style })}
                                    ${td(row[3], { style: cols[3].style })}
                                    ${td(row[4], { style: cols[4].style })}
                                    ${td(row[5], { style: cols[5].style })}
                                    ${td(row[6], { style: cols[6].style })}
                                </tr>`
                                )
                                .join('')}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    `;
}

function renderShearAnalogyBlock(cells) {
    const cols = [
        { col: 'Z', header: 'Z50' },
        { col: 'AA', header: 'AA50', unit: 'AA51', style: 'sci2' },
        { col: 'AC', header: 'AC50', unit: 'AC51', style: 'sci2' },
        { col: 'AE', header: 'AE50', unit: 'AE51', style: 'fixed0' },
        { col: 'AF', header: 'AF50', unit: null, style: 'sci2' }
    ];

    const rows = [];
    for (let r = 52; r <= 58; r++) {
        rows.push(cols.map((c) => get(cells, `${c.col}${r}`)));
    }

    return `
        <div class="excel-block">
            <div class="excel-title">${formatCell(get(cells, 'Z49') || 'Shear Analogy Method')}</div>
            <div class="p-3">
                <div class="table-responsive">
                    <table class="excel-table excel-green">
                        <thead>
                            <tr>
                                ${cols.map((c) => `<th>${formatCell(get(cells, c.header))}</th>`).join('')}
                            </tr>
                            <tr>
                                ${cols
                                    .map((c) => `<th class="excel-units-row">${c.unit ? formatCell(get(cells, c.unit)) : ''}</th>`)
                                    .join('')}
                            </tr>
                        </thead>
                        <tbody>
                            ${rows
                                .map(
                                    (row) => `
                                <tr>
                                    ${td(row[0])}
                                    ${td(row[1], { style: cols[1].style })}
                                    ${td(row[2], { style: cols[2].style })}
                                    ${td(row[3], { style: cols[3].style })}
                                    ${td(row[4], { style: cols[4].style })}
                                </tr>`
                                )
                                .join('')}
                            <tr>
                                <td colspan="3"></td>
                                ${td(get(cells, 'AE59'))}
                                ${td(get(cells, 'AF59'), { style: 'sci2' })}
                            </tr>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    `;
}

function renderGammaBlock(cells) {
    const metaRows = [
        ['Z69', 'AA69', 'AB69', 'AD69', 'AE69'],
        ['Z70', 'AA70', 'AB70', '', '']
    ].map((refs) => refs.map((r) => (r ? get(cells, r) : '')));

    const cols = [
        { col: 'Z', header: 'Z72' },
        { col: 'AA', header: 'AA72', unit: 'AA73', style: 'fixed1' },
        { col: 'AB', header: 'AB72', unit: null },
        { col: 'AC', header: 'AC72', unit: 'AC73', style: 'fixed1' },
        { col: 'AD', header: 'AD72', unit: null, style: 'fixed2' },
        { col: 'AE', header: 'AE72', unit: 'AE73', style: 'fixed0' },
        { col: 'AF', header: 'AF72', unit: 'AF73', style: 'fixed1' },
        { col: 'AG', header: 'AG72', unit: null },
        { col: 'AH', header: 'AH72', unit: null, style: 'fixed3' },
        { col: 'AI', header: 'AI72', unit: null },
        { col: 'AJ', header: 'AJ72', unit: 'AJ73', style: 'fixed1' },
        { col: 'AK', header: 'AK72', unit: 'AK73', style: 'sci2' },
        { col: 'AL', header: 'AL72', unit: 'AL73', style: 'sci2' },
        { col: 'AM', header: 'AM72', unit: null, style: 'sci2' }
    ];

    const rows = [];
    for (let r = 74; r <= 80; r++) {
        rows.push(cols.map((c) => get(cells, `${c.col}${r}`)));
    }

    return `
        <div class="excel-block">
            <div class="excel-title">${formatCell(get(cells, 'Z67') || 'Gamma Method')}</div>
            <div class="p-3">
                <div class="table-responsive mb-3">
                    <table class="excel-table">
                        <tbody>
                            ${metaRows
                                .map(
                                    (row) => `
                                <tr>
                                    ${td(row[0])}
                                    ${td(row[1], { style: 'fixed0' })}
                                    ${td(row[2])}
                                    ${td(row[3])}
                                    ${td(row[4], { style: 'fixed0' })}
                                </tr>`
                                )
                                .join('')}
                        </tbody>
                    </table>
                </div>

                <div class="table-responsive">
                    <table class="excel-table excel-green">
                        <thead>
                            <tr>
                                ${cols.map((c) => `<th>${formatCell(get(cells, c.header))}</th>`).join('')}
                            </tr>
                            <tr>
                                ${cols
                                    .map((c) => `<th class="excel-units-row">${c.unit ? formatCell(get(cells, c.unit)) : ''}</th>`)
                                    .join('')}
                            </tr>
                        </thead>
                        <tbody>
                            ${rows
                                .map(
                                    (row) => `
                                <tr>
                                    ${row
                                        .map((v, idx) => td(v, { style: cols[idx].style || null }))
                                        .join('')}
                                </tr>`
                                )
                                .join('')}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    `;
}

function renderShearOutput(result) {
    const cells = result.excel.cells || {};
    return `
        ${renderSectionProperties(cells)}
        ${renderShearAnalogyBlock(cells)}
    `;
}

function renderGammaOutput(result) {
    const cells = result.excel.cells || {};
    return `
        ${renderSectionProperties(cells)}
        ${renderGammaBlock(cells)}
    `;
}

async function initCalculator() {
    const root = document.getElementById('calculator-root');
    if (!root) return;

    root.innerHTML = `
        <div id="loading" class="alert alert-info">Memuat model Excel…</div>
        <form id="calculator-form" class="card mb-3 d-none">
            <div class="card-body">
                <div class="row g-3">
                    <div class="col-12 col-md-4">
                        <label class="form-label">Analytical Method</label>
                        <select class="form-select" name="method">
                            <option value="shear-analogy">Shear Analogy</option>
                            <option value="gamma">Gamma</option>
                        </select>
                    </div>
                    <div class="col-12 col-md-4">
                        <label class="form-label">Grade</label>
                        <select class="form-select" name="grade"></select>
                    </div>
                    <div class="col-12 col-md-4">
                        <label class="form-label">Total Layers</label>
                        <select class="form-select" name="layerCount"></select>
                    </div>
                    <div class="col-12 col-md-4">
                        <label class="form-label">Thickness Each Layer (mm)</label>
                        <input class="form-control" name="thicknessMm" type="number" step="0.001" min="1" />
                    </div>
                    <div class="col-12 col-md-4">
                        <label class="form-label">b<sub>eff</sub> (mm)</label>
                        <input class="form-control" name="beffMm" type="number" step="1" min="1" />
                    </div>
                    <div class="col-12 col-md-4">
                        <label class="form-label">Length (m)</label>
                        <input class="form-control" name="lengthM" type="number" step="0.001" min="0.1" />
                    </div>
                    <div class="col-12 d-flex gap-2">
                        <button class="btn btn-primary" type="submit">Calculate</button>
                        <div id="calculator-error" class="alert alert-danger py-2 px-3 mb-0 d-none flex-grow-1"></div>
                    </div>
                </div>
            </div>
        </form>

        <div class="excel-like">
            <div id="output-shear"></div>
            <div id="output-gamma" class="d-none"></div>
        </div>
    `;

    const loadingEl = document.getElementById('loading');
    const form = document.getElementById('calculator-form');
    const methodEl = form.elements.method;
    const layerCountEl = form.elements.layerCount;
    const gradeEl = form.elements.grade;

    let xlsxBuffer;
    try {
        xlsxBuffer = await fetch('/floor-panel-properties.xlsx').then((r) => r.arrayBuffer());
    } catch (err) {
        if (loadingEl) {
            loadingEl.classList.remove('alert-info');
            loadingEl.classList.add('alert-warning');
            loadingEl.textContent = 'Tidak bisa memuat floor-panel-properties.xlsx otomatis. Pilih file .xlsx secara manual.';
        }

        xlsxBuffer = await new Promise((resolve, reject) => {
            const input = document.createElement('input');
            input.type = 'file';
            input.accept = '.xlsx';
            input.className = 'form-control mt-2';
            input.addEventListener('change', () => {
                const file = input.files && input.files[0];
                if (!file) return;
                const reader = new FileReader();
                reader.onload = () => resolve(reader.result);
                reader.onerror = () => reject(reader.error || new Error('Gagal membaca file.'));
                reader.readAsArrayBuffer(file);
            });
            if (loadingEl) loadingEl.appendChild(input);
        });
    }

    const excel = await PanelProperties.loadFromXlsxBuffer(xlsxBuffer);

    const grades = getGradeList(excel);
    gradeEl.innerHTML = grades.map((g) => `<option value="${g}">${g}</option>`).join('');

    const defaultGrade = String(excel.evaluate('B9') || grades[0] || '');
    if (defaultGrade) gradeEl.value = defaultGrade;

    const defaultThickness = excel.evaluate('F9');
    const defaultLayers = excel.evaluate('D9');
    const defaultLength = excel.evaluate('H9');
    const defaultBeff = excel.evaluate('D18');

    form.elements.thicknessMm.value = defaultThickness || 35;
    form.elements.layerCount.value = defaultLayers || 5;
    form.elements.lengthM.value = defaultLength || 5;
    form.elements.beffMm.value = defaultBeff || 1000;

    renderLayerCountSelect(layerCountEl, methodEl.value);
    setOutputVisibility(methodEl.value);

    if (loadingEl) loadingEl.classList.add('d-none');
    form.classList.remove('d-none');

    methodEl.addEventListener('change', () => {
        renderLayerCountSelect(layerCountEl, methodEl.value);
        setOutputVisibility(methodEl.value);
    });

    form.addEventListener('submit', (e) => {
        e.preventDefault();

        const errorEl = document.getElementById('calculator-error');
        const outputShear = document.getElementById('output-shear');
        const outputGamma = document.getElementById('output-gamma');

        try {
            const layup = buildLayupFromForm(excel, form.elements);
            const result = PanelProperties.calculate(layup);

            if (errorEl) errorEl.classList.add('d-none');
            if (result.method === 'gamma') {
                if (outputGamma) outputGamma.innerHTML = renderGammaOutput(result);
            } else {
                if (outputShear) outputShear.innerHTML = renderShearOutput(result);
            }
        } catch (err) {
            if (errorEl) {
                errorEl.textContent = err instanceof Error ? err.message : String(err);
                errorEl.classList.remove('d-none');
            }
        }
    });
}

document.addEventListener('DOMContentLoaded', initCalculator);
