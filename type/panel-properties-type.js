export class PanelPropertiesType {
    constructor({ method, beffMm, lengthM, totalThicknessMm, neutralAxisYmm, eiEff } = {}) {
        this.method = method || '';
        this.beffMm = beffMm == null ? 0 : Number(beffMm);
        this.lengthM = lengthM == null ? 0 : Number(lengthM);
        this.totalThicknessMm = totalThicknessMm == null ? 0 : Number(totalThicknessMm);
        this.neutralAxisYmm = neutralAxisYmm == null ? 0 : Number(neutralAxisYmm);
        this.eiEff = eiEff == null ? 0 : Number(eiEff);
        this.excel = { cells: {}, range: null };
        this.layers = [];
    }

    setLayers(layerProperties) {
        this.layers = Array.isArray(layerProperties) ? layerProperties : [];
        return this;
    }

    setExcel({ cells, range }) {
        this.excel = {
            cells: cells || {},
            range: range || null
        };
        return this;
    }
}
