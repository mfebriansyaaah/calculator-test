export class CLTLayupType {
    constructor({ method, beffMm, lengthM } = {}) {
        this.name = 'CLT Layup';
        this.method = method || 'shear-analogy';
        this.beffMm = beffMm == null ? 1000 : Number(beffMm);
        this.lengthM = lengthM == null ? 5 : Number(lengthM);
        /**
         * @type {CLTLayerType[]}
         */
        this.layers = [];
    }

    addLayer(layer) {
        this.layers.push(layer);
        return this;
    }

    setLayers(layers) {
        this.layers = Array.isArray(layers) ? layers : [];
        return this;
    }

    getLayers() {
        return this.layers.slice();
    }

    getLayerCount() {
        return this.layers.length;
    }

    getTotalThicknessMm() {
        return this.layers.reduce((sum, layer) => sum + Number(layer.thicknessMm || 0), 0);
    }

    isSymmetric() {
        const n = this.layers.length;
        for (let i = 0; i < Math.floor(n / 2); i++) {
            const a = this.layers[i];
            const b = this.layers[n - 1 - i];
            if (Number(a.thicknessMm) !== Number(b.thicknessMm)) return false;
            if (Number(a.orientationDeg) !== Number(b.orientationDeg)) return false;
        }
        return true;
    }

    validate() {
        const n = this.layers.length;
        if (this.method === 'gamma') {
            if (!(n === 3 || n === 5)) {
                return { ok: false, message: 'Gamma hanya bisa 3 atau 5 layer.' };
            }
        } else {
            if (n < 3 || n > 9) {
                return { ok: false, message: 'Shear Analogy hanya bisa 3 sampai 9 layer.' };
            }
            if (!this.isSymmetric()) {
                return { ok: false, message: 'Shear Analogy harus simetris dari atas ke bawah.' };
            }
        }

        if (!(this.beffMm > 0)) return { ok: false, message: 'beff harus > 0.' };
        if (!(this.lengthM > 0)) return { ok: false, message: 'Length harus > 0.' };
        if (this.layers.some((l) => !(Number(l.thicknessMm) > 0))) {
            return { ok: false, message: 'Thickness tiap layer harus > 0.' };
        }

        return { ok: true, message: '' };
    }
}
