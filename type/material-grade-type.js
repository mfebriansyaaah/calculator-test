export class MaterialGrade {
    constructor({ name, e, e90, g, g90 }) {
        this.name = name;
        this.e = Number(e);
        this.e90 = Number(e90);
        this.g = Number(g);
        this.g90 = Number(g90);
    }

    getEForOrientation(orientationDeg) {
        return orientationDeg === 90 ? this.e90 : this.e;
    }
}
