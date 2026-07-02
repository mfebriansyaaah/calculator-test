export class CLTLayerType {
    constructor({ index, thicknessMm, orientationDeg, grade }) {
        this.index = Number(index);
        this.thicknessMm = Number(thicknessMm);
        this.orientationDeg = orientationDeg === 90 ? 90 : 0;
        this.grade = grade;
    }

    getE() {
        return this.grade.getEForOrientation(this.orientationDeg);
    }

    getG() {
        return this.orientationDeg === 90 ? this.grade.g90 : this.grade.g;
    }
} 
