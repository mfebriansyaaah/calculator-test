export class CLTLayerPropertiesType {
    constructor({
        index,
        orientationDeg,
        thicknessMm,
        e,
        g,
        areaMm2,
        centroidYmm,
        iCentroidMm4,
        distanceToNaMm,
        gamma,
        eiContribution
    }) {
        this.index = index;
        this.orientationDeg = orientationDeg;
        this.thicknessMm = thicknessMm;
        this.e = e;
        this.g = g;
        this.areaMm2 = areaMm2;
        this.centroidYmm = centroidYmm;
        this.iCentroidMm4 = iCentroidMm4;
        this.distanceToNaMm = distanceToNaMm;
        this.gamma = gamma;
        this.eiContribution = eiContribution;
    }
}
