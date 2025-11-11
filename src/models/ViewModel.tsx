import { EntityMetadataCollection } from "@pptb/types/dataverseAPI";
import { makeAutoObservable } from "mobx";

export class ViewModel {
    metadataLoaded: boolean = false;
    metadata: EntityMetadataCollection | null;

    constructor() {
        this.metadata = null;
        this.metadataLoaded = false;

        makeAutoObservable(this);
    }
}