import { EntityMetadataCollection } from "@pptb/types/dataverseAPI";
import { makeAutoObservable } from "mobx";

export type ColumnTestType = "matchesRegex" | "matchesMetadata" | "containsData";

export type ColumnTestDefinition =
    | {
        id: string;
        type: "matchesRegex";
        regexPattern: string;
    }
    | {
        id: string;
        type: "matchesMetadata";
    }
    | {
        id: string;
        type: "containsData";
    };

export class ViewModel {
    metadataLoaded: boolean = false;
    metadata: EntityMetadataCollection | null;
    testsByTableAndColumn: Record<string, Record<string, ColumnTestDefinition[]>>;

    constructor() {
        this.metadata = null;
        this.metadataLoaded = false;
        this.testsByTableAndColumn = {};

        makeAutoObservable(this);
    }

    private createTestId(): string {
        return `${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;
    }

    private ensureColumnTestArray(tableLogicalName: string, columnLogicalName: string): ColumnTestDefinition[] {
        if (!this.testsByTableAndColumn[tableLogicalName]) {
            this.testsByTableAndColumn[tableLogicalName] = {};
        }

        if (!this.testsByTableAndColumn[tableLogicalName][columnLogicalName]) {
            this.testsByTableAndColumn[tableLogicalName][columnLogicalName] = [];
        }

        return this.testsByTableAndColumn[tableLogicalName][columnLogicalName];
    }

    private removeTestType(tableLogicalName: string, columnLogicalName: string, type: ColumnTestType): void {
        const existingTests = this.testsByTableAndColumn[tableLogicalName]?.[columnLogicalName];

        if (!existingTests) {
            return;
        }

        this.testsByTableAndColumn[tableLogicalName][columnLogicalName] = existingTests.filter((test) => test.type !== type);

        if (this.testsByTableAndColumn[tableLogicalName][columnLogicalName].length === 0) {
            delete this.testsByTableAndColumn[tableLogicalName][columnLogicalName];
        }
    }

    setRegexTest(tableLogicalName: string, columnLogicalName: string, regexPattern: string | null): void {
        if (regexPattern === null) {
            this.removeTestType(tableLogicalName, columnLogicalName, "matchesRegex");
            return;
        }

        const existingTests = this.ensureColumnTestArray(tableLogicalName, columnLogicalName);
        const otherTests = existingTests.filter((test) => test.type !== "matchesRegex");

        this.testsByTableAndColumn[tableLogicalName][columnLogicalName] = [...otherTests, {
            id: this.createTestId(),
            type: "matchesRegex",
            regexPattern,
        }];
    }

    setMatchesMetadataTest(tableLogicalName: string, columnLogicalName: string, enabled: boolean): void {
        if (!enabled) {
            this.removeTestType(tableLogicalName, columnLogicalName, "matchesMetadata");
            return;
        }

        const existingTests = this.ensureColumnTestArray(tableLogicalName, columnLogicalName);
        const hasMetadataTest = existingTests.some((test) => test.type === "matchesMetadata");

        if (hasMetadataTest) {
            return;
        }

        this.testsByTableAndColumn[tableLogicalName][columnLogicalName] = [
            ...existingTests,
            {
                id: this.createTestId(),
                type: "matchesMetadata",
            },
        ];
    }

    setContainsDataTest(tableLogicalName: string, columnLogicalName: string, enabled: boolean): void {
        if (!enabled) {
            this.removeTestType(tableLogicalName, columnLogicalName, "containsData");
            return;
        }

        const existingTests = this.ensureColumnTestArray(tableLogicalName, columnLogicalName);
        const alreadyExists = existingTests.some((test) => test.type === "containsData");

        if (alreadyExists) {
            return;
        }

        this.testsByTableAndColumn[tableLogicalName][columnLogicalName] = [
            ...existingTests,
            { id: this.createTestId(), type: "containsData" },
        ];
    }

    getTestsForTable(tableLogicalName: string): Record<string, ColumnTestDefinition[]> {
        return this.testsByTableAndColumn[tableLogicalName] ?? {};
    }

    getTestsForColumn(tableLogicalName: string, columnLogicalName: string): ColumnTestDefinition[] {
        return this.testsByTableAndColumn[tableLogicalName]?.[columnLogicalName] ?? [];
    }
}