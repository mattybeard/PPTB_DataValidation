import { observer } from "mobx-react";
import { useContext, useEffect, useMemo, useState } from "react";
import { ServiceProviderContext } from "../framework/ServiceProvider";
import { ViewModel } from "../models/ViewModel";
import {
  Checkbox,
  DefaultButton,
  DetailsList,
  DetailsRow,
  Dialog,
  DialogFooter,
  DialogType,
  Dropdown,
  DropdownMenuItemType,
  IColumn,
  IDropdownOption,
  IDropdownStyles,
  IStackTokens,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  Selection,
  SelectionMode,
  Spinner,
  Stack,
  TextField,
  Text,
} from "@fluentui/react";

interface TableSelectionProps {
  connection: ToolBoxAPI.DataverseConnection | null;
}

interface AttributeRow {
  key: string;
  logicalName: string;
  displayName: string;
  fieldType: string;
  testCount: number;
}

type SortableColumnKey = "logicalName" | "displayName" | "fieldType" | "testCount";
type TestDialogMode = "single" | "bulk";

export const TableSelection = observer(
  (props: TableSelectionProps): React.ReactElement => {
    const [selectedTable, setSelectedTable] = useState<string | null>(null);
    const [attributeRows, setAttributeRows] = useState<AttributeRow[]>([]);
    const [attributesLoading, setAttributesLoading] = useState(false);
    const [attributesError, setAttributesError] = useState<string | null>(null);
    const [isAddTestDialogOpen, setIsAddTestDialogOpen] = useState(false);
    const [testDialogMode, setTestDialogMode] = useState<TestDialogMode>("single");
    const [activeColumnLogicalName, setActiveColumnLogicalName] = useState<string | null>(null);
    const [selectedColumnLogicalNames, setSelectedColumnLogicalNames] = useState<string[]>([]);
    const [sortColumnKey, setSortColumnKey] = useState<SortableColumnKey>("logicalName");
    const [isSortDescending, setIsSortDescending] = useState(false);
    const [isMatchesRegexEnabled, setIsMatchesRegexEnabled] = useState(false);
    const [isMatchesMetadataEnabled, setIsMatchesMetadataEnabled] = useState(false);
    const [isContainsDataEnabled, setIsContainsDataEnabled] = useState(false);
    const [regexPattern, setRegexPattern] = useState("");
    const [addTestDialogError, setAddTestDialogError] = useState<string | null>(null);
    const [runTestsMessage, setRunTestsMessage] = useState<string | null>(null);

    const context = useContext(ServiceProviderContext);
    const vm = context.get<ViewModel>("vm");

    const stackTokens: IStackTokens = { childrenGap: 20 };
    const dropdownStyles: Partial<IDropdownStyles> = {
      dropdown: { width: 360 },
    };

    const selection = useMemo(() => {
      return new Selection({
        getKey: (item) => (item as AttributeRow).logicalName,
        onSelectionChanged: () => {
          const selectedItems = selection.getSelection() as AttributeRow[];
          setSelectedColumnLogicalNames(selectedItems.map((item) => item.logicalName));
        },
      });
    }, []);

    const tableOptions = useMemo((): IDropdownOption[] => {
      if (!vm.metadata) return [];

      const options: IDropdownOption[] = [];
      options.push({ key: 'headerRow', text: 'Tables', itemType: DropdownMenuItemType.Header });

      vm.metadata.value.forEach((entity) => {
        const displayLabel = entity.DisplayName?.LocalizedLabels?.[0]?.Label;
        const label = displayLabel ? `${displayLabel} (${entity.LogicalName})` : entity.LogicalName;
        options.push({ key: entity.LogicalName, text: label });
      });

      return options;
    }, [vm.metadata]);

    const selectedTestsByColumn: Record<string, number> = {};

    if (selectedTable) {
      const testsForTable = vm.getTestsForTable(selectedTable);

      Object.keys(testsForTable).forEach((columnLogicalName) => {
        selectedTestsByColumn[columnLogicalName] = testsForTable[columnLogicalName].length;
      });
    }

    const totalSelectedTests = Object.values(selectedTestsByColumn).reduce((sum, count) => sum + count, 0);

    const handleColumnHeaderClick = (_event?: React.MouseEvent<HTMLElement>, column?: IColumn): void => {
      const clickedColumnKey = column?.key as SortableColumnKey | undefined;

      if (!clickedColumnKey) {
        return;
      }

      if (sortColumnKey === clickedColumnKey) {
        setIsSortDescending((previous) => !previous);
      } else {
        setSortColumnKey(clickedColumnKey);
        setIsSortDescending(false);
      }
    };

    const sortedAttributeRows = useMemo((): AttributeRow[] => {
      const items = attributeRows.map((row) => ({
        ...row,
        testCount: selectedTestsByColumn[row.logicalName] ?? 0,
      }));

      items.sort((left, right) => {
        if (sortColumnKey === "testCount") {
          const leftValue = left.testCount;
          const rightValue = right.testCount;

          if (leftValue < rightValue) return isSortDescending ? 1 : -1;
          if (leftValue > rightValue) return isSortDescending ? -1 : 1;
          return 0;
        }

        const leftValue = left[sortColumnKey].toLowerCase();
        const rightValue = right[sortColumnKey].toLowerCase();

        if (leftValue < rightValue) return isSortDescending ? 1 : -1;
        if (leftValue > rightValue) return isSortDescending ? -1 : 1;
        return 0;
      });

      return items;
    }, [attributeRows, selectedTestsByColumn, sortColumnKey, isSortDescending]);

    const detailsListColumns = useMemo((): IColumn[] => {
      return [
        {
          key: "logicalName",
          name: "Logical Name",
          fieldName: "logicalName",
          minWidth: 170,
          maxWidth: 240,
          isResizable: true,
          isSorted: sortColumnKey === "logicalName",
          isSortedDescending: isSortDescending,
          onColumnClick: handleColumnHeaderClick,
        },
        {
          key: "displayName",
          name: "Display Name",
          fieldName: "displayName",
          minWidth: 220,
          maxWidth: 320,
          isResizable: true,
          isSorted: sortColumnKey === "displayName",
          isSortedDescending: isSortDescending,
          onColumnClick: handleColumnHeaderClick,
        },
        {
          key: "fieldType",
          name: "Field Type",
          fieldName: "fieldType",
          minWidth: 120,
          maxWidth: 180,
          isResizable: true,
          isSorted: sortColumnKey === "fieldType",
          isSortedDescending: isSortDescending,
          onColumnClick: handleColumnHeaderClick,
          onRender: (item: AttributeRow) => item.fieldType.replace(/Type$/i, ""),
        },
        {
          key: "testCount",
          name: "Test Count",
          fieldName: "testCount",
          minWidth: 100,
          maxWidth: 120,
          isResizable: true,
          isSorted: sortColumnKey === "testCount",
          isSortedDescending: isSortDescending,
          onColumnClick: handleColumnHeaderClick,
        },
        {
          key: "actions",
          name: "Actions",
          minWidth: 160,
          maxWidth: 180,
          onRender: (item: AttributeRow) => {
            return (
              <DefaultButton
                text="Edit Test(s)"
                onClick={() => {
                  if (!selectedTable) {
                    return;
                  }

                  const existingRegexTest = vm
                    .getTestsForColumn(selectedTable, item.logicalName)
                    .find((test) => test.type === "matchesRegex");
                  const existingMatchesMetadataTest = vm
                    .getTestsForColumn(selectedTable, item.logicalName)
                    .find((test) => test.type === "matchesMetadata");
                  const existingContainsDataTest = vm
                    .getTestsForColumn(selectedTable, item.logicalName)
                    .find((test) => test.type === "containsData");

                  setTestDialogMode("single");
                  setActiveColumnLogicalName(item.logicalName);
                  setIsMatchesRegexEnabled(Boolean(existingRegexTest));
                  setIsMatchesMetadataEnabled(Boolean(existingMatchesMetadataTest));
                  setIsContainsDataEnabled(Boolean(existingContainsDataTest));
                  setRegexPattern(existingRegexTest?.regexPattern ?? "");
                  setAddTestDialogError(null);
                  setIsAddTestDialogOpen(true);
                }}
              />
            );
          },
        },
      ];
    }, [selectedTable, sortColumnKey, isSortDescending, vm, selectedTestsByColumn]);

    const getDisplayLabel = (displayName: any): string => {
      return (
        displayName?.UserLocalizedLabel?.Label ??
        displayName?.LocalizedLabels?.[0]?.Label ??
        "N/A"
      );
    };

    const getFieldType = (attribute: any): string => {
      return attribute?.AttributeTypeName?.Value ?? attribute?.AttributeType ?? "Unknown";
    };

    useEffect(() => {
      selection.setItems(sortedAttributeRows, false);
    }, [selection, sortedAttributeRows]);

    useEffect(() => {
      if (!props.connection) {
        setSelectedTable(null);
        setAttributeRows([]);
        setAttributesError(null);
        setIsAddTestDialogOpen(false);
        setTestDialogMode("single");
        setActiveColumnLogicalName(null);
        setSelectedColumnLogicalNames([]);
        selection.setAllSelected(false);
        setIsMatchesRegexEnabled(false);
        setIsMatchesMetadataEnabled(false);
        setIsContainsDataEnabled(false);
        setRegexPattern("");
        setAddTestDialogError(null);
        setRunTestsMessage(null);
        return;
      }

      if (!vm.metadataLoaded) {
        window.dataverseAPI.getAllEntitiesMetadata().then((entityMetadata) => {
          vm.metadata = entityMetadata;
          vm.metadataLoaded = true;
        });
      }
    }, [props.connection, selection, vm]);

    useEffect(() => {
      if (!selectedTable) {
        setAttributeRows([]);
        setAttributesError(null);
        setAttributesLoading(false);
        return;
      }

      setSelectedColumnLogicalNames([]);
      selection.setAllSelected(false);

      setAttributesLoading(true);
      setAttributesError(null);

      window.dataverseAPI
        .getEntityRelatedMetadata(selectedTable, "Attributes")
        .then((result: any) => {
          const excludedFieldTypes = new Set(["VirtualType", "UniqueIdentifierType", "UniqueidentifierType"]);

          const rows: AttributeRow[] = (result?.value ?? [])
            .map((attribute: any) => {
              const displayName = getDisplayLabel(attribute.DisplayName);
              const fieldType = getFieldType(attribute);

              return {
                key: attribute.LogicalName,
                logicalName: attribute.LogicalName,
                displayName,
                fieldType,
                testCount: 0,
              };
            })
            .filter((row: AttributeRow) => row.displayName !== "N/A")
            .filter((row: AttributeRow) => !excludedFieldTypes.has(row.fieldType));

          setAttributeRows(rows);
        })
        .catch((error: Error) => {
          setAttributesError(error.message);
          setAttributeRows([]);
        })
        .finally(() => {
          setAttributesLoading(false);
        });
    }, [selectedTable, selection]);

    const handleAddTestDialogDismiss = () => {
      setIsAddTestDialogOpen(false);
      setTestDialogMode("single");
      setActiveColumnLogicalName(null);
      setIsMatchesRegexEnabled(false);
      setIsMatchesMetadataEnabled(false);
      setIsContainsDataEnabled(false);
      setRegexPattern("");
      setAddTestDialogError(null);
    };

    const handleSaveTest = () => {
      if (!selectedTable) {
        handleAddTestDialogDismiss();
        return;
      }

      const targetColumns = testDialogMode === "bulk"
        ? selectedColumnLogicalNames
        : activeColumnLogicalName
          ? [activeColumnLogicalName]
          : [];

      if (targetColumns.length === 0) {
        handleAddTestDialogDismiss();
        return;
      }

      if (!isMatchesRegexEnabled && !isMatchesMetadataEnabled && !isContainsDataEnabled) {
        targetColumns.forEach((columnLogicalName) => {
          vm.setRegexTest(selectedTable, columnLogicalName, null);
          vm.setMatchesMetadataTest(selectedTable, columnLogicalName, false);
          vm.setContainsDataTest(selectedTable, columnLogicalName, false);
        });
        setRunTestsMessage(null);
        handleAddTestDialogDismiss();
        return;
      }

      targetColumns.forEach((columnLogicalName) => {
        vm.setMatchesMetadataTest(selectedTable, columnLogicalName, isMatchesMetadataEnabled);
      });

      targetColumns.forEach((columnLogicalName) => {
        vm.setContainsDataTest(selectedTable, columnLogicalName, isContainsDataEnabled);
      });

      if (!isMatchesRegexEnabled) {
        targetColumns.forEach((columnLogicalName) => {
          vm.setRegexTest(selectedTable, columnLogicalName, null);
        });
        setRunTestsMessage(null);
        handleAddTestDialogDismiss();
        return;
      }

      const trimmedRegexPattern = regexPattern.trim();

      if (!trimmedRegexPattern) {
        setAddTestDialogError("Regex pattern is required when Matches Regex is selected.");
        return;
      }

      try {
        new RegExp(trimmedRegexPattern);
      } catch {
        setAddTestDialogError("Regex pattern is not valid. Please correct it and try again.");
        return;
      }

      targetColumns.forEach((columnLogicalName) => {
        vm.setRegexTest(selectedTable, columnLogicalName, trimmedRegexPattern);
      });
      setRunTestsMessage(null);
      handleAddTestDialogDismiss();
    };

    const handleBulkAddTests = () => {
      if (!selectedTable || selectedColumnLogicalNames.length === 0) {
        return;
      }

      setTestDialogMode("bulk");
      setActiveColumnLogicalName(null);
      setIsMatchesRegexEnabled(false);
      setIsMatchesMetadataEnabled(false);
      setIsContainsDataEnabled(false);
      setRegexPattern("");
      setAddTestDialogError(null);
      setIsAddTestDialogOpen(true);
    };

    const handleRunSelectedTests = () => {
      setRunTestsMessage(`Prepared to run ${totalSelectedTests} selected test(s). Execution wiring will be added next.`);
    };

    if (!props.connection) {
      return (
        <div className="card">
          <MessageBar messageBarType={MessageBarType.warning}>
            Connect to a Dataverse environment to select a table and configure column tests.
          </MessageBar>
        </div>
      );
    }

    const showColumnsSection = Boolean(selectedTable);

    return (
      <>
        <div className="card">
          <Stack tokens={stackTokens}>
            <Text variant="xLarge">Dataverse Table Test Setup</Text>

            {!vm.metadataLoaded && (
              <div>
                <Spinner label="Loading table metadata" />
              </div>
            )}

            {vm.metadataLoaded && (
              <>
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }}>
                  <Text variant="medium" style={{ whiteSpace: "nowrap" }}>Select a table:</Text>
                  <Dropdown
                    placeholder="Choose a table..."
                    options={tableOptions}
                    styles={dropdownStyles}
                    selectedKey={selectedTable ?? undefined}
                    onChange={(_, option) => {
                      setSelectedTable(option?.key as string | null);
                      setRunTestsMessage(null);
                    }}
                  />
                </Stack>

                {showColumnsSection && (
                  <>
                    <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                      <Text variant="large">Columns for {selectedTable}</Text>
                      <DefaultButton
                        text={`Add Test To All Selected (${selectedColumnLogicalNames.length})`}
                        disabled={selectedColumnLogicalNames.length === 0}
                        onClick={handleBulkAddTests}
                      />
                    </Stack>

                    {attributesLoading && <Spinner label="Loading columns" />}

                    {attributesError && (
                      <MessageBar messageBarType={MessageBarType.error}>
                        Failed to load columns: {attributesError}
                      </MessageBar>
                    )}

                    {!attributesLoading && !attributesError && sortedAttributeRows.length === 0 && (
                      <MessageBar messageBarType={MessageBarType.warning}>
                        No columns were returned for this table.
                      </MessageBar>
                    )}

                    {!attributesLoading && !attributesError && sortedAttributeRows.length > 0 && (
                      <DetailsList
                        items={sortedAttributeRows}
                        columns={detailsListColumns}
                        setKey="tableColumns"
                        selection={selection}
                        selectionMode={SelectionMode.multiple}
                        checkboxVisibility={2}
                        compact
                        onRenderRow={(rowProps) => {
                          if (!rowProps) return null;
                          return (
                            <DetailsRow
                              {...rowProps}
                              styles={{ cell: { display: "flex", alignItems: "center" } }}
                            />
                          );
                        }}
                      />
                    )}

                    <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                      <Text variant="medium">
                        Selected rows: {selectedColumnLogicalNames.length} | Selected tests: {totalSelectedTests}
                      </Text>

                      <PrimaryButton
                        text="Run Selected Tests"
                        disabled={totalSelectedTests === 0}
                        onClick={handleRunSelectedTests}
                      />
                    </Stack>

                    {runTestsMessage && (
                      <MessageBar messageBarType={MessageBarType.info}>{runTestsMessage}</MessageBar>
                    )}
                  </>
                )}
              </>
            )}
          </Stack>
        </div>

        <Dialog
          hidden={!isAddTestDialogOpen}
          onDismiss={handleAddTestDialogDismiss}
          dialogContentProps={{
            type: DialogType.normal,
            title: "Edit Test(s)",
            subText: testDialogMode === "bulk"
              ? `Apply the same tests to ${selectedColumnLogicalNames.length} selected column(s).`
              : activeColumnLogicalName
                ? `Configure tests for ${activeColumnLogicalName}.`
                : "Create a test for this column.",
          }}
        >
          <Stack tokens={{ childrenGap: 12 }}>
            {addTestDialogError && (
              <MessageBar messageBarType={MessageBarType.error}>
                {addTestDialogError}
              </MessageBar>
            )}

            <Checkbox
              label="Matches Regex"
              checked={isMatchesRegexEnabled}
              onChange={(_, checked) => {
                const isChecked = Boolean(checked);
                setIsMatchesRegexEnabled(isChecked);
                if (!isChecked) {
                  setRegexPattern("");
                }
                setAddTestDialogError(null);
              }}
            />

            {isMatchesRegexEnabled && (
              <TextField
                label="Regex Pattern"
                placeholder="e.g. ^[A-Z]{3}\\d{4}$"
                value={regexPattern}
                styles={{ root: { paddingLeft: 28 } }}
                onChange={(_, value) => {
                  setRegexPattern(value ?? "");
                  setAddTestDialogError(null);
                }}
              />
            )}

            <Checkbox
              label="Matches Metadata"
              checked={isMatchesMetadataEnabled}
              onChange={(_, checked) => {
                setIsMatchesMetadataEnabled(Boolean(checked));
                setAddTestDialogError(null);
              }}
            />

            <Checkbox
              label="Contains Data"
              checked={isContainsDataEnabled}
              onChange={(_, checked) => {
                setIsContainsDataEnabled(Boolean(checked));
                setAddTestDialogError(null);
              }}
            />
          </Stack>

          <DialogFooter>
            <PrimaryButton text="Save Test" onClick={handleSaveTest} />
            <DefaultButton text="Cancel" onClick={handleAddTestDialogDismiss} />
          </DialogFooter>
        </Dialog>
      </>
    );
});