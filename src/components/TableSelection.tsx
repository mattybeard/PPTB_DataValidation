import { observer } from "mobx-react";
import { useContext, useEffect, useMemo, useState } from "react";
import { ServiceProviderContext } from "../framework/ServiceProvider";
import { ViewModel } from "../models/ViewModel";
import {
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
  ProgressIndicator,
  Selection,
  SelectionMode,
  Spinner,
  Stack,
  Text,
  TextField,
} from "@fluentui/react";
import { ColumnTestResult, TestResultsPanel, TestRunResults } from "./TestResultsPanel";
import { EditTestsDialog } from "./EditTestsDialog";

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
    const [attributeMetadataByLogicalName, setAttributeMetadataByLogicalName] =
      useState<Record<string, any>>({});
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
    const [isRunningTests, setIsRunningTests] = useState(false);
    const [testResults, setTestResults] = useState<TestRunResults | null>(null);
    const [progressPercentage, setProgressPercentage] = useState(0);
    const [progressDescription, setProgressDescription] = useState("");
    const [columnFilter, setColumnFilter] = useState("");
    const [isRunDialogOpen, setIsRunDialogOpen] = useState(false);
    const [isResultsFullscreen, setIsResultsFullscreen] = useState(false);

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

      // Filter by logical name or display name
      const filteredItems = columnFilter.trim() === ""
        ? items
        : items.filter((row) => {
            const filterLower = columnFilter.toLowerCase();
            return (
              row.logicalName.toLowerCase().includes(filterLower) ||
              row.displayName.toLowerCase().includes(filterLower)
            );
          });

      filteredItems.sort((left, right) => {
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

      return filteredItems;
    }, [attributeRows, selectedTestsByColumn, sortColumnKey, isSortDescending, columnFilter]);

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
        setAttributeMetadataByLogicalName({});
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
        setTestResults(null);
        setIsRunningTests(false);
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
        setAttributeMetadataByLogicalName({});
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
          const metadataMap: Record<string, any> = {};

          const rows: AttributeRow[] = (result?.value ?? [])
            .map((attribute: any) => {
              const displayName = getDisplayLabel(attribute.DisplayName);
              const fieldType = getFieldType(attribute);

              metadataMap[attribute.LogicalName] = attribute;

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
          setAttributeMetadataByLogicalName(metadataMap);
        })
        .catch((error: Error) => {
          setAttributesError(error.message);
          setAttributeRows([]);
          setAttributeMetadataByLogicalName({});
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
        setTestResults(null);
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
        setTestResults(null);
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
      setTestResults(null);
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

    const handleToggleSelectAll = () => {
      if (sortedAttributeRows.length === 0) {
        return;
      }

      const allRowsSelected = selectedColumnLogicalNames.length === sortedAttributeRows.length;
      selection.setAllSelected(!allRowsSelected);
    };

    const handleRunSelectedTests = async () => {
      if (!selectedTable) return;

      const testsForTable = vm.getTestsForTable(selectedTable);
      const columnLogicalNames = Object.keys(testsForTable);
      if (columnLogicalNames.length === 0) return;

      const totalConfiguredTests = columnLogicalNames.reduce(
        (sum, logicalName) => sum + (testsForTable[logicalName]?.length ?? 0),
        0
      );

      setIsRunningTests(true);
      setTestResults(null);
      setRunTestsMessage(null);
      setProgressPercentage(0);
      setProgressDescription("Fetching records...");
      setIsRunDialogOpen(true);
      setIsResultsFullscreen(false);

      try {
        let estimatedTotalRecords = 0;
        const selectedTableMetadata = vm.metadata?.value?.find(
          (entity: any) => entity?.LogicalName === selectedTable
        );
        const primaryIdAttribute = selectedTableMetadata?.PrimaryIdAttribute || `${selectedTable}id`;

        const countFetchXml = `<fetch aggregate="true">\n  <entity name="${selectedTable}">\n    <attribute name="${primaryIdAttribute}" alias="recordcount" aggregate="count" />\n  </entity>\n</fetch>`;

        console.log("[fetchXml] count", countFetchXml);

        try {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const countResult = (await window.dataverseAPI.fetchXmlQuery(countFetchXml)) as any;
          const countRow = Array.isArray(countResult?.value) ? countResult.value[0] : null;
          const parsedCount = Number(countRow?.recordcount ?? countRow?.RecordCount ?? 0);
          if (Number.isFinite(parsedCount) && parsedCount > 0) {
            estimatedTotalRecords = parsedCount;
            setProgressDescription(
              `Found ${estimatedTotalRecords.toLocaleString()} records. Fetching page 1...`
            );
          }
        } catch {
          setProgressDescription("Fetching records...");
        }

        const attributeXml = columnLogicalNames
          .map((name) => `    <attribute name="${name}" />`)
          .join("\n");

        const allRecords: Record<string, unknown>[] = [];
        let page = 1;
        let moreRecords = true;

        while (moreRecords) {
          const fetchXml = `<fetch count="2000" page="${page}">\n  <entity name="${selectedTable}">\n${attributeXml}\n  </entity>\n</fetch>`;

          console.log(`[fetchXml] page=${page}`, fetchXml);

          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const result = (await window.dataverseAPI.fetchXmlQuery(fetchXml)) as any;
          const records: Record<string, unknown>[] = result?.value ?? [];
          allRecords.push(...records);
          moreRecords = Boolean(result?.["@Microsoft.Dynamics.CRM.morerecords"]);

          if (estimatedTotalRecords > 0) {
            const fetchPercentage = Math.round(
              (Math.min(allRecords.length, estimatedTotalRecords) / estimatedTotalRecords) * 100
            );
            setProgressPercentage(Math.max(0, Math.min(100, fetchPercentage)));
          }

          setProgressDescription(
            moreRecords
              ? `Fetching records... ${allRecords.length.toLocaleString()} / ${estimatedTotalRecords > 0 ? estimatedTotalRecords.toLocaleString() : "?"} (page ${page})`
              : `Fetched ${allRecords.length.toLocaleString()} records.`
          );
          page++;
        }

        const columnResults: ColumnTestResult[] = [];
        setProgressDescription(`Running tests on ${allRecords.length.toLocaleString()} records...`);
        const totalTestIterations = totalConfiguredTests * allRecords.length;
        const totalWorkUnits = allRecords.length + totalTestIterations;
        let currentTestIteration = 0;
        let lastYield = 0;

        for (const columnLogicalName of columnLogicalNames) {
          const tests = testsForTable[columnLogicalName];
          const attrRow = attributeRows.find((r) => r.logicalName === columnLogicalName);
          const displayName = attrRow?.displayName ?? columnLogicalName;

          for (const test of tests) {
            if (test.type === "containsData") {
              let passCount = 0;
              let failCount = 0;

              let recordIdx = 0;
              for (const record of allRecords) {
                recordIdx++;
                currentTestIteration++;
                const now = Date.now();
                if (now - lastYield > 100) {
                  lastYield = now;
                  const completedWorkUnits = allRecords.length + currentTestIteration;
                  const combinedPercentage = totalWorkUnits > 0
                    ? Math.round((completedWorkUnits / totalWorkUnits) * 100)
                    : 100;
                  setProgressPercentage(Math.max(0, Math.min(100, combinedPercentage)));
                  setProgressDescription(`Testing "${displayName}" — containsData: row ${recordIdx.toLocaleString()} / ${allRecords.length.toLocaleString()}`);
                  await new Promise<void>(resolve => setTimeout(resolve, 0));
                }
                const value = record[columnLogicalName];
                const hasData =
                  value !== null && value !== undefined && String(value).trim() !== "";
                if (hasData) passCount++;
                else failCount++;
              }

              columnResults.push({
                columnLogicalName,
                displayName,
                testType: "containsData",
                passCount,
                failCount,
                totalCount: allRecords.length,
              });
            }

            if (test.type === "matchesRegex") {
              let passCount = 0;
              let failCount = 0;

              let regex: RegExp;
              try {
                regex = new RegExp(test.regexPattern);
              } catch {
                setRunTestsMessage(
                  `Skipping invalid regex for ${displayName} (${columnLogicalName}): ${test.regexPattern}`
                );
                continue;
              }

              let recordIdx = 0;
              for (const record of allRecords) {
                recordIdx++;
                currentTestIteration++;
                const now = Date.now();
                if (now - lastYield > 100) {
                  lastYield = now;
                  const completedWorkUnits = allRecords.length + currentTestIteration;
                  const combinedPercentage = totalWorkUnits > 0
                    ? Math.round((completedWorkUnits / totalWorkUnits) * 100)
                    : 100;
                  setProgressPercentage(Math.max(0, Math.min(100, combinedPercentage)));
                  setProgressDescription(`Testing "${displayName}" — matchesRegex: row ${recordIdx.toLocaleString()} / ${allRecords.length.toLocaleString()}`);
                  await new Promise<void>(resolve => setTimeout(resolve, 0));
                }
                const value = record[columnLogicalName];
                const valueAsText = value === null || value === undefined ? "" : String(value);
                const isMatch = regex.test(valueAsText);
                if (isMatch) passCount++;
                else failCount++;
              }

              columnResults.push({
                columnLogicalName,
                displayName,
                testType: "matchesRegex",
                passCount,
                failCount,
                totalCount: allRecords.length,
              });
            }

            if (test.type === "matchesMetadata") {
              let passCount = 0;
              let failCount = 0;
              const attributeMetadata = attributeMetadataByLogicalName[columnLogicalName];
              const attributeType = getFieldType(attributeMetadata);
              const parseNumericRange = () => {
                const minValue = Number(attributeMetadata?.MinValue);
                const maxValue = Number(attributeMetadata?.MaxValue);

                return {
                  minValue: Number.isFinite(minValue) ? minValue : null,
                  maxValue: Number.isFinite(maxValue) ? maxValue : null,
                };
              };
              const optionValueSet = new Set<string>();
              const optionCollections = [
                attributeMetadata?.OptionSet?.Options,
                attributeMetadata?.GlobalOptionSet?.Options,
              ];

              optionCollections.forEach((options: any) => {
                if (!Array.isArray(options)) {
                  return;
                }

                options.forEach((option: any) => {
                  const optionValue = option?.Value;
                  if (optionValue !== null && optionValue !== undefined) {
                    optionValueSet.add(String(optionValue));
                  }
                });
              });

              let recordIdx = 0;
              for (const record of allRecords) {
                recordIdx++;
                currentTestIteration++;
                const now = Date.now();
                if (now - lastYield > 100) {
                  lastYield = now;
                  const completedWorkUnits = allRecords.length + currentTestIteration;
                  const combinedPercentage = totalWorkUnits > 0
                    ? Math.round((completedWorkUnits / totalWorkUnits) * 100)
                    : 100;
                  setProgressPercentage(Math.max(0, Math.min(100, combinedPercentage)));
                  setProgressDescription(`Testing "${displayName}" — matchesMetadata: row ${recordIdx.toLocaleString()} / ${allRecords.length.toLocaleString()}`);
                  await new Promise<void>(resolve => setTimeout(resolve, 0));
                }

                const value = record[columnLogicalName];

                // "Contains Data" is a separate rule; empty values are treated as metadata-valid.
                if (value === null || value === undefined || String(value).trim() === "") {
                  passCount++;
                  continue;
                }

                if (attributeType === "StringType") {
                  const maxLength = Number(attributeMetadata?.MaxLength);
                  if (!Number.isFinite(maxLength) || maxLength <= 0) {
                    passCount++;
                    continue;
                  }

                  const valueLength = String(value).length;
                  if (valueLength <= maxLength) passCount++;
                  else failCount++;
                  continue;
                }

                if (attributeType === "IntegerType") {
                  const numericValue = Number(value);
                  const { minValue, maxValue } = parseNumericRange();

                  if (!Number.isInteger(numericValue)) {
                    failCount++;
                    continue;
                  }

                  const meetsMin = minValue !== null ? numericValue >= minValue : true;
                  const meetsMax = maxValue !== null ? numericValue <= maxValue : true;

                  if (meetsMin && meetsMax) passCount++;
                  else failCount++;
                  continue;
                }

                if (attributeType === "DecimalType" || attributeType === "MoneyType") {
                  const numericValue = Number(value);
                  const { minValue, maxValue } = parseNumericRange();

                  if (!Number.isFinite(numericValue)) {
                    failCount++;
                    continue;
                  }

                  const meetsMin = minValue !== null ? numericValue >= minValue : true;
                  const meetsMax = maxValue !== null ? numericValue <= maxValue : true;

                  if (meetsMin && meetsMax) passCount++;
                  else failCount++;
                  continue;
                }

                if (
                  attributeType === "PicklistType" ||
                  attributeType === "StateType" ||
                  attributeType === "StatusType"
                ) {
                  if (optionValueSet.size === 0) {
                    passCount++;
                    continue;
                  }

                  const isValidOption = optionValueSet.has(String(value));

                  if (isValidOption) passCount++;
                  else failCount++;
                  continue;
                }

                // Until other metadata checks are implemented, unsupported types pass by default.
                passCount++;
              }

              columnResults.push({
                columnLogicalName,
                displayName,
                testType: "matchesMetadata",
                passCount,
                failCount,
                totalCount: allRecords.length,
              });
            }
          }
        }

        setTestResults({ totalRecords: allRecords.length, columnResults });
        setProgressPercentage(100);
        setProgressDescription("Completed.");
        setIsRunDialogOpen(false);
        setIsResultsFullscreen(true);
      } catch (error) {
        setRunTestsMessage(
          `Test execution failed: ${error instanceof Error ? error.message : String(error)}`
        );
      } finally {
        setIsRunningTests(false);
      }
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
    const allRowsSelected =
      sortedAttributeRows.length > 0 && selectedColumnLogicalNames.length === sortedAttributeRows.length;

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
                      setTestResults(null);
                      setIsResultsFullscreen(false);
                    }}
                  />
                </Stack>

                {showColumnsSection && (
                  <>
                    {isResultsFullscreen && testResults && !isRunningTests ? (
                      <Stack tokens={{ childrenGap: 12 }}>
                        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                          <Text variant="large">Latest Test Run Results</Text>
                          <PrimaryButton
                            text="Start Again"
                            onClick={() => {
                              setIsResultsFullscreen(false);
                              setTestResults(null);
                              setRunTestsMessage(null);
                              setProgressPercentage(0);
                              setProgressDescription("");
                            }}
                          />
                        </Stack>
                        <TestResultsPanel results={testResults} />
                      </Stack>
                    ) : (
                      <>
                    <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                      <Text variant="large">Columns for {selectedTable}</Text>
                      <Stack horizontal tokens={{ childrenGap: 8 }}>
                        <DefaultButton
                          text={allRowsSelected ? "Clear Selection" : "Select All"}
                          disabled={sortedAttributeRows.length === 0}
                          onClick={handleToggleSelectAll}
                        />
                        <DefaultButton
                          text={`Add Test To All Selected (${selectedColumnLogicalNames.length})`}
                          disabled={selectedColumnLogicalNames.length === 0}
                          onClick={handleBulkAddTests}
                        />
                        <DefaultButton
                          text="Clear All Tests"
                          disabled={totalSelectedTests === 0}
                          onClick={() => {
                            if (selectedTable) {
                              vm.clearAllTestsForTable(selectedTable);
                              setTestResults(null);
                            }
                          }}
                        />
                        <PrimaryButton
                          text="Run Selected Tests"
                          disabled={totalSelectedTests === 0 || isRunningTests}
                          onClick={handleRunSelectedTests}
                        />
                      </Stack>
                    </Stack>

                    {attributesLoading && <Spinner label="Loading columns" />}

                    {attributesError && (
                      <MessageBar messageBarType={MessageBarType.error}>
                        Failed to load columns: {attributesError}
                      </MessageBar>
                    )}

                    {!attributesLoading && !attributesError && attributeRows.length > 0 && (
                      <TextField
                        placeholder="Search by logical or display name..."
                        value={columnFilter}
                        onChange={(_, value) => setColumnFilter(value ?? "")}
                        styles={{ root: { maxWidth: 500 } }}
                      />
                    )}

                    {!attributesLoading && !attributesError && attributeRows.length === 0 && (
                      <MessageBar messageBarType={MessageBarType.warning}>
                        No columns were returned for this table.
                      </MessageBar>
                    )}

                    {!attributesLoading && !attributesError && attributeRows.length > 0 && columnFilter.trim() !== "" && sortedAttributeRows.length === 0 && (
                      <MessageBar messageBarType={MessageBarType.info}>
                        No columns match your filter.
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

                      <Stack horizontal tokens={{ childrenGap: 8 }}>
                        <DefaultButton
                          text="Clear All Tests"
                          disabled={totalSelectedTests === 0}
                          onClick={() => {
                            if (selectedTable) {
                              vm.clearAllTestsForTable(selectedTable);
                              setTestResults(null);
                            }
                          }}
                        />
                        <PrimaryButton
                          text="Run Selected Tests"
                          disabled={totalSelectedTests === 0 || isRunningTests}
                          onClick={handleRunSelectedTests}
                        />
                      </Stack>
                    </Stack>

                    {!isRunDialogOpen && runTestsMessage && (
                      <MessageBar messageBarType={MessageBarType.error}>{runTestsMessage}</MessageBar>
                    )}
                      </>
                    )}
                  </>
                )}
              </>
            )}
          </Stack>
        </div>

        <EditTestsDialog
          isOpen={isAddTestDialogOpen}
          dialogMode={testDialogMode}
          activeColumnLogicalName={activeColumnLogicalName}
          selectedColumnCount={selectedColumnLogicalNames.length}
          addTestDialogError={addTestDialogError}
          isMatchesRegexEnabled={isMatchesRegexEnabled}
          regexPattern={regexPattern}
          isMatchesMetadataEnabled={isMatchesMetadataEnabled}
          isContainsDataEnabled={isContainsDataEnabled}
          onDismiss={handleAddTestDialogDismiss}
          onSave={handleSaveTest}
          onMatchesRegexChange={(enabled) => {
            setIsMatchesRegexEnabled(enabled);
            if (!enabled) {
              setRegexPattern("");
            }
            setAddTestDialogError(null);
          }}
          onRegexPatternChange={(pattern) => {
            setRegexPattern(pattern);
            setAddTestDialogError(null);
          }}
          onMatchesMetadataChange={(enabled) => {
            setIsMatchesMetadataEnabled(enabled);
            setAddTestDialogError(null);
          }}
          onContainsDataChange={(enabled) => {
            setIsContainsDataEnabled(enabled);
            setAddTestDialogError(null);
          }}
        />

        <Dialog
          hidden={!isRunDialogOpen}
          onDismiss={() => setIsRunDialogOpen(false)}
          modalProps={{ isBlocking: isRunningTests }}
          dialogContentProps={{
            type: DialogType.largeHeader,
            title: "Running Selected Tests",
            subText: "Fetching and validating records. This window updates as progress is made.",
          }}
        >
          <Stack tokens={{ childrenGap: 12 }}>
            <ProgressIndicator
              label={progressDescription || (isRunningTests ? "Starting..." : "Done.")}
              percentComplete={Math.max(0, Math.min(progressPercentage, 100)) / 100}
              description={`${Math.max(0, Math.min(progressPercentage, 100))}%`}
            />

            {runTestsMessage && (
              <MessageBar messageBarType={MessageBarType.error}>{runTestsMessage}</MessageBar>
            )}

            {!runTestsMessage && !isRunningTests && (
              <MessageBar messageBarType={MessageBarType.warning}>
                Test run completed.
              </MessageBar>
            )}
          </Stack>

          <DialogFooter>
            <DefaultButton
              text="Close"
              onClick={() => setIsRunDialogOpen(false)}
              disabled={isRunningTests}
            />
          </DialogFooter>
        </Dialog>
      </>
    );
});