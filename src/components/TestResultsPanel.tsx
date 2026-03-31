import React, { useEffect, useMemo, useState } from "react";
import {
  DetailsList,
  IColumn,
  MessageBar,
  MessageBarType,
  SelectionMode,
  Stack,
} from "@fluentui/react";

export interface ColumnTestResult {
  columnLogicalName: string;
  displayName: string;
  testType: string;
  passCount: number;
  failCount: number;
  totalCount: number;
}

export interface TestRunResults {
  totalRecords: number;
  columnResults: ColumnTestResult[];
}

type ResultSortKey = "column" | "test" | "passed" | "failed" | "rate";

const getTestTypeLabel = (testType: string): string => {
  if (testType === "containsData") return "Contains Data";
  if (testType === "matchesRegex") return "Matches Regex";
  if (testType === "matchesMetadata") return "Matches Metadata";
  return testType;
};

export const TestResultsPanel = ({
  results,
}: {
  results: TestRunResults;
}): React.ReactElement => {
  const [sortKey, setSortKey] = useState<ResultSortKey>("rate");
  const [isSortDescending, setIsSortDescending] = useState(false);

  // Default to failed % descending, which is equivalent to pass % ascending.
  useEffect(() => {
    setSortKey("rate");
    setIsSortDescending(false);
  }, [results]);

  const totalChecks = results.columnResults.reduce((sum, item) => sum + item.totalCount, 0);
  const totalPassed = results.columnResults.reduce((sum, item) => sum + item.passCount, 0);
  const overallRate = totalChecks > 0 ? Math.round((totalPassed / totalChecks) * 100) : 0;
  const barType =
    overallRate >= 90
      ? MessageBarType.success
      : overallRate >= 70
        ? MessageBarType.warning
        : MessageBarType.error;

  const getRate = (item: ColumnTestResult): number =>
    item.totalCount > 0 ? item.passCount / item.totalCount : 0;

  const sortedItems = useMemo(() => {
    const items = [...results.columnResults];

    items.sort((left, right) => {
      if (sortKey === "column") {
        const l = `${left.displayName} (${left.columnLogicalName})`.toLowerCase();
        const r = `${right.displayName} (${right.columnLogicalName})`.toLowerCase();
        if (l < r) return isSortDescending ? 1 : -1;
        if (l > r) return isSortDescending ? -1 : 1;
        return 0;
      }

      if (sortKey === "test") {
        const l = getTestTypeLabel(left.testType).toLowerCase();
        const r = getTestTypeLabel(right.testType).toLowerCase();
        if (l < r) return isSortDescending ? 1 : -1;
        if (l > r) return isSortDescending ? -1 : 1;
        return 0;
      }

      if (sortKey === "passed") {
        if (left.passCount < right.passCount) return isSortDescending ? 1 : -1;
        if (left.passCount > right.passCount) return isSortDescending ? -1 : 1;
        return 0;
      }

      if (sortKey === "failed") {
        if (left.failCount < right.failCount) return isSortDescending ? 1 : -1;
        if (left.failCount > right.failCount) return isSortDescending ? -1 : 1;
        return 0;
      }

      const l = getRate(left);
      const r = getRate(right);
      if (l < r) return isSortDescending ? 1 : -1;
      if (l > r) return isSortDescending ? -1 : 1;
      return 0;
    });

    return items;
  }, [results.columnResults, sortKey, isSortDescending]);

  const handleColumnClick = (
    _event?: React.MouseEvent<HTMLElement>,
    column?: IColumn
  ): void => {
    const clickedKey = column?.key as ResultSortKey | undefined;
    if (!clickedKey) {
      return;
    }

    if (clickedKey === sortKey) {
      setIsSortDescending((previous) => !previous);
      return;
    }

    setSortKey(clickedKey);
    setIsSortDescending(false);
  };

  const columns: IColumn[] = [
    {
      key: "column",
      name: "Column",
      minWidth: 180,
      maxWidth: 280,
      isResizable: true,
      isSorted: sortKey === "column",
      isSortedDescending: isSortDescending,
      onColumnClick: handleColumnClick,
      onRender: (item: ColumnTestResult) => `${item.displayName} (${item.columnLogicalName})`,
    },
    {
      key: "test",
      name: "Test",
      minWidth: 110,
      maxWidth: 140,
      isSorted: sortKey === "test",
      isSortedDescending: isSortDescending,
      onColumnClick: handleColumnClick,
      onRender: (item: ColumnTestResult) => getTestTypeLabel(item.testType),
    },
    {
      key: "passed",
      name: "Passed",
      minWidth: 65,
      maxWidth: 80,
      isSorted: sortKey === "passed",
      isSortedDescending: isSortDescending,
      onColumnClick: handleColumnClick,
      onRender: (item: ColumnTestResult) => String(item.passCount),
    },
    {
      key: "failed",
      name: "Failed",
      minWidth: 65,
      maxWidth: 80,
      isSorted: sortKey === "failed",
      isSortedDescending: isSortDescending,
      onColumnClick: handleColumnClick,
      onRender: (item: ColumnTestResult) => String(item.failCount),
    },
    {
      key: "rate",
      name: "Pass %",
      minWidth: 65,
      maxWidth: 80,
      isSorted: sortKey === "rate",
      isSortedDescending: isSortDescending,
      onColumnClick: handleColumnClick,
      onRender: (item: ColumnTestResult) =>
        `${item.totalCount > 0 ? Math.round((item.passCount / item.totalCount) * 100) : 0}%`,
    },
  ];

  return (
    <Stack tokens={{ childrenGap: 8 }}>
      <MessageBar messageBarType={barType}>
        <strong>
          {results.totalRecords} records checked - {overallRate}% pass rate ({totalPassed}/
          {totalChecks} checks)
        </strong>
      </MessageBar>
      <DetailsList
        items={sortedItems}
        setKey="testResults"
        columns={columns}
        compact
        selectionMode={SelectionMode.none}
      />
    </Stack>
  );
};
