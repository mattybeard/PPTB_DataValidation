import React from "react";
import {
  Checkbox,
  DefaultButton,
  Dialog,
  DialogFooter,
  DialogType,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  Stack,
  TextField,
} from "@fluentui/react";

export interface EditTestsDialogProps {
  isOpen: boolean;
  dialogMode: "single" | "bulk";
  activeColumnLogicalName: string | null;
  selectedColumnCount: number;
  addTestDialogError: string | null;
  isMatchesRegexEnabled: boolean;
  regexPattern: string;
  isMatchesMetadataEnabled: boolean;
  isContainsDataEnabled: boolean;
  onDismiss: () => void;
  onSave: () => void;
  onMatchesRegexChange: (enabled: boolean) => void;
  onRegexPatternChange: (pattern: string) => void;
  onMatchesMetadataChange: (enabled: boolean) => void;
  onContainsDataChange: (enabled: boolean) => void;
}

export const EditTestsDialog = ({
  isOpen,
  dialogMode,
  activeColumnLogicalName,
  selectedColumnCount,
  addTestDialogError,
  isMatchesRegexEnabled,
  regexPattern,
  isMatchesMetadataEnabled,
  isContainsDataEnabled,
  onDismiss,
  onSave,
  onMatchesRegexChange,
  onRegexPatternChange,
  onMatchesMetadataChange,
  onContainsDataChange,
}: EditTestsDialogProps): React.ReactElement => {
  return (
    <Dialog
      hidden={!isOpen}
      onDismiss={onDismiss}
      dialogContentProps={{
        type: DialogType.normal,
        title: "Edit Test(s)",
        subText:
          dialogMode === "bulk"
            ? `Apply the same tests to ${selectedColumnCount} selected column(s).`
            : activeColumnLogicalName
              ? `Configure tests for ${activeColumnLogicalName}.`
              : "Create a test for this column.",
      }}
    >
      <Stack tokens={{ childrenGap: 12 }}>
        {addTestDialogError && (
          <MessageBar messageBarType={MessageBarType.error}>{addTestDialogError}</MessageBar>
        )}

        <Checkbox
          label="Matches Regex"
          checked={isMatchesRegexEnabled}
          onChange={(_, checked) => {
            const isChecked = Boolean(checked);
            onMatchesRegexChange(isChecked);
          }}
        />

        {isMatchesRegexEnabled && (
          <TextField
            label="Regex Pattern"
            placeholder="e.g. ^[A-Z]{3}\\d{4}$"
            value={regexPattern}
            styles={{ root: { paddingLeft: 28 } }}
            onChange={(_, value) => {
              onRegexPatternChange(value ?? "");
            }}
          />
        )}

        <Checkbox
          label="Matches Metadata"
          checked={isMatchesMetadataEnabled}
          onChange={(_, checked) => {
            onMatchesMetadataChange(Boolean(checked));
          }}
        />

        <Checkbox
          label="Contains Data"
          checked={isContainsDataEnabled}
          onChange={(_, checked) => {
            onContainsDataChange(Boolean(checked));
          }}
        />
      </Stack>

      <DialogFooter>
        <PrimaryButton text="Save Test" onClick={onSave} />
        <DefaultButton text="Cancel" onClick={onDismiss} />
      </DialogFooter>
    </Dialog>
  );
};
