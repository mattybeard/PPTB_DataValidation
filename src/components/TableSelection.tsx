import { observer } from "mobx-react";
import { useContext } from "react";
import { ServiceProviderContext } from "../framework/ServiceProvider";
import { ViewModel } from "../models/ViewModel";
import { Dropdown, DropdownMenuItemType, IDropdownOption, IDropdownStyles, IStackTokens, Spinner, Stack } from "@fluentui/react";

export const TableSelection = observer(
  (): React.ReactElement => {
    const context = useContext(ServiceProviderContext);
    const vm = context.get<ViewModel>("vm");

    const stackTokens: IStackTokens = { childrenGap: 20 };
    const dropdownStyles: Partial<IDropdownStyles> = {
      dropdown: { width: 300 },
  };

    const generateOptions = (): IDropdownOption[] => {
      if (!vm.metadata) return [];
      const options: IDropdownOption[] = [];
      options.push({ key: 'headerRow', text: 'Tables', itemType: DropdownMenuItemType.Header });
      
      vm.metadata.value.forEach((entity) => {
        options.push({ key: entity.LogicalName, text: entity.LogicalName });
      });
      return options;
    };

    return (
    <>
      <div className="card">
        <Stack tokens={stackTokens}>
          {!vm.metadataLoaded && (
            <>
              <div>
                <Spinner label="Loading Metadata" />
              </div>
            </>
          )}
          {vm.metadataLoaded && (
            <>
              <Dropdown
              placeholder="Select a table"
              label="Select a table"
              options={generateOptions()}
              styles={dropdownStyles}
              />
            </>
          )}
        </Stack>
      </div>    
    </>
    );
});