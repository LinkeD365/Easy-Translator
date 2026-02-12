import React from "react";

import { observer } from "mobx-react";
import { Table, ViewModel, LabelOptions, LanguageDef } from "../model/viewModel";
import { dvService } from "../services/dataverseService";
import {
  Button,
  Caption1,
  Checkbox,
  Combobox,
  Divider,
  Dropdown,
  Field,
  Input,
  List,
  ListItem,
  Option,
  ProgressBar,
  Radio,
  RadioGroup,
  SelectionItemId,
  Toolbar,
  ToolbarGroup,
  tokens,
} from "@fluentui/react-components";
import { Solution } from "../model/solution";
import { ArrowExportUpFilled } from "@fluentui/react-icons";
import { languageService } from "../services/languageService";

interface ExportPanelProps {
  dvSvc: dvService;
  vm: ViewModel;
  lgSvc: languageService;
  onLog: (message: string, type?: "info" | "success" | "warning" | "error") => void;
}

interface EmptyStateProps {
  message: string;
}

const EmptyState: React.FC<EmptyStateProps> = ({ message }) => (
  <div
    style={{
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      height: "100%",
      padding: "20px",
      textAlign: "center",
    }}
  >
    <Caption1>{message}</Caption1>
  </div>
);

function useDebounce<T>(value: T, delay: number): T {
  const [debouncedValue, setDebouncedValue] = React.useState(value);

  React.useEffect(() => {
    const timer = setTimeout(() => setDebouncedValue(value), delay);
    return () => clearTimeout(timer);
  }, [value, delay]);

  return debouncedValue;
}

export const ExportPanel = observer((props: ExportPanelProps): React.JSX.Element => {
  const { dvSvc, vm, lgSvc, onLog } = props;
  const [solutions, setSolutions] = React.useState<Solution[]>([]);
  const [selectedSolution, setSelectedSolution] = React.useState<Solution | null>(null);
  const [selectedTables, setSelectedTables] = React.useState<SelectionItemId[]>([]);
  const [tables, setTables] = React.useState<Table[]>([]);
  const [filteredTables, setFilteredTables] = React.useState<Table[]>([]);
  const [searchQuery, setSearchQuery] = React.useState<string>("");
  const [loading, setLoading] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string | null>(null);
  const [labelOption, setLabelOption] = React.useState<string>(String(vm.options.labelOptions));
  const [languages, setLanguages] = React.useState<LanguageDef[]>([]);

  const debouncedSearch = useDebounce(searchQuery, 300);

  React.useEffect(() => {
    onLog("Export panel loaded");
    const fetchSolutions = async () => {
      try {
        const sols = await dvSvc.getSolutions();
        setSolutions(sols);
        onLog("Solutions loaded", "success");
        const langs = await dvSvc.getLanguages();
        setLanguages(langs);
        vm.allLanguages = langs;
        onLog("Languages loaded", "success");
      } catch (err) {
        const errorMsg = err instanceof Error ? err.message : "Failed to load initial data";
        setError(errorMsg);
        onLog(errorMsg, "error");
      }
    };
    fetchSolutions();
  }, [dvSvc, onLog, vm]);

  React.useEffect(() => {
    if (debouncedSearch.trim() === "") {
      setFilteredTables(tables);
    } else {
      const query = debouncedSearch.toLowerCase();
      setFilteredTables(
        tables.filter(
          (table) => table.label.toLowerCase().includes(query) || table.logicalName.toLowerCase().includes(query),
        ),
      );
    }
  }, [debouncedSearch, tables]);

  React.useEffect(() => {
    if (!selectedSolution) {
      setTables([]);
      setFilteredTables([]);
      setLoading(false);
      setError(null);
      return;
    }
    const fetchTables = async () => {
      setLoading(true);
      setError(null);
      onLog(`Loading tables for ${selectedSolution.name}`);
      try {
        const solutionTables = await dvSvc.getSolutionTables(selectedSolution.solutionId);
        setTables(solutionTables);
        setFilteredTables(solutionTables);
        setSearchQuery("");
        onLog("Tables loaded", "success");
      } catch (err) {
        const errorMsg = err instanceof Error ? err.message : "Failed to load tables";
        setError(errorMsg);
        onLog(errorMsg, "error");
      } finally {
        setLoading(false);
      }
    };

    fetchTables();
  }, [dvSvc, onLog, selectedSolution]);

  const exportExcel = React.useCallback(async () => {
    try {
      vm.selectedTables = tables.filter((table) => selectedTables.includes(table.id));
      vm.solution = selectedSolution || undefined;
      await lgSvc.exportTranslations();
    } catch (err) {
      const errorMsg = err instanceof Error ? err.message : "Export failed";
      setError(errorMsg);
      onLog(errorMsg, "error");
    }
  }, [tables, selectedTables, selectedSolution, vm, lgSvc, onLog]);

  const loadAllTables = React.useCallback(async () => {
    setLoading(true);
    setSelectedSolution(null);
    setError(null);
    onLog(`Loading all tables from environment`);
    try {
      const allTables = await dvSvc.getAllTables();
      setTables(allTables);
      setFilteredTables(allTables);
      setSearchQuery("");
      onLog("All tables loaded", "success");
    } catch (err) {
      const errorMsg = err instanceof Error ? err.message : "Failed to load all tables";
      setError(errorMsg);
      onLog(errorMsg, "error");
    } finally {
      setLoading(false);
    }
  }, [dvSvc, onLog]);

  const selectLang = React.useCallback(
    (code?: string): void => {
      vm.selectedLanguage = vm.allLanguages?.find((lang) => lang.code === code);
    },
    [vm],
  );

  function toggleAllGlobalOptions(checked: boolean): void {
    vm.options.optionSets = checked;
    vm.options.globalOptionSets = checked;
    vm.options.siteMaps = checked;
    vm.options.dashboards = checked;
  }

  function allGlobalOptionsChecked(): boolean {
    return !!vm.options.optionSets && !!vm.options.globalOptionSets && !!vm.options.siteMaps && !!vm.options.dashboards;
  }

  function someGlobalOptionsChecked(): boolean {
    return !!vm.options.optionSets || !!vm.options.globalOptionSets || !!vm.options.siteMaps || !!vm.options.dashboards;
  }

  function toggleAllTableOptions(checked: boolean): void {
    vm.options.table = checked;
    vm.options.fields = checked;
    vm.options.localOptionSets = checked;
    vm.options.booleanOptions = checked;
    vm.options.views = checked;
    vm.options.charts = checked;
    vm.options.forms = checked;
    vm.options.formTabs = checked;
    vm.options.formSections = checked;
    vm.options.formFields = checked;
    vm.options.relationships = checked;
  }

  function allTableOptionsChecked(): boolean {
    return (
      !!vm.options.table &&
      !!vm.options.fields &&
      !!vm.options.localOptionSets &&
      !!vm.options.booleanOptions &&
      !!vm.options.views &&
      !!vm.options.charts &&
      !!vm.options.forms &&
      !!vm.options.formTabs &&
      !!vm.options.formSections &&
      !!vm.options.formFields &&
      !!vm.options.relationships
    );
  }

  function someTableOptionsChecked(): boolean {
    return (
      !!vm.options.table ||
      !!vm.options.fields ||
      !!vm.options.localOptionSets ||
      !!vm.options.booleanOptions ||
      !!vm.options.views ||
      !!vm.options.charts ||
      !!vm.options.forms ||
      !!vm.options.formTabs ||
      !!vm.options.formSections ||
      !!vm.options.formFields ||
      !!vm.options.relationships
    );
  }

  const toggleAllTables = React.useCallback(
    (checked: boolean): void => {
      if (checked) {
        setSelectedTables(filteredTables.map((table) => table.id));
      } else {
        setSelectedTables([]);
      }
    },
    [filteredTables],
  );

  const allTablesChecked = React.useCallback((): boolean => {
    if (filteredTables.length === 0) return false;
    return filteredTables.every((table) => selectedTables.includes(table.id));
  }, [filteredTables, selectedTables]);

  const someTablesChecked = React.useCallback((): boolean => {
    if (filteredTables.length === 0) return false;
    const selectedInFiltered = filteredTables.filter((table) => selectedTables.includes(table.id)).length;
    return selectedInFiltered > 0 && selectedInFiltered < filteredTables.length;
  }, [filteredTables, selectedTables]);

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100%", gap: "16px" }}>
      <Toolbar style={{ justifyContent: "space-between" }}>
        <ToolbarGroup>
          <Combobox
            placeholder="Select a Solution"
            disabled={vm.exporting}
            value={selectedSolution?.name}
            onOptionSelect={(_, data) =>
              setSelectedSolution(solutions.find((sol) => sol.uniqueName === data.optionValue) || null)
            }
          >
            {solutions.map((sol) => (
              <Option key={sol.uniqueName} value={sol.uniqueName}>
                {sol.name}
              </Option>
            ))}
          </Combobox>
          <Button appearance="secondary" disabled={vm.exporting || loading} onClick={loadAllTables}>
            All Tables
          </Button>
        </ToolbarGroup>
        <ToolbarGroup>
          <Button
            icon={<ArrowExportUpFilled />}
            appearance="subtle"
            disabled={selectedTables.length === 0 || vm.exporting}
            onClick={exportExcel}
          >
            Export
          </Button>
        </ToolbarGroup>
      </Toolbar>
      {error && (
        <div
          style={{
            padding: "12px",
            backgroundColor: tokens.colorPaletteRedBackground2,
            color: tokens.colorPaletteRedForeground1,
            borderRadius: "4px",
            margin: "0 16px",
          }}
        >
          {error}
        </div>
      )}
      {vm.exporting && (
        <Field style={{ margin: "20px" }} validationMessage={vm.message} validationState="none">
          <ProgressBar thickness="large" value={vm.exportpercentage}></ProgressBar>
        </Field>
      )}
      {!vm.exporting && (
        <div style={{ display: "grid", gridTemplateColumns: "1fr 2fr", gap: "16px", flex: 1, minHeight: 0 }}>
          <div style={{ overflow: "auto", border: `1px solid ${tokens.colorNeutralStroke1}`, borderRadius: "4px" }}>
            {tables.length === 0 && !loading ? (
              <EmptyState message="Please select a solution or click All Tables to view available tables" />
            ) : loading ? (
              <EmptyState message="Loading tables..." />
            ) : (
              <div style={{ display: "flex", flexDirection: "column", height: "100%" }}>
                <div
                  style={{
                    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
                    display: "flex",
                    alignItems: "center",
                    gap: "8px",
                  }}
                >
                  <Checkbox
                    checked={allTablesChecked() ? true : someTablesChecked() ? "mixed" : false}
                    onChange={(_, data) => toggleAllTables(data.checked === true)}
                  />
                  <Input
                    placeholder="Search tables..."
                    value={searchQuery}
                    onChange={(_, data) => setSearchQuery(data.value)}
                    style={{ flex: 1 }}
                    aria-label="Search tables"
                  />
                </div>
                {filteredTables.length === 0 && searchQuery.trim() !== "" ? (
                  <EmptyState message={`No tables found matching "${searchQuery}"`} />
                ) : (
                  <List
                    selectionMode="multiselect"
                    selectedItems={selectedTables}
                    onSelectionChange={(_, data) => setSelectedTables(data.selectedItems)}
                    aria-label="List of Tables"
                    style={{ flex: 1, overflow: "auto" }}
                  >
                    {filteredTables.map((table) => (
                      <ListItem
                        key={table.id}
                        value={table.id}
                        style={{ whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}
                      >
                        {table.label} <span style={{ fontSize: "0.85em", color: "gray" }}> ({table.logicalName})</span>
                      </ListItem>
                    ))}
                  </List>
                )}
              </div>
            )}
          </div>
          <div style={{ overflow: "auto", height: "100%" }}>
            <Divider>Global Options</Divider>
            <Checkbox
              checked={allGlobalOptionsChecked() ? true : someGlobalOptionsChecked() ? "mixed" : false}
              label="Select All"
              onChange={(_, data) => toggleAllGlobalOptions(data.checked === true)}
            />
            <Checkbox
              checked={!!vm.options.optionSets}
              label="Export Option Sets"
              onChange={(_, data) => (vm.options.optionSets = data.checked === true)}
            />
            <Checkbox
              checked={!!vm.options.globalOptionSets}
              label="Export Global Option Sets"
              onChange={(_, data) => (vm.options.globalOptionSets = data.checked === true)}
            />
            <Checkbox
              checked={!!vm.options.siteMaps}
              label="Export Site Maps"
              onChange={(_, data) => (vm.options.siteMaps = data.checked === true)}
            />
            <Checkbox
              checked={!!vm.options.dashboards}
              label="Export Dashboards"
              onChange={(_, data) => (vm.options.dashboards = data.checked === true)}
            />
            <Divider>Table Related Options</Divider>
            <Checkbox
              checked={allTableOptionsChecked() ? true : someTableOptionsChecked() ? "mixed" : false}
              label="Select All"
              onChange={(_, data) => toggleAllTableOptions(data.checked === true)}
            />
            <Checkbox
              checked={!!vm.options.table}
              label="Export Tables"
              onChange={(_, data) => (vm.options.table = data.checked === true)}
            />
            <Checkbox
              checked={!!vm.options.fields}
              label="Export Fields"
              onChange={(_, data) => (vm.options.fields = data.checked === true)}
            />
            <Checkbox
              checked={!!vm.options.localOptionSets}
              label="Export Local Option Sets"
              onChange={(_, data) => (vm.options.localOptionSets = data.checked === true)}
            />
            <Checkbox
              checked={!!vm.options.booleanOptions}
              label="Export Boolean Options"
              onChange={(_, data) => (vm.options.booleanOptions = data.checked === true)}
            />
            <Checkbox
              checked={!!vm.options.views}
              label="Export Views"
              onChange={(_, data) => (vm.options.views = data.checked === true)}
            />
            <Checkbox
              checked={!!vm.options.charts}
              label="Export Charts"
              onChange={(_, data) => (vm.options.charts = data.checked === true)}
            />
            <Checkbox
              checked={!!vm.options.forms}
              label="Export Forms"
              onChange={(_, data) => (vm.options.forms = data.checked === true)}
            />
            <Checkbox
              checked={!!vm.options.formTabs}
              label="Export Form Tabs"
              onChange={(_, data) => (vm.options.formTabs = data.checked === true)}
            />
            <Checkbox
              checked={!!vm.options.formSections}
              label="Export Form Sections"
              onChange={(_, data) => (vm.options.formSections = data.checked === true)}
            />
            <Checkbox
              checked={!!vm.options.formFields}
              label="Export Form Fields"
              onChange={(_, data) => (vm.options.formFields = data.checked === true)}
            />
            <Checkbox
              checked={!!vm.options.relationships}
              label="Export Relationships"
              onChange={(_, data) => (vm.options.relationships = data.checked === true)}
            />
            <Divider>Label & Languages</Divider>

            <Field label="Label Options">
              <RadioGroup
                value={labelOption}
                defaultValue={String(LabelOptions.both)}
                onChange={(_, data) => {
                  setLabelOption(data.value);
                  vm.options.labelOptions = parseInt(data.value) as LabelOptions;
                }}
              >
                <Radio value={String(LabelOptions.both)} label="Both" />
                <Radio value={String(LabelOptions.names)} label="Names" />
                <Radio value={String(LabelOptions.descriptions)} label="Descriptions" />
              </RadioGroup>
            </Field>
            <Checkbox
              checked={!!vm.options.exportAllLanguages}
              label="Export All Languages"
              onChange={(_, data) => (vm.options.exportAllLanguages = data.checked === true)}
            />
            <Dropdown
              disabled={vm.options.exportAllLanguages}
              onOptionSelect={(_, data) => selectLang(data.optionValue)}
            >
              {languages.map((lang) => (
                <Option key={lang.code} value={lang.code} text={lang.name}>
                  {lang.name}
                </Option>
              ))}
            </Dropdown>
          </div>
        </div>
      )}
    </div>
  );
});
