import React from "react";

import { observer } from "mobx-react";
import { Table, ViewModel, LabelOptions, LanguageDef } from "../model/viewModel";
import { dvService } from "../services/dataverseService";
import {
  Button,
  Checkbox,
  Combobox,
  Divider,
  Dropdown,
  Field,
  List,
  ListItem,
  Option,
  ProgressBar,
  Radio,
  RadioGroup,
  SelectionItemId,
  Toolbar,
  ToolbarGroup,
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
export const ExportPanel = observer((props: ExportPanelProps): React.JSX.Element => {
  const { dvSvc, vm, lgSvc, onLog } = props;
  const [solutions, setSolutions] = React.useState<Solution[]>([]);
  const [selectedSolution, setSelectedSolution] = React.useState<Solution | null>(null);
  const [selectedTables, setSelectedTables] = React.useState<SelectionItemId[]>([]);
  const [tables, setTables] = React.useState<Table[]>([]);
  const [labelOption, setLabelOption] = React.useState<string>(String(vm.options.labelOptions));
  const [languages, setLanguages] = React.useState<LanguageDef[]>([]);

  React.useEffect(() => {
    onLog("Export panel loaded");
    const fetchSolutions = async () => {
      try {
        const sols = await dvSvc.getSolutions(); // Assumes getSolutions returns array of { uniqueName, friendlyName }
        setSolutions(sols);
        onLog("Solutions loaded", "success");
        const langs = await dvSvc.getLanguages();
        console.log("langs", langs);
        setLanguages(langs);
        vm.allLanguages = langs;
        onLog("Languages loaded", "success");
      } catch (err) {
        onLog("Failed to load solutions", "error");
      }
    };
    fetchSolutions();
  }, [dvSvc, onLog]);

  React.useEffect(() => {
    if (!selectedSolution) {
      setTables([]);
      return;
    }
    const fetchTables = async () => {
      onLog(`Loading tables for ${selectedSolution.name}`);
      try {
        const solutionTables = await dvSvc.getSolutionTables(selectedSolution.solutionId);
        setTables(solutionTables);
        onLog("Tables loaded", "success");
      } catch (err) {
        onLog("Failed to load tables", "error");
      }
    };

    fetchTables();
  }, [dvSvc, onLog, selectedSolution]);

  async function exportExcel() {
    vm.selectedTables = tables.filter((table) => selectedTables.includes(table.id));
    //await window.toolboxAPI.utils.showLoading("Exporting translations...");
    await lgSvc.exportTranslations();
    //window.toolboxAPI.utils.hideLoading();
  }
  function selectLang(code?: string): void {
    console.log(code);
    vm.selectedLanguage = vm.allLanguages?.find((lang) => lang.code === code);
  }

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100%", gap: "16px" }}>
      <Toolbar style={{ justifyContent: "space-between" }}>
        <ToolbarGroup>
          <Combobox
            placeholder="Select a Solution"
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
        </ToolbarGroup>
        <ToolbarGroup>
          <Button icon={<ArrowExportUpFilled />} disabled={selectedTables.length === 0} onClick={exportExcel}>
            Export
          </Button>
        </ToolbarGroup>
      </Toolbar>
      {vm.exporting && (
        <Field style={{margin: "20px"}} validationMessage={vm.message} validationState="none">
          <ProgressBar thickness="large" value={vm.exportpercentage}></ProgressBar>
        </Field>
      )}
      {!vm.exporting && (
        <div style={{ display: "grid", gridTemplateColumns: "1fr 2fr", gap: "16px", flex: 1, minHeight: 0 }}>
          <div style={{ overflow: "auto" }}>
            <List
              selectionMode="multiselect"
              selectedItems={selectedTables}
              onSelectionChange={(_, data) => setSelectedTables(data.selectedItems)}
              aria-label="List of Tables"
            >
              {tables.map((table) => (
                <ListItem key={table.id} value={table.id}>
                  {table.label}
                </ListItem>
              ))}
            </List>
          </div>
          <div style={{ overflow: "auto", height: "100%" }}>
            <Divider>Global Options</Divider>
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
              {languages.length}
              <RadioGroup
                value={labelOption}
                defaultValue={String(LabelOptions.both)}
                onChange={(_, data) => {
                  setLabelOption(data.value);
                  vm.options.labelOptions = data.value as unknown as LabelOptions;
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
