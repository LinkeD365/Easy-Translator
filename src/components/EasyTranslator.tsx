import React from "react";

import { observer } from "mobx-react";
import { ViewModel } from "../model/viewModel";
import { dvService } from "../services/dataverseService";
import { Tab, TabList } from "@fluentui/react-components";
import { ExportPanel } from "./ExportPanel";
import { languageService } from "../services/languageService";

interface EasyTranslatorProps {
  dvSvc: dvService;
  vm: ViewModel;
  lgSvc: languageService;
  onLog: (message: string, type?: "info" | "success" | "warning" | "error") => void;
}
export const EasyTranslator = observer((props: EasyTranslatorProps): React.JSX.Element => {
  const { dvSvc, vm, lgSvc, onLog } = props;

  const [selectedTab, setSelectedTab] = React.useState<string>("export");

  const tabs = (
    <TabList selectedValue={selectedTab} onTabSelect={(_, data) => setSelectedTab(data.value as string)}>
      <Tab value="export">Export</Tab>
      <Tab value="import">Import</Tab>
    </TabList>
  );
  return (
    <div style={{ height: "100vh", display: "flex", flexDirection: "column" }}>
      {tabs}
      {selectedTab === "export" && <ExportPanel dvSvc={dvSvc} lgSvc={lgSvc} vm={vm} onLog={onLog} />}
    </div>
  );
});
