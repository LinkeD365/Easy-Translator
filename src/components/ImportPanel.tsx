import React from "react";
import { observer } from "mobx-react";
import { ViewModel } from "../model/viewModel";
import { dvService } from "../services/dataverseService";
import { Button, Caption1, Caption2, Field, Input, Label, ProgressBar, tokens } from "@fluentui/react-components";
import { languageService } from "../services/languageService";
import { ArrowUploadFilled, FolderOpenRegular } from "@fluentui/react-icons";

interface ImportPanelProps {
  dvSvc: dvService;
  vm: ViewModel;
  lgSvc: languageService;
  onLog: (message: string, type?: "info" | "success" | "warning" | "error") => void;
}

export const ImportPanel = observer((props: ImportPanelProps): React.JSX.Element => {
  const { vm, onLog } = props;
  const [batchCount, setBatchCount] = React.useState<string>("10");
  const [selectedFile, setSelectedFile] = React.useState<File | null>(null);
  const [isDarkTheme, setIsDarkTheme] = React.useState(() => {
    return document.body.getAttribute("data-theme") === "dark";
  });

  React.useEffect(() => {
    const observer = new MutationObserver(() => {
      setIsDarkTheme(document.body.getAttribute("data-theme") === "dark");
    });
    observer.observe(document.body, { attributes: true, attributeFilter: ["data-theme"] });
    return () => observer.disconnect();
  }, []);

  const warningBgColor = isDarkTheme ? "#3C3C1F" : "#DAA520";
  const warningBorderColor = isDarkTheme ? "#B8860B" : "#B8860B";

  const fileInputRef = React.useRef<HTMLInputElement>(null);

  const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (file) {
      setSelectedFile(file);
      onLog(`File selected: ${file.name}`, "info");
    }
  };

  const handleImport = async () => {
    if (!selectedFile) return;
    console.log(window.toolboxAPI.connections.getActiveConnection());
    try {
      await props.lgSvc.importTranslations(selectedFile, parseInt(batchCount) || 10);
    } catch (err) {
      onLog(`Import error: ${(err as Error).message}`, "error");
    }
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100%", gap: "16px", padding: "16px" }}>
      {vm.exporting && (
        <div style={{ display: "flex", flexDirection: "column", gap: "12px", margin: "20px" }}>
          {/* Overall Progress */}
          <Field validationMessage={vm.message} validationState="none">
            <Label weight="semibold">Overall Progress</Label>
            <ProgressBar thickness="large" value={vm.exportpercentage}></ProgressBar>
          </Field>

          {/* Batch Progress */}
          <Field validationMessage={vm.batchMessage} validationState="none">
            <Label weight="semibold">Current Process Progress</Label>
            <ProgressBar thickness="medium" value={vm.batchProgress}></ProgressBar>
          </Field>
        </div>
      )}
      {!vm.exporting && (
        <>
          {/* Warning Messages */}
          <div style={{ display: "flex", flexDirection: "column", gap: "12px" }}>
            <div
              className="warningBox"
              // style={{
              //   display: "flex",
              //   gap: "12px",
              //   padding: "16px",
              //   backgroundColor: warningBgColor,
              //   border: `2px solid ${warningBorderColor}`,
              //   borderRadius: "4px",
              // }}
            >
              <div style={{ fontSize: "40px", minWidth: "40px", lineHeight: 1 }}>⚠️</div>
              <div>
                <div style={{ fontWeight: 600, marginBottom: "8px" }}>
                  The Excel file used to import translations must meet the following conditions:
                </div>
                <ul style={{ margin: "0", paddingLeft: "20px" }}>
                  <li>Worksheet names have not been changed</li>
                  <li>Cells with colored backgrounds have not been changed</li>
                  <li>No rows have been added to any worksheet in the file</li>
                  <li>No worksheets have been added to the file</li>
                  <li>The only added columns are for additional languages</li>
                </ul>
              </div>
            </div>

            <div
              style={{
                display: "flex",
                gap: "12px",
                padding: "16px",
                backgroundColor: warningBgColor,
                border: `2px solid ${warningBorderColor}`,
                borderRadius: "4px",
              }}
            >
              <div style={{ fontSize: "40px", minWidth: "40px", lineHeight: 1 }}>⚠️</div>
              <div>
                <Caption1 style={{ fontWeight: 600, marginBottom: "8px" }}>
                  Prior to using this tool ensure that you have a backup of the current customizations
                </Caption1>
                <Caption2>
                  If there is a problem the backup can be re-imported to restore original translations
                </Caption2>
              </div>
            </div>
          </div>

          {/* Import Settings */}
          <div style={{ backgroundColor: tokens.colorNeutralBackground2, padding: "16px", borderRadius: "4px" }}>
            <Label weight="semibold" style={{ marginBottom: "12px", display: "block" }}>
              Import Settings
            </Label>
            <Field label="Batch count">
              <Input
                type="number"
                value={batchCount}
                onChange={(e) => setBatchCount(e.target.value)}
                min="1"
                style={{ width: "80px" }}
              />
            </Field>
          </div>

          {/* File Upload */}
          <div style={{ display: "flex", flexDirection: "column", gap: "12px" }}>
            <Field label="Select Excel File">
              <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx,.xls"
                onChange={handleFileUpload}
                style={{ display: "none" }}
              />
              <div style={{ display: "flex", gap: "8px", alignItems: "center" }}>
                <Button
                  icon={<FolderOpenRegular />}
                  appearance="secondary"
                  onClick={() => fileInputRef.current?.click()}
                >
                  Choose File
                </Button>
                <Caption1 style={{ flex: 1 }}>{selectedFile ? selectedFile.name : "No file selected"}</Caption1>
                <Button
                  icon={<ArrowUploadFilled />}
                  appearance="primary"
                  onClick={handleImport}
                  disabled={!selectedFile}
                >
                  Import
                </Button>
              </div>
            </Field>
          </div>
        </>
      )}
    </div>
  );
});
