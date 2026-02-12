import React, { useCallback, useEffect, useMemo, useState } from "react";
import { FluentProvider, webLightTheme, webDarkTheme } from "@fluentui/react-components";

import { useConnection, useEventLog, useToolboxEvents } from "./hooks/useToolboxAPI";
import { dvService } from "./services/dataverseService";
import { ViewModel } from "./model/viewModel";
import { EasyTranslator } from "./components/EasyTranslator";
import { languageService } from "./services/languageService";



function App() {
  const { connection, refreshConnection } = useConnection();
  const { addLog } = useEventLog();
  const [theme, setTheme] = React.useState<"light" | "dark">("light");

  // Handle platform events
  const handleEvent = useCallback(
    (event: string, _data: any) => {
      switch (event) {
        case "connection:updated":
        case "connection:created":
          refreshConnection();
          break;

        case "connection:deleted":
          refreshConnection();
          break;

        case "terminal:output":
        case "terminal:command:completed":
        case "terminal:error":
          // Terminal events handled by dedicated components
          break;
        case "settings:updated":
          if (_data && _data.theme) {
            document.body.setAttribute("data-theme", _data.theme);
            document.body.setAttribute("data-ag-theme-mode", _data.theme);
            setTheme(_data.theme);
          }
            break;
        default:
          addLog(`Unhandled event: ${event}`, "warning");
          break;
      }
    },
    [refreshConnection, addLog],
  );

  useToolboxEvents(handleEvent);

  // Add initial log (run only once on mount)
  useEffect(() => {
    (async () => {
      const currentTheme = await window.toolboxAPI.utils.getCurrentTheme();
      document.body.setAttribute("data-theme", currentTheme);
      document.body.setAttribute("data-ag-theme-mode", currentTheme);
      setTheme(currentTheme);
    })();
    addLog("FlowFinder initialized", "success");
  }, [addLog]);

  // Get theme from Toolbox API
  useEffect(() => {
    const getTheme = async () => {
      try {
        const currentTheme = await window.toolboxAPI.utils.getCurrentTheme();
        setTheme(currentTheme === "dark" ? "dark" : "light");
      } catch (error) {
        console.error("Error getting theme:", error);
      }
    };
    getTheme();
  }, []);
  const [vm] = useState(() => new ViewModel());
  const dvSvc = useMemo(() => {
    if (!connection) return null;
    return new dvService({
      connection: connection,
      dvApi: window.dataverseAPI,
      onLog: addLog,
    });
  }, [connection, addLog]);

  const languageSvc = useMemo(() => {
    if (!dvSvc) return null;
    return new languageService({
      dvSvc: dvSvc,
      vm: vm,
      onLog: addLog,
    });
  }, [dvSvc, vm, addLog]);

  return (
    <FluentProvider theme={theme === "dark" ? webDarkTheme : webLightTheme}>
      <EasyTranslator dvSvc={dvSvc!} vm={vm} lgSvc={languageSvc!} onLog={addLog} />
    </FluentProvider>
  );
}

export default App;
