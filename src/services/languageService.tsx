import { ViewModel } from "../model/viewModel";
import { dvService } from "./dataverseService";
import { exportLanguageService } from "./exportLanguageService";
import { importLanguageService } from "./importLanguageService";

interface languageServiceProps {
  dvSvc: dvService;
  vm: ViewModel;
  onLog: (message: string, type?: "info" | "success" | "warning" | "error") => void;
}

export class languageService {
  private _export: exportLanguageService;
  private _import: importLanguageService;
  onLog: (message: string, type?: "info" | "success" | "warning" | "error") => void;

  constructor(props: languageServiceProps) {
    this._export = new exportLanguageService(props);
    this._import = new importLanguageService(props);
    this.onLog = props.onLog;
  }

  async exportTranslations(): Promise<void> {
    return this._export.exportTranslations();
  }

  async importTranslations(file: File, batchCount: number): Promise<void> {
    return this._import.importTranslations(file, batchCount);
  }
}
