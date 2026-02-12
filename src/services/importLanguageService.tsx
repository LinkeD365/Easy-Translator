import { ViewModel } from "../model/viewModel";
import { dvService } from "./dataverseService";
import ExcelJS from "exceljs";

export interface ImportLanguageServiceProps {
  dvSvc: dvService;
  vm: ViewModel;
  onLog: (message: string, type?: "info" | "success" | "warning" | "error") => void;
}

export class importLanguageService {
  private dvSvc: dvService;
  private vm: ViewModel;
  private totalSheets: number = 0;
  onLog: (message: string, type?: "info" | "success" | "warning" | "error") => void;

  constructor(props: ImportLanguageServiceProps) {
    this.dvSvc = props.dvSvc;
    this.vm = props.vm;
    this.onLog = props.onLog;
  }

  /**
   * Import translations from an Excel file, replicating the MsCrmTools.Translator import logic.
   * Iterates over worksheets and applies updates per sheet type via the Dataverse API.
   */
  async importTranslations(file: File, batchCount: number): Promise<void> {
    try {
      this.vm.exporting = true;
      this.vm.exportpercentage = 0;
      this.vm.batchProgress = 0;
      this.vm.message = "Reading Excel file...";
      this.vm.batchMessage = "";
      this.onLog("Starting import...", "info");

      const arrayBuffer = await file.arrayBuffer();
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(arrayBuffer);

      const sheetNames = workbook.worksheets.map((ws) => ws.name);
      this.totalSheets = sheetNames.length;
      let processedSheets = 0;

      // Track forms/dashboards/sitemaps that need content updates after all sheets are processed
      const formUpdates: Map<string, { formId: string; formXml: string }> = new Map();
      const dashboardUpdates: Map<string, { formId: string; formXml: string }> = new Map();
      const siteMapUpdates: Map<string, { siteMapId: string; sitemapXml: string }> = new Map();

      for (const sheet of workbook.worksheets) {
        try {
          const progress = processedSheets / this.totalSheets;
          this.vm.exportpercentage = progress;

          switch (sheet.name) {
            case "Entities":
              this.vm.message = "Importing entity translations...";
              this.onLog("Importing entity translations...", "info");
              await this.importEntities(sheet, batchCount);
              break;

            case "Attributes":
              this.vm.message = "Importing attribute translations...";
              this.onLog("Importing attribute translations...", "info");
              await this.importAttributes(sheet, batchCount);
              break;

            case "Relationships":
              this.vm.message = "Importing relationship translations...";
              this.onLog("Importing relationship translations...", "info");
              await this.importRelationships(sheet, batchCount, "OneToMany");
              break;

            case "RelationshipsNN":
              this.vm.message = "Importing NN relationship translations...";
              this.onLog("Importing NN relationship translations...", "info");
              await this.importRelationships(sheet, batchCount, "ManyToMany");
              break;

            case "Global OptionSets":
              this.vm.message = "Importing global option set translations...";
              this.onLog("Importing global option set translations...", "info");
              await this.importGlobalOptionSets(sheet, batchCount);
              break;

            case "Local OptionSets":
            case "OptionSets":
              this.vm.message = "Importing option set translations...";
              this.onLog("Importing option set translations...", "info");
              await this.importLocalOptionSets(sheet, batchCount);
              break;

            case "Booleans":
              this.vm.message = "Importing boolean translations...";
              this.onLog("Importing boolean translations...", "info");
              await this.importBooleans(sheet, batchCount);
              break;

            case "Views":
              this.vm.message = "Importing view translations...";
              this.onLog("Importing view translations...", "info");
              await this.importViews(sheet, batchCount);
              break;

            case "Charts":
              this.vm.message = "Importing chart translations...";
              this.onLog("Importing chart translations...", "info");
              await this.importCharts(sheet, batchCount);
              break;

            case "Forms":
              this.vm.message = "Importing form name translations...";
              this.onLog("Importing form name translations...", "info");
              await this.importFormNames(sheet, batchCount);
              break;

            case "Forms Tabs":
              this.vm.message = "Preparing form tab translations...";
              this.onLog("Preparing form tab translations...", "info");
              await this.prepareFormContent(sheet, formUpdates, "tab");
              break;

            case "Forms Sections":
              this.vm.message = "Preparing form section translations...";
              this.onLog("Preparing form section translations...", "info");
              await this.prepareFormContent(sheet, formUpdates, "section");
              break;

            case "Forms Fields":
              this.vm.message = "Preparing form field translations...";
              this.onLog("Preparing form field translations...", "info");
              await this.prepareFormContent(sheet, formUpdates, "field");
              break;

            case "Dashboards":
              this.vm.message = "Importing dashboard translations...";
              this.onLog("Importing dashboard translations...", "info");
              await this.importFormNames(sheet, batchCount);
              break;

            case "Dashboards Tabs":
              this.vm.message = "Preparing dashboard tab translations...";
              this.onLog("Preparing dashboard tab translations...", "info");
              await this.prepareFormContent(sheet, dashboardUpdates, "tab");
              break;

            case "Dashboards Sections":
              this.vm.message = "Preparing dashboard section translations...";
              this.onLog("Preparing dashboard section translations...", "info");
              await this.prepareFormContent(sheet, dashboardUpdates, "section");
              break;

            case "Dashboards Fields":
              this.vm.message = "Preparing dashboard field translations...";
              this.onLog("Preparing dashboard field translations...", "info");
              await this.prepareFormContent(sheet, dashboardUpdates, "field");
              break;

            case "SiteMap Areas":
              this.vm.message = "Preparing sitemap area translations...";
              this.onLog("Preparing sitemap area translations...", "info");
              await this.prepareSiteMapContent(sheet, siteMapUpdates, "Area");
              break;

            case "SiteMap Groups":
              this.vm.message = "Preparing sitemap group translations...";
              this.onLog("Preparing sitemap group translations...", "info");
              await this.prepareSiteMapContent(sheet, siteMapUpdates, "Group");
              break;

            case "SiteMap SubAreas":
              this.vm.message = "Preparing sitemap subarea translations...";
              this.onLog("Preparing sitemap subarea translations...", "info");
              await this.prepareSiteMapContent(sheet, siteMapUpdates, "SubArea");
              break;

            default:
              this.onLog(`Skipping unknown sheet: ${sheet.name}`, "warning");
              break;
          }
        } catch (err) {
          this.onLog(`Error processing sheet "${sheet.name}": ${(err as Error).message}`, "error");
        } finally {
          processedSheets++;
        }
      }

      // Apply accumulated form content updates
      if (formUpdates.size > 0) {
        this.vm.message = "Importing form content translations...";
        this.onLog(`Updating ${formUpdates.size} forms with content changes...`, "info");
        await this.applyFormContentUpdates(formUpdates, batchCount);
      }

      // Apply accumulated dashboard content updates
      if (dashboardUpdates.size > 0) {
        this.vm.message = "Importing dashboard content translations...";
        this.onLog(`Updating ${dashboardUpdates.size} dashboards with content changes...`, "info");
        await this.applyFormContentUpdates(dashboardUpdates, batchCount);
      }

      // Apply accumulated sitemap updates
      if (siteMapUpdates.size > 0) {
        this.vm.message = "Importing sitemap translations...";
        this.onLog(`Updating ${siteMapUpdates.size} sitemaps...`, "info");
        await this.applySiteMapUpdates(siteMapUpdates);
      }

      // Publish all customizations
      this.vm.message = "Publishing customizations...";
      this.vm.exportpercentage = 0.95;
      this.vm.batchProgress = 0;
      this.vm.batchMessage = "";
      this.onLog("Publishing customizations...", "info");
      await this.dvSvc.publishAllCustomizations();

      this.vm.message = "Import complete!";
      this.vm.exportpercentage = 1;
      this.vm.batchProgress = 1;
      this.vm.batchMessage = "Complete!";
      this.onLog("Import completed successfully!", "success");
    } catch (err) {
      this.onLog(`Import failed: ${(err as Error).message}`, "error");
    } finally {
      this.vm.exporting = false;
      this.vm.batchProgress = 0;
      this.vm.batchMessage = "";
    }
  }

  // ────────────────────────────────
  // Import helpers – one per sheet type
  // ────────────────────────────────

  /**
   * Helper to read the LCID codes from the header row starting at a given column index.
   */
  private getLanguageColumns(sheet: ExcelJS.Worksheet, startCol: number): number[] {
    const headerRow = sheet.getRow(1);
    const lcids: number[] = [];
    for (let col = startCol; col <= sheet.columnCount; col++) {
      const val = headerRow.getCell(col).value;
      if (val != null) {
        lcids.push(Number(val));
      }
    }
    return lcids;
  }

  /**
   * Helper to read label values from a row starting at a column.
   * Returns a map of LCID -> label string, only for non-empty cells.
   */
  private readLabels(row: ExcelJS.Row, startCol: number, lcids: number[]): Record<number, string> {
    const labels: Record<number, string> = {};
    for (let i = 0; i < lcids.length; i++) {
      const cell = row.getCell(startCol + i);
      if (cell.value != null && String(cell.value).trim() !== "") {
        labels[lcids[i]] = String(cell.value);
      }
    }
    return labels;
  }

  /**
   * Strip guid braces: "{GUID}" -> "GUID"
   */
  private stripBraces(val: string): string {
    return val?.replace(/[{}]/g, "") ?? "";
  }

  /**
   * Import Entities sheet
   * Columns: Entity Id | Entity Logical Name | Type | LCID1 | LCID2 | ...
   */
  private async importEntities(sheet: ExcelJS.Worksheet, batchCount: number): Promise<void> {
    const lcids = this.getLanguageColumns(sheet, 4); // cols 4+ are languages
    // Group rows by entity logical name
    const entityMap: Map<
      string,
      {
        displayName: Record<number, string>;
        displayCollectionName: Record<number, string>;
        description: Record<number, string>;
      }
    > = new Map();

    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // skip header
      const entityLogicalName = String(row.getCell(2).value ?? "").trim();
      const type = String(row.getCell(3).value ?? "").trim();
      if (!entityLogicalName || !type) return;

      if (!entityMap.has(entityLogicalName)) {
        entityMap.set(entityLogicalName, { displayName: {}, displayCollectionName: {}, description: {} });
      }
      const entry = entityMap.get(entityLogicalName)!;
      const labels = this.readLabels(row, 4, lcids);

      if (type === "DisplayName") entry.displayName = { ...entry.displayName, ...labels };
      else if (type === "DisplayCollectionName")
        entry.displayCollectionName = { ...entry.displayCollectionName, ...labels };
      else if (type === "Description") entry.description = { ...entry.description, ...labels };
    });

    let processed = 0;
    const total = entityMap.size;

    for (const [logicalName, data] of entityMap) {
      try {
        await this.dvSvc.updateEntityMetadata(
          logicalName,
          Object.keys(data.displayName).length > 0 ? data.displayName : null,
          Object.keys(data.displayCollectionName).length > 0 ? data.displayCollectionName : null,
          Object.keys(data.description).length > 0 ? data.description : null,
        );
        processed++;

        // Update batch progress
        this.vm.batchProgress = processed / total;
        this.vm.batchMessage = `Processing entity ${processed} of ${total}: ${logicalName}`;

        if (processed % batchCount === 0) {
          this.onLog(`Entities: ${processed}/${total} updated`, "info");
        }
      } catch (err) {
        this.onLog(`Failed to update entity ${logicalName}: ${(err as Error).message}`, "error");
      }
    }

    this.vm.batchProgress = 0;
    this.vm.batchMessage = "";
    this.onLog(`Entities: ${processed}/${total} updated`, "success");
  }

  /**
   * Import Attributes sheet
   * Columns: Attribute Id | Entity Logical Name | Attribute Logical Name | Type | LCID1 | ...
   */
  private async importAttributes(sheet: ExcelJS.Worksheet, batchCount: number): Promise<void> {
    const lcids = this.getLanguageColumns(sheet, 5); // cols 5+ are languages
    // Group by entity+attribute
    const attrMap: Map<
      string,
      {
        entityLogicalName: string;
        attributeLogicalName: string;
        displayName: Record<number, string>;
        description: Record<number, string>;
      }
    > = new Map();

    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      const entityLogicalName = String(row.getCell(2).value ?? "").trim();
      const attributeLogicalName = String(row.getCell(3).value ?? "").trim();
      const type = String(row.getCell(4).value ?? "").trim();
      if (!entityLogicalName || !attributeLogicalName || !type) return;

      const key = `${entityLogicalName}|${attributeLogicalName}`;
      if (!attrMap.has(key)) {
        attrMap.set(key, { entityLogicalName, attributeLogicalName, displayName: {}, description: {} });
      }
      const entry = attrMap.get(key)!;
      const labels = this.readLabels(row, 5, lcids);

      if (type === "DisplayName") entry.displayName = { ...entry.displayName, ...labels };
      else if (type === "Description") entry.description = { ...entry.description, ...labels };
    });

    let processed = 0;
    const total = attrMap.size;

    for (const [, data] of attrMap) {
      try {
        await this.dvSvc.updateAttributeMetadata(
          data.entityLogicalName,
          data.attributeLogicalName,
          Object.keys(data.displayName).length > 0 ? data.displayName : null,
          Object.keys(data.description).length > 0 ? data.description : null,
        );
        processed++;

        // Update batch progress
        this.vm.batchProgress = processed / total;
        this.vm.batchMessage = `Processing attribute ${processed} of ${total}: ${data.entityLogicalName}.${data.attributeLogicalName}`;

        if (processed % batchCount === 0) {
          this.vm.exportpercentage += processed / total / 100;
          this.vm.message = `Importing attributes... (${data.entityLogicalName}.${data.attributeLogicalName})`;
          this.onLog(`Attributes: ${processed}/${total} updated`, "info");
        }
      } catch (err) {
        this.onLog(
          `Failed to update attribute ${data.attributeLogicalName} on ${data.entityLogicalName}: ${(err as Error).message}`,
          "error",
        );
      }
    }

    this.vm.batchProgress = 0;
    this.vm.batchMessage = "";
    this.onLog(`Attributes: ${processed}/${total} updated`, "success");
  }

  /**
   * Import Relationships / RelationshipsNN sheets
   * Relationships columns: Entity | Relationship Id | Relationship Name | Relationship entity | LCID1 | ...
   * RelationshipsNN columns: Entity | Relationship Id | Relationship Intersect Entity | LCID1 | ...
   */
  private async importRelationships(
    sheet: ExcelJS.Worksheet,
    batchCount: number,
    relationType: "OneToMany" | "ManyToMany",
  ): Promise<void> {
    const startCol = relationType === "ManyToMany" ? 4 : 5;
    const lcids = this.getLanguageColumns(sheet, startCol);

    let processed = 0;
    const rows: { schemaName: string; labels: Record<number, string> }[] = [];
    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      // For OneToMany: col 3 is schema name; for ManyToMany col 3 is intersect entity
      const schemaName =
        relationType === "ManyToMany"
          ? String(row.getCell(3).value ?? "").trim()
          : String(row.getCell(3).value ?? "").trim();
      const labels = this.readLabels(row, startCol, lcids);
      if (schemaName && Object.keys(labels).length > 0) {
        rows.push({ schemaName, labels });
      }
    });

    const total = rows.length;
    const relType = relationType === "ManyToMany" ? "ManyToManyRelationship" : "OneToManyRelationship";

    for (const item of rows) {
      try {
        await this.dvSvc.updateRelationshipLabel(item.schemaName, item.labels, relType);
        processed++;

        // Update batch progress
        this.vm.batchProgress = processed / total;
        this.vm.batchMessage = `Processing relationship ${processed} of ${total}: ${item.schemaName}`;

        if (processed % batchCount === 0) {
          this.onLog(`Relationships: ${processed}/${total} updated`, "info");
        }
      } catch (err) {
        this.onLog(`Failed to update relationship ${item.schemaName}: ${(err as Error).message}`, "error");
      }
    }

    this.vm.batchProgress = 0;
    this.vm.batchMessage = "";
    this.onLog(`Relationships: ${processed}/${total} updated`, "success");
  }

  /**
   * Import Global OptionSets sheet
   * Columns: OptionSet Id | OptionSet Name | Value | Type | LCID1 | ...
   */
  private async importGlobalOptionSets(sheet: ExcelJS.Worksheet, batchCount: number): Promise<void> {
    const lcids = this.getLanguageColumns(sheet, 5); // cols 5+ are languages
    let processed = 0;
    const rows: { optionSetName: string; value: number; type: string; labels: Record<number, string> }[] = [];

    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      const optionSetName = String(row.getCell(2).value ?? "").trim();
      const value = Number(row.getCell(3).value ?? 0);
      const type = String(row.getCell(4).value ?? "Label").trim();
      const labels = this.readLabels(row, 5, lcids);
      if (optionSetName && Object.keys(labels).length > 0) {
        rows.push({ optionSetName, value, type, labels });
      }
    });

    const total = rows.length;

    for (const item of rows) {
      try {
        await this.dvSvc.updateOptionValue(
          null,
          null,
          item.optionSetName,
          item.value,
          item.labels,
          item.type === "Description",
        );
        processed++;

        // Update batch progress
        this.vm.batchProgress = processed / total;
        this.vm.batchMessage = `Processing global option set ${processed} of ${total}: ${item.optionSetName}`;

        if (processed % batchCount === 0) {
          this.onLog(`Global OptionSets: ${processed}/${total} updated`, "info");
        }
      } catch (err) {
        this.onLog(
          `Failed to update global optionset ${item.optionSetName} value ${item.value}: ${(err as Error).message}`,
          "error",
        );
      }
    }

    this.vm.batchProgress = 0;
    this.vm.batchMessage = "";
    this.onLog(`Global OptionSets: ${processed}/${total} updated`, "success");
  }

  /**
   * Import Local OptionSets sheet
   * Columns: Attribute Id | Entity Logical Name | Attribute Logical Name | Attribute Type | Value | Type | LCID1 | ...
   */
  private async importLocalOptionSets(sheet: ExcelJS.Worksheet, batchCount: number): Promise<void> {
    const lcids = this.getLanguageColumns(sheet, 7); // cols 7+ are languages
    let processed = 0;
    const rows: {
      entityLogicalName: string;
      attributeLogicalName: string;
      value: number;
      type: string;
      labels: Record<number, string>;
    }[] = [];

    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      const entityLogicalName = String(row.getCell(2).value ?? "").trim();
      const attributeLogicalName = String(row.getCell(3).value ?? "").trim();
      const value = Number(row.getCell(5).value ?? 0);
      const type = String(row.getCell(6).value ?? "Label").trim();
      const labels = this.readLabels(row, 7, lcids);
      if (entityLogicalName && attributeLogicalName && Object.keys(labels).length > 0) {
        rows.push({ entityLogicalName, attributeLogicalName, value, type, labels });
      }
    });

    const total = rows.length;

    for (const item of rows) {
      try {
        await this.dvSvc.updateOptionValue(
          item.entityLogicalName,
          item.attributeLogicalName,
          null,
          item.value,
          item.labels,
          item.type === "Description",
        );
        processed++;

        // Update batch progress
        this.vm.batchProgress = processed / total;
        this.vm.batchMessage = `Processing option set ${processed} of ${total}: ${item.entityLogicalName}.${item.attributeLogicalName}`;

        if (processed % batchCount === 0) {
          this.onLog(`OptionSets: ${processed}/${total} updated`, "info");
        }
      } catch (err) {
        this.onLog(
          `Failed to update optionset ${item.attributeLogicalName} value ${item.value}: ${(err as Error).message}`,
          "error",
        );
      }
    }

    this.vm.batchProgress = 0;
    this.vm.batchMessage = "";
    this.onLog(`OptionSets: ${processed}/${total} updated`, "success");
  }

  /**
   * Import Booleans sheet
   * Columns: Attribute Id | Entity Logical Name | Attribute Logical Name | Attribute Type | Value | Type | LCID1 | ...
   */
  private async importBooleans(sheet: ExcelJS.Worksheet, batchCount: number): Promise<void> {
    // Booleans have the same layout as local option sets
    await this.importLocalOptionSets(sheet, batchCount);
  }

  /**
   * Import Views sheet
   * Columns: View Id | Entity Logical Name | View Type | Type | LCID1 | ...
   */
  private async importViews(sheet: ExcelJS.Worksheet, batchCount: number): Promise<void> {
    const lcids = this.getLanguageColumns(sheet, 5); // cols 5+ are languages
    let processed = 0;
    const rows: { viewId: string; type: string; labels: Record<number, string> }[] = [];

    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      const viewId = this.stripBraces(String(row.getCell(1).value ?? ""));
      const type = String(row.getCell(4).value ?? "").trim();
      const labels = this.readLabels(row, 5, lcids);
      if (viewId && Object.keys(labels).length > 0) {
        rows.push({ viewId, type, labels });
      }
    });

    const total = rows.length;

    for (const item of rows) {
      try {
        const attrName = item.type === "Description" ? "description" : "name";
        await this.dvSvc.setLocLabels("savedqueries", item.viewId, attrName, item.labels);
        processed++;

        // Update batch progress
        this.vm.batchProgress = processed / total;
        this.vm.batchMessage = `Processing view ${processed} of ${total}: ${item.viewId}`;

        if (processed % batchCount === 0) {
          this.onLog(`Views: ${processed}/${total} updated`, "info");
        }
      } catch (err) {
        this.onLog(`Failed to update view ${item.viewId}: ${(err as Error).message}`, "error");
      }
    }

    this.vm.batchProgress = 0;
    this.vm.batchMessage = "";
    this.onLog(`Views: ${processed}/${total} updated`, "success");
  }

  /**
   * Import Charts sheet
   * Columns: Chart Id | Entity Logical Name | Type | LCID1 | ...
   */
  private async importCharts(sheet: ExcelJS.Worksheet, batchCount: number): Promise<void> {
    const lcids = this.getLanguageColumns(sheet, 4); // cols 4+ are languages
    let processed = 0;
    const rows: { chartId: string; type: string; labels: Record<number, string> }[] = [];

    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      const chartId = this.stripBraces(String(row.getCell(1).value ?? ""));
      const type = String(row.getCell(3).value ?? "").trim();
      const labels = this.readLabels(row, 4, lcids);
      if (chartId && Object.keys(labels).length > 0) {
        rows.push({ chartId, type, labels });
      }
    });

    const total = rows.length;

    for (const item of rows) {
      try {
        const attrName = item.type === "Description" ? "description" : "name";
        await this.dvSvc.setLocLabels("savedqueryvisualizations", item.chartId, attrName, item.labels);
        processed++;

        // Update batch progress
        this.vm.batchProgress = processed / total;
        this.vm.batchMessage = `Processing chart ${processed} of ${total}: ${item.chartId}`;

        if (processed % batchCount === 0) {
          this.onLog(`Charts: ${processed}/${total} updated`, "info");
        }
      } catch (err) {
        this.onLog(`Failed to update chart ${item.chartId}: ${(err as Error).message}`, "error");
      }
    }

    this.vm.batchProgress = 0;
    this.vm.batchMessage = "";
    this.onLog(`Charts: ${processed}/${total} updated`, "success");
  }

  /**
   * Import form/dashboard name translations (Forms / Dashboards sheets)
   * Forms columns: Form Unique Id | Form Id | Entity Logical Name | Form Type | Type | LCID1 | ...
   * Dashboards columns: Form Unique Id | Form Id | Type | LCID1 | ...
   */
  private async importFormNames(sheet: ExcelJS.Worksheet, batchCount: number): Promise<void> {
    // Detect layout: Forms has 5 header cols before LCIDs, Dashboards has 3
    const headerRow = sheet.getRow(1);
    let langStartCol = 4; // default for Dashboards
    for (let col = 1; col <= sheet.columnCount; col++) {
      const val = String(headerRow.getCell(col).value ?? "");
      const num = Number(val);
      if (!isNaN(num) && num > 1000) {
        langStartCol = col;
        break;
      }
    }

    const lcids = this.getLanguageColumns(sheet, langStartCol);
    let processed = 0;
    const rows: { formId: string; type: string; labels: Record<number, string> }[] = [];

    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      const formId = this.stripBraces(String(row.getCell(2).value ?? ""));
      // Type column is the one right before the LCIDs
      const type = String(row.getCell(langStartCol - 1).value ?? "").trim();
      const labels = this.readLabels(row, langStartCol, lcids);
      if (formId && Object.keys(labels).length > 0) {
        rows.push({ formId, type, labels });
      }
    });

    const total = rows.length;

    for (const item of rows) {
      try {
        const attrName = item.type === "description" || item.type === "Description" ? "description" : "name";
        await this.dvSvc.setLocLabels("systemforms", item.formId, attrName, item.labels);
        processed++;

        // Update batch progress
        this.vm.batchProgress = processed / total;
        this.vm.batchMessage = `Processing form ${processed} of ${total}: ${item.formId}`;

        if (processed % batchCount === 0) {
          this.onLog(`Form names: ${processed}/${total} updated`, "info");
        }
      } catch (err) {
        this.onLog(`Failed to update form name ${item.formId}: ${(err as Error).message}`, "error");
      }
    }

    this.vm.batchProgress = 0;
    this.vm.batchMessage = "";
    this.onLog(`Form names: ${processed}/${total} updated`, "success");
  }

  /**
   * Prepare form content (tabs, sections, fields) for batch update.
   * Reads from the sheet and accumulates XML changes into the formUpdates map.
   *
   * Forms Tabs columns:    Tab Id | Entity Logical Name | Form Name | Form Unique Id | Form Id | LCID1 | ...
   * Forms Sections columns: Section Id | Entity Logical Name | Form Name | Form Unique Id | Form Id | Tab Name | LCID1 | ...
   * Forms Fields columns:  Label Id | Entity Logical Name | Form Name | Form Unique Id | Form Id | Tab Name | Section Name | Attribute | Display Name | LCID1 | ...
   */
  private async prepareFormContent(
    sheet: ExcelJS.Worksheet,
    formUpdates: Map<string, { formId: string; formXml: string }>,
    contentType: "tab" | "section" | "field",
  ): Promise<void> {
    // Detect language start column
    const headerRow = sheet.getRow(1);
    let langStartCol = 6; // default for tabs
    for (let col = 1; col <= sheet.columnCount; col++) {
      const val = String(headerRow.getCell(col).value ?? "");
      const num = Number(val);
      if (!isNaN(num) && num > 1000) {
        langStartCol = col;
        break;
      }
    }
    const lcids = this.getLanguageColumns(sheet, langStartCol);

    // Determine Form Id column position (varies by content type)
    let formIdCol: number;
    if (contentType === "tab") formIdCol = 5;
    else if (contentType === "section") formIdCol = 5;
    else formIdCol = 5; // field
    let totalRows = 0;
    let processedRows = 0;

    // Count total rows first
    totalRows = sheet.rowCount - 1; // exclude header
    // sheet.eachRow((row, rowNumber) => {
    //   if (rowNumber > 1) totalRows++;
    // });

    for (let i = 2; i <= sheet.rowCount; i++) {
      const row = sheet.getRow(i);

      const elementId = String(row.getCell(1).value ?? "").trim();
      const formId = this.stripBraces(String(row.getCell(formIdCol).value ?? ""));
      if (!formId || !elementId) return;

      const labels = this.readLabels(row, langStartCol, lcids);
      if (Object.keys(labels).length === 0) return;

      // Fetch or reuse form XML
      if (!formUpdates.has(formId)) {
        try {
          const form = await this.dvSvc.getFormById(formId);
          formUpdates.set(formId, {
            formId,
            formXml: (form.formxml as string) ?? "",
          });
        } catch (err) {
          this.onLog(`Could not retrieve form ${formId}: ${(err as Error).message}`, "error");
          return;
        }
      }

      const formData = formUpdates.get(formId)!;
      const xmlDoc = new DOMParser().parseFromString(formData.formXml, "application/xml");

      // Find the element by its id attribute
      let element: Element | null = null;
      const tagName = contentType === "tab" ? "tab" : contentType === "section" ? "section" : "cell";
      const elements = xmlDoc.getElementsByTagName(tagName);
      for (let i = 0; i < elements.length; i++) {
        if (elements[i].getAttribute("id") === elementId || elements[i].getAttribute("name") === elementId) {
          element = elements[i];
          break;
        }
      }

      if (!element) {
        this.onLog(`Could not find ${contentType} with id ${elementId} in form ${formId}`, "warning");
        return;
      }

      // Update or create labels element
      let labelsEl = element.querySelector("labels");
      if (!labelsEl) {
        labelsEl = xmlDoc.createElement("labels");
        element.insertBefore(labelsEl, element.firstChild);
      }

      // Clear existing labels and re-add
      while (labelsEl.firstChild) {
        labelsEl.removeChild(labelsEl.firstChild);
      }

      for (const [lcid, labelText] of Object.entries(labels)) {
        const labelEl = xmlDoc.createElement("label");
        labelEl.setAttribute("description", labelText);
        labelEl.setAttribute("languagecode", String(lcid));
        labelsEl.appendChild(labelEl);
      }

      // Serialize back
      const serializer = new XMLSerializer();
      formData.formXml = serializer.serializeToString(xmlDoc);

      // Update batch progress
      processedRows++;
      this.vm.batchProgress = processedRows / totalRows;
      this.vm.batchMessage = `Preparing ${contentType} ${processedRows} of ${totalRows}: ${elementId}`;
    }

    this.vm.batchProgress = 0;
    this.vm.batchMessage = "";
  }

  /**
   * Apply accumulated form content updates to Dataverse
   */
  private async applyFormContentUpdates(
    formUpdates: Map<string, { formId: string; formXml: string }>,
    batchCount: number,
  ): Promise<void> {
    let processed = 0;
    const total = formUpdates.size;

    for (const [, data] of formUpdates) {
      try {
        await this.dvSvc.updateForm(data.formId, { formxml: data.formXml });
        processed++;

        // Update batch progress
        this.vm.batchProgress = processed / total;
        this.vm.batchMessage = `Updating form ${processed} of ${total}: ${data.formId}`;

        if (processed % batchCount === 0) {
          this.onLog(`Forms content: ${processed}/${total} updated`, "info");
        }
      } catch (err) {
        this.onLog(`Failed to update form ${data.formId}: ${(err as Error).message}`, "error");
      }
    }

    this.vm.batchProgress = 0;
    this.vm.batchMessage = "";
    this.onLog(`Forms content: ${processed}/${total} updated`, "success");
  }

  /**
   * Prepare sitemap content for update.
   * SiteMap Areas columns:    SiteMap Name | SiteMap Id | Area Id | Type | LCID1 | ...
   * SiteMap Groups columns:   SiteMap Name | SiteMap Id | Area Id | Group Id | Type | LCID1 | ...
   * SiteMap SubAreas columns: SiteMap Name | SiteMap Id | Area Id | Group Id | SubArea Id | Type | LCID1 | ...
   */
  private async prepareSiteMapContent(
    sheet: ExcelJS.Worksheet,
    siteMapUpdates: Map<string, { siteMapId: string; sitemapXml: string }>,
    elementType: "Area" | "Group" | "SubArea",
  ): Promise<void> {
    // Detect language start column
    const headerRow = sheet.getRow(1);
    let langStartCol = 5;
    for (let col = 1; col <= sheet.columnCount; col++) {
      const val = String(headerRow.getCell(col).value ?? "");
      const num = Number(val);
      if (!isNaN(num) && num > 1000) {
        langStartCol = col;
        break;
      }
    }
    const lcids = this.getLanguageColumns(sheet, langStartCol);

    // Determine element ID column and SiteMap Id column
    const siteMapIdCol = 2;
    let elementIdCol: number;
    let typeCol: number;
    if (elementType === "Area") {
      elementIdCol = 3;
      typeCol = 4;
    } else if (elementType === "Group") {
      elementIdCol = 4;
      typeCol = 5;
    } else {
      elementIdCol = 5;
      typeCol = 6;
    }

    let totalRows = 0;
    let processedRows = 0;

    // Count total rows first
    sheet.eachRow((_, rowNumber) => {
      if (rowNumber > 1) totalRows++;
    });

    sheet.eachRow(async (row, rowNumber) => {
      if (rowNumber === 1) return;

      const siteMapId = String(row.getCell(siteMapIdCol).value ?? "").trim();
      const elementId = String(row.getCell(elementIdCol).value ?? "").trim();
      const type = String(row.getCell(typeCol).value ?? "").trim();
      if (!siteMapId || !elementId) return;

      const labels = this.readLabels(row, langStartCol, lcids);
      if (Object.keys(labels).length === 0) return;

      // Fetch or reuse sitemap XML
      if (!siteMapUpdates.has(siteMapId)) {
        try {
          const sm = await this.dvSvc.getSiteMapById(siteMapId);
          siteMapUpdates.set(siteMapId, {
            siteMapId,
            sitemapXml: (sm.sitemapxml as string) ?? "",
          });
        } catch (err) {
          this.onLog(`Could not retrieve sitemap ${siteMapId}: ${(err as Error).message}`, "error");
          return;
        }
      }

      const smData = siteMapUpdates.get(siteMapId)!;
      const xmlDoc = new DOMParser().parseFromString(smData.sitemapXml, "application/xml");

      // Find element by Id attribute
      const elements = xmlDoc.getElementsByTagName(elementType);
      let targetElement: Element | null = null;
      for (let i = 0; i < elements.length; i++) {
        if (elements[i].getAttribute("Id") === elementId) {
          targetElement = elements[i];
          break;
        }
      }

      if (!targetElement) {
        this.onLog(`Could not find ${elementType} with Id ${elementId} in sitemap ${siteMapId}`, "warning");
        return;
      }

      // Determine which child element to update (Titles or Descriptions)
      const containerName = type === "Description" ? "Descriptions" : "Titles";
      const labelAttr = type === "Description" ? "Description" : "Title";

      let container = targetElement.querySelector(`:scope > ${containerName}`);
      if (!container) {
        container = xmlDoc.createElement(containerName);
        targetElement.appendChild(container);
      }

      // Remove existing labels and add updated ones
      while (container.firstChild) {
        container.removeChild(container.firstChild);
      }

      for (const [lcid, labelText] of Object.entries(labels)) {
        const titleEl = xmlDoc.createElement(containerName === "Titles" ? "Title" : "Description");
        titleEl.setAttribute("LCID", String(lcid));
        titleEl.setAttribute(labelAttr, labelText);
        container.appendChild(titleEl);
      }

      const serializer = new XMLSerializer();
      smData.sitemapXml = serializer.serializeToString(xmlDoc);

      // Update batch progress
      processedRows++;
      this.vm.batchProgress = processedRows / totalRows;
      this.vm.batchMessage = `Preparing sitemap ${elementType} ${processedRows} of ${totalRows}: ${elementId}`;
    });

    this.vm.batchProgress = 0;
    this.vm.batchMessage = "";
  }

  /**
   * Apply accumulated sitemap updates
   */
  private async applySiteMapUpdates(
    siteMapUpdates: Map<string, { siteMapId: string; sitemapXml: string }>,
  ): Promise<void> {
    let processed = 0;
    const total = siteMapUpdates.size;

    for (const [, data] of siteMapUpdates) {
      try {
        await this.dvSvc.updateSiteMap(data.siteMapId, data.sitemapXml);
        processed++;

        // Update batch progress
        this.vm.batchProgress = processed / total;
        this.vm.batchMessage = `Updating sitemap ${processed} of ${total}: ${data.siteMapId}`;
      } catch (err) {
        this.onLog(`Failed to update sitemap ${data.siteMapId}: ${(err as Error).message}`, "error");
      }
    }

    this.vm.batchProgress = 0;
    this.vm.batchMessage = "";
    this.onLog(`SiteMaps: ${processed}/${total} updated`, "success");
  }
}
