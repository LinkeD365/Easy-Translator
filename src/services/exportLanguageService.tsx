import { LabelOptions, LanguageDef, ViewModel } from "../model/viewModel";
import { dvService } from "./dataverseService";
import ExcelJS from "exceljs";

export interface ExportLanguageServiceProps {
  dvSvc: dvService;
  vm: ViewModel;
  onLog: (message: string, type?: "info" | "success" | "warning" | "error") => void;
}

export class exportLanguageService {
  private dvSvc: dvService;
  private vm: ViewModel;
  private outputLangs: LanguageDef[];
  private baseLanguage: LanguageDef | undefined;
  onLog: (message: string, type?: "info" | "success" | "warning" | "error") => void;

  constructor(props: ExportLanguageServiceProps) {
    this.dvSvc = props.dvSvc;
    this.vm = props.vm;
    this.onLog = props.onLog;
    this.outputLangs = [];
  }

  async exportTranslations(): Promise<void> {
    try {
      this.onLog("Exporting translations...", "info");
      this.vm.exporting = true;
      this.vm.message = "Getting base language...";
      this.vm.exportpercentage = 0.05;
      this.baseLanguage = await this.dvSvc.getBaseLanguage();
      this.onLog(`Base language is ${this.baseLanguage.name} (${this.baseLanguage.code})`, "info");

      if (this.vm.options.exportAllLanguages) this.outputLangs = this.vm.allLanguages as LanguageDef[];
      else this.outputLangs = [this.vm.selectedLanguage as LanguageDef];

      if (!this.outputLangs.some((ld) => ld.code === this.baseLanguage?.code))
        this.outputLangs.push(this.baseLanguage as LanguageDef);

      const workbook = new ExcelJS.Workbook();

      this.vm.message = "Exporting table info...";
      this.vm.exportpercentage = 0.1;
      await this.exportTableInfo(workbook);

      this.vm.message = "Exporting attributes...";
      this.vm.exportpercentage = 0.15;
      await this.exportAttributes(workbook);

      this.vm.message = "Exporting relationships...";
      this.vm.exportpercentage = 0.2;
      await this.exportRelationships(workbook);

      this.vm.message = "Exporting option sets...";
      this.vm.exportpercentage = 0.3;
      if (this.vm.options.localOptionSets || this.vm.options.globalOptionSets) {
        await this.exportOptionSets(workbook);
      }
      this.vm.message = "Exporting boolean options...";
      this.vm.exportpercentage = 0.4;
      await this.exportBooleans(workbook);

      this.vm.message = "Exporting views...";
      this.vm.exportpercentage = 0.5;
      await this.exportViews(workbook);

      this.vm.message = "Exporting charts...";
      this.vm.exportpercentage = 0.6;
      await this.exportCharts(workbook);

      this.vm.message = "Exporting forms...";
      this.vm.exportpercentage = 0.7;
      await this.exportForms(workbook);

      this.vm.message = "Exporting site map...";
      this.vm.exportpercentage = 0.8;
      await this.exportSiteMap(workbook);

      this.vm.message = "Exporting dashboards...";
      this.vm.exportpercentage = 0.9;
      await this.exportDashboards(workbook);

      this.vm.message = "Finalizing export...";
      this.vm.exportpercentage = 0.95;

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      this.downloadBlob(blob, "translations.xlsx");

      this.vm.exporting = false;
      return;
    } catch (err) {
      this.onLog("Export failed.", "error");
      return Promise.reject(err);
    }
  }

  private async exportTableInfo(workbook: ExcelJS.Workbook) {
    const wsheet = workbook.addWorksheet("Entities");

    wsheet.addRow(["Entity Id", "Entity Logical Name", "Type", ...this.outputLangs.map((lang) => lang.code)]);

    await Promise.all(
      this.vm.selectedTables.map(async (table) => {
        await this.dvSvc.getTableMeta(table);
      }),
    );

    this.vm.selectedTables.forEach((table) => {
      const langProps =
        this.vm.options.labelOptions == LabelOptions.both
          ? table.langProps
          : this.vm.options.labelOptions == LabelOptions.names
            ? table.langProps.filter((lp) => lp.name === "DisplayCollectionName" || lp.name === "DisplayName")
            : table.langProps.filter((lp) => lp.name === "Description");
      langProps.forEach((langProp) => {
        const row = [
          `{${table.id}}`,
          table.logicalName,
          langProp.name,
          ...langProp.translation.map((trans) => trans.translation),
        ];

        wsheet.addRow(row);
      });
    });

    this.styleSheet(wsheet);
  }

  private async exportAttributes(workbook: ExcelJS.Workbook) {
    if (!this.vm.options.fields && !this.vm.options.formFields) return;
    const wsheet = workbook.addWorksheet("Attributes");
    wsheet.addRow([
      "Attribute Id",
      "Entity Logical Name",
      "Attribute Logical Name",
      "Type",
      ...this.outputLangs.map((lang) => lang.code),
    ]);

    await Promise.all(
      this.vm.selectedTables.map(async (table) => {
        this.onLog(`Fetching attributes for ${table.logicalName}...`, "info");
        await this.dvSvc.getTableFields(table);
      }),
    );
    this.vm.selectedTables.forEach((table) => {
      table.fields.forEach((fld) => {
        const langProps =
          this.vm.options.labelOptions == LabelOptions.both
            ? fld.langProps
            : fld.langProps.filter(
                (lp) =>
                  lp.name === (this.vm.options.labelOptions == LabelOptions.names ? "DisplayName" : "Description"),
              );
        langProps.forEach((lgProp) => {
          const row = [
            `{${fld.id}}`,
            table.logicalName,
            fld.name,
            lgProp.name,
            ...lgProp.translation.map((trans) => trans.translation),
          ];

          wsheet.addRow(row);
        });
      });
    });
  }

  private async exportRelationships(workbook: ExcelJS.Workbook) {
    return new Promise<void>(async (resolve, reject) => {
      try {
        if (!this.vm.options.relationships || this.vm.options.labelOptions == LabelOptions.descriptions) return;
        let wsheet = workbook.addWorksheet("Relationships");
        wsheet.addRow([
          "Entity",
          "Relationship Id",
          "Relationship Name",
          "Relationship entity",
          ...this.outputLangs.map((lang) => lang.code),
        ]);
        await Promise.all(
          this.vm.selectedTables.map(async (table) => {
            await this.dvSvc.getRelationships(table);
          }),
        );

        this.vm.selectedTables.forEach((table) => {
          table.relationships
            .filter((rel) => rel.type !== "ManyToManyRelationship")
            .forEach((rel) => {
              rel.langProps.forEach((lgProp) => {
                const row = [
                  table.logicalName,
                  `{${rel.id}}`,
                  ...Object.values(rel.props ?? {}).map((prop) => prop),
                  ...lgProp.translation.map((trans) => trans.translation),
                ];

                wsheet.addRow(row);
              });
            });
        });

        this.styleSheet(wsheet);

        wsheet = workbook.addWorksheet("RelationshipsNN");
        wsheet.addRow([
          "Entity",
          "Relationship Id",
          "Relationship Intersect Entity",
          ...this.outputLangs.map((lang) => lang.code),
        ]);

        this.vm.selectedTables.forEach((table) => {
          table.relationships
            .filter((rel) => rel.type === "ManyToManyRelationship")
            .forEach((rel) => {
              rel.langProps.forEach((lgProp) => {
                const row = [
                  table.logicalName,
                  `{${rel.id}}`,
                  rel.props?.IntersectEntityName,
                  ...lgProp.translation.map((trans) => trans.translation),
                ];

                wsheet.addRow(row);
              });
            });
        });
        this.onLog("Relationships loaded", "success");
        this.styleSheet(wsheet);
        resolve();
      } catch (err) {
        reject(err);
      }
    });
  }

  private async exportBooleans(workbook: ExcelJS.Workbook) {
    const wsheet = workbook.addWorksheet("Booleans");
    wsheet.addRow([
      "Attribute Id",
      "Entity Logical Name",
      "Attribute Logical Name",
      "Attribute Type",
      "Value",
      "Type",
      ...this.outputLangs.map((lang) => lang.code),
    ]);

    await Promise.all(
      this.vm.selectedTables.map(async (table) => {
        this.onLog(`Fetching boolean options for ${table.logicalName}...`, "info");
        await this.dvSvc.getBooleans(table);
      }),
    );
    this.vm.selectedTables.forEach((table) => {
      table.optionSets
        .filter((os) => os.type == "Boolean")
        .forEach((optionSet) => {
          optionSet.langProps
            .filter(
              (lp) =>
                this.vm.options.labelOptions === LabelOptions.both ||
                (this.vm.options.labelOptions === LabelOptions.names && lp.name === "Label") ||
                (this.vm.options.labelOptions === LabelOptions.descriptions && lp.name === "Description"),
            )
            .forEach((langProp) => {
              const row = [
                `{${optionSet.id}}`,
                table.logicalName,
                optionSet.attributeLogicalName,
                optionSet.type,
                optionSet.optionValue,
                langProp.name,
                ...this.outputLangs.map((lang) => {
                  const trans = langProp.translation.find((t) => t.code === lang.code);
                  return trans ? trans.translation : "";
                }),
              ];
              wsheet.addRow(row);
            });
        });
    });

    this.styleSheet(wsheet);
  }

  private async exportOptionSets(workbook: ExcelJS.Workbook) {
    let wsheet = workbook.addWorksheet("Local OptionSets");
    wsheet.addRow([
      "Attribute Id",
      "Entity Logical Name",
      "Attribute Logical Name",
      "Attribute Type",
      "Value",
      "Type",
      ...this.outputLangs.map((lang) => lang.code),
    ]);

    await Promise.all(
      this.vm.selectedTables.map(async (table) => {
        await this.dvSvc.getOptionSets(table);
      }),
    );

    if (this.vm.options.localOptionSets) {
      this.vm.selectedTables.forEach((table) => {
        table.optionSets
          .filter((os) => !os.isGlobal)
          .forEach((optionSet) => {
            optionSet.langProps
              .filter(
                (lp) =>
                  this.vm.options.labelOptions === LabelOptions.both ||
                  (this.vm.options.labelOptions === LabelOptions.names && lp.name === "Label") ||
                  (this.vm.options.labelOptions === LabelOptions.descriptions && lp.name === "Description"),
              )
              .forEach((langProp) => {
                const row = [
                  `{${optionSet.id}}`,
                  table.logicalName,
                  optionSet.attributeLogicalName,
                  optionSet.type,
                  optionSet.optionValue,
                  langProp.name,
                  ...this.outputLangs.map((lang) => {
                    const trans = langProp.translation.find((t) => t.code === lang.code);
                    return trans ? trans.translation : "";
                  }),
                ];
                wsheet.addRow(row);
              });
          });
      });
    }

    this.styleSheet(wsheet);

    wsheet = workbook.addWorksheet("Global OptionSets");
    wsheet.addRow(["OptionSet Id", "OptionSet Name", "Value", "Type", ...this.outputLangs.map((lang) => lang.code)]);

    this.vm.selectedTables.forEach((table) => {
      table.optionSets
        .filter((os) => os.isGlobal && os.type != "Boolean")
        .forEach((optionSet) => {
          optionSet.langProps
            .filter(
              (lp) =>
                this.vm.options.labelOptions === LabelOptions.both ||
                (this.vm.options.labelOptions === LabelOptions.names && lp.name === "Label") ||
                (this.vm.options.labelOptions === LabelOptions.descriptions && lp.name === "Description"),
            )
            .forEach((langProp) => {
              const row = [
                `{${optionSet.id}}`,
                table.logicalName,
                optionSet.attributeLogicalName,
                optionSet.type,
                optionSet.optionValue,
                langProp.name,
                ...this.outputLangs.map((lang) => {
                  const trans = langProp.translation.find((t) => t.code === lang.code);
                  return trans ? trans.translation : "";
                }),
              ];
              wsheet.addRow(row);
            });
        });
    });

    this.styleSheet(wsheet);
  }

  private async exportViews(workbook: ExcelJS.Workbook) {
    this.onLog("Exporting views...", "info");
    const wsheet = workbook.addWorksheet("Views");
    wsheet.addRow([
      "View Id",
      "Entity Logical Name",
      "View Type",
      "Type",
      ...this.outputLangs.map((lang) => lang.code),
    ]);

    await Promise.all(
      this.vm.selectedTables.map(async (table) => {
        await this.dvSvc.getViews(table);
      }),
    );

    this.vm.selectedTables.forEach((table) => {
      table.views.forEach((view) => {
        view.langProps.forEach((langProp) => {
          const row = [
            `{${view.id}}`,
            table.logicalName,
            view.type,
            langProp.name,
            ...langProp.translation.map((trans) => trans.translation),
          ];
          wsheet.addRow(row);
        });
      });
    });
  }

  private async exportCharts(workbook: ExcelJS.Workbook) {
    this.onLog("Exporting charts...", "info");
    const wsheet = workbook.addWorksheet("Charts");
    wsheet.addRow(["Chart Id", "Entity Logical Name", "Type", ...this.outputLangs.map((lang) => lang.code)]);
    await Promise.all(
      this.vm.selectedTables.map(async (table) => {
        await this.dvSvc.getCharts(table);
      }),
    );

    this.vm.selectedTables.forEach((table) => {
      table.charts.forEach((chart) => {
        chart.langProps.forEach((langProp) => {
          const row = [
            `{${chart.id}}`,
            table.logicalName,
            langProp.name,
            ...langProp.translation.map((trans) => trans.translation),
          ];
          wsheet.addRow(row);
        });
      });
    });
  }

  private async exportForms(workbook: ExcelJS.Workbook) {
    let wsheet = workbook.addWorksheet("Forms");
    wsheet.addRow([
      "Form Unique Id",
      "Form Id",
      "Entity Logical Name",
      "Form Type",
      "Type",
      ...this.outputLangs.map((lang) => lang.code),
    ]);
    const formsExport =
      this.vm.options.forms || this.vm.options.formFields || this.vm.options.formSections || this.vm.options.formTabs;
    if (formsExport || this.vm.options.dashboards) {
      await this.dvSvc.getUserLanguage().then((result: { uiLocale: string; locale: string; userid: string }) => {
        this.onLog(`User language is ${result.uiLocale} (${result.locale})`, "info");
        this.vm.uiLocale = result.uiLocale;
        this.vm.locale = result.locale;
        this.vm.userId = result.userid;
      });
      let currentLang = this.vm.uiLocale;
      for (const lang of this.outputLangs) {
        if (currentLang !== lang.code) {
          this.vm.message = `Updating language to ${lang.name}...`;
          await this.dvSvc.updateLanguage(lang.code, this.vm.userId);
          currentLang = lang.code;
        }

        if (formsExport) {
          this.vm.message = `Fetching forms for language ${lang.name}...`;
          await Promise.all(
            this.vm.selectedTables.map(async (table) => {
              await this.dvSvc.getForms(table, lang, lang.code === this.baseLanguage?.code ? true : false);
            }),
          );
        }
        if (this.vm.options.dashboards) {
          this.vm.message = `Fetching dashboards for language ${lang.name}...`;
          const dashboards = await this.dvSvc.getDashboards(this.vm.solution?.solutionId ?? "");
          if (lang.code === this.baseLanguage?.code) {
            dashboards.forEach((d) => {
              if (d.props) d.props.base = true;
              else d.props = { base: true };
            });
          }
          this.vm.dashboards.push(...dashboards);
        }
      }
    }
    this.vm.message = "Reverting language...";
    await this.dvSvc.updateLanguage(this.vm.uiLocale, this.vm.userId);
    if (this.vm.options.forms)
      this.vm.selectedTables.forEach((table) => {
        table.forms
          .filter((form) => form.props?.base)
          .forEach((form) => {
            form.langProps.forEach((langProp) => {
              const row = [
                `{${form.props?.uniqueName ?? ""}}`,
                `{${form.id}}`,
                table.logicalName,
                form.type,
                langProp.name,
                ...langProp.translation.map((trans) => trans.translation),
              ];
              wsheet.addRow(row);
            });
          });
      });
    this.styleSheet(wsheet);
    wsheet = workbook.addWorksheet("Forms Tabs");
    wsheet.addRow([
      "Tab Id",
      "Entity Logical Name",
      "Form Name",
      "Form Unique Id",
      "Form Id",
      ...this.outputLangs.map((lang) => lang.code),
    ]);
    if (this.vm.options.formTabs) {
      this.vm.selectedTables.forEach((table) => {
        table.forms.forEach((form) => {
          if (form.props) {
            const xmlDoc = new DOMParser().parseFromString(form.props.formXml, "application/xml");
            const tabs = xmlDoc.getElementsByTagName("tab");
            Array.from(tabs).forEach((element) => {
              const labels = element.querySelector("labels");
              const firstLabel = labels?.firstElementChild;

              const existingTab = table.tabs.find((tab) => tab.id === element.getAttribute("id"));

              if (existingTab) {
                const labelProp = existingTab.langProps.find((lp) => lp.name === "Label");
                if (labelProp) {
                  labelProp.translation.push({
                    code: form.props?.lang ?? "",
                    translation: firstLabel?.getAttribute("description") ?? "",
                  });
                }
              } else {
                table.tabs.push({
                  id: element.getAttribute("id") ?? "",
                  name: form.name,
                  props: { uniqueName: form.props?.uniqueName ?? "", lang: form.props?.lang ?? "", formId:form.id },
                  langProps: [
                    {
                      name: "Label",

                      translation: [
                        {
                          code: form.props?.lang ?? "",
                          translation: firstLabel?.getAttribute("description") ?? "",
                        },
                      ],
                    },
                  ],
                });
              }
            });
          }
        });
        this.onLog("Exporting form tabs", "info");
        table.tabs.forEach((tab) => {
          const row = [
            `${tab.id}`,
            table.logicalName,
            tab.name,
            `{${tab.props?.uniqueName ?? ""}}`,
            `{${tab.props?.formId ?? ""}}`,
            ...tab.langProps
              .filter((lp) => lp.name === "Label")
              .flatMap((lp) => lp.translation.map((trans) => trans.translation)),
          ];
          wsheet.addRow(row);
        });
      });
    }

    this.styleSheet(wsheet);

    wsheet = workbook.addWorksheet("Forms Sections");
    wsheet.addRow([
      "Section Id",
      "Entity Logical Name",
      "Form Name",
      "Form Unique Id",
      "Form Id",
      "Tab Name",
      ...this.outputLangs.map((lang) => lang.code),
    ]);

    if (this.vm.options.formSections) {
      this.vm.selectedTables.forEach((table) => {
        table.forms.forEach((form) => {
          if (form.props) {
            const xmlDoc = new DOMParser().parseFromString(form.props.formXml, "application/xml");
            const sections = xmlDoc.getElementsByTagName("section");
            Array.from(sections).forEach((element) => {
              const labels = element.querySelector("labels");
              const firstLabel = labels?.firstElementChild;
              const parentTab = element.closest("tab");
              const parentTabName = parentTab?.getAttribute("name") ?? "";
              const existingSection = table.sections.find((sec) => sec.id === element.getAttribute("id"));
              if (existingSection) {
                const labelProp = existingSection.langProps.find((lp) => lp.name === "Label");
                if (labelProp) {
                  labelProp.translation.push({
                    code: form.props?.lang ?? "",
                    translation: firstLabel?.getAttribute("description") ?? "",
                  });
                }
              } else {
                table.sections.push({
                  id: element.getAttribute("id") ?? "",
                  name: form.name,
                  props: {
                    uniqueName: form.props?.uniqueName ?? "",
                    lang: form.props?.lang ?? "",
                    formId: form.id,
                    tabName: parentTabName,
                  },
                  langProps: [
                    {
                      name: "Label",
                      translation: [
                        {
                          code: form.props?.lang ?? "",
                          translation: firstLabel?.getAttribute("description") ?? "",
                        },
                      ],
                    },
                  ],
                });
              }
            });
          }
        });

        table.sections.forEach((section) => {
          const row = [
            `${section.id}`,
            table.logicalName,
            section.name,
            `{${section.props?.uniqueName ?? ""}}`,
            `{${section.props?.formId ?? ""}}`,
            section.props?.tabName ?? "",
            ...section.langProps
              .filter((lp) => lp.name === "Label")
              .flatMap((lp) => lp.translation.map((trans) => trans.translation)),
          ];
          wsheet.addRow(row);
        });
      });
    }

    this.styleSheet(wsheet);

    wsheet = workbook.addWorksheet("Forms Fields");
    wsheet.addRow([
      "Label Id",
      "Entity Logical Name",
      "Form Name",
      "Form Unique Id",
      "Form Id",
      "Tab Name",
      "Section Name",
      "Attribute",
      "Display Name",
      ...this.outputLangs.map((lang) => lang.code),
    ]);

    if (this.vm.options.formFields) {
      this.vm.selectedTables.forEach((table) => {
        table.forms.forEach((form) => {
          if (form.props) {
            const xmlDoc = new DOMParser().parseFromString(form.props.formXml, "application/xml");
            const fields = xmlDoc.getElementsByTagName("cell");
            Array.from(fields).forEach((element) => {
              const labels = element.querySelector("labels");
              const firstLabel = labels?.firstElementChild;
              const parentTab = element.closest("tab");
              const tab = parentTab?.querySelector("labels")?.firstElementChild?.getAttribute("description") ?? "";
              const parentSection = element.closest("section");
              const section =
                parentSection?.querySelector("labels")?.firstElementChild?.getAttribute("description") ?? "";
              const control = element.querySelector("control");
              const attributeName = control?.getAttribute("id") ?? "";
              const displayName = table.fields
                .find((f) => f.name === attributeName)
                ?.langProps.find((lp) => lp.name === "DisplayName")?.translation[0].translation;

              const existingField = table.formFields.find((field) => field.id === element.getAttribute("id"));
              if (existingField) {
                const labelProp = existingField.langProps.find((lp) => lp.name === "Label");
                if (labelProp) {
                  labelProp.translation.push({
                    code: form.props?.lang ?? "",
                    translation: firstLabel?.getAttribute("description") ?? "",
                  });
                }
              } else {
                table.formFields.push({
                  id: element.getAttribute("id") ?? "",
                  name: form.name,
                  type: tab,
                  props: {
                    uniqueName: form.props?.uniqueName ?? "",
                    formId: form.id,
                    attributeName: attributeName,
                    displayName: displayName ?? attributeName ?? "",
                    tabName: tab,
                    sectionName: section,
                    lang: form.props?.lang ?? "",
                  },
                  langProps: [
                    {
                      name: "Label",
                      translation: [
                        {
                          code: form.props?.lang ?? "",
                          translation: firstLabel?.getAttribute("description") ?? "",
                        },
                      ],
                    },
                  ],
                });
              }
            });
          }
        });
        table.formFields.forEach((field) => {
          const row = [
            `${field.id}`,
            table.logicalName,
            field.name,
            `{${field.props?.uniqueName ?? ""}}`,
            `{${field.props?.formId ?? ""}}`,
            field.props?.tabName ?? "",
            field.props?.sectionName ?? "",
            field.props?.attributeName ?? "",
            field.props?.displayName ?? "",
            ...field.langProps
              .filter((lp) => lp.name === "Label")
              .flatMap((lp) => lp.translation.map((trans) => trans.translation)),
          ];
          wsheet.addRow(row);
        });
      });
    }

    this.styleSheet(wsheet);
  }

  private async exportSiteMap(workbook: ExcelJS.Workbook) {
    const areaSheet = workbook.addWorksheet("SiteMap Areas");
    areaSheet.addRow(["SiteMap Name", "SiteMap Id", "Area Id", "Type", ...this.outputLangs.map((lang) => lang.code)]);
    const groupSheet = workbook.addWorksheet("SiteMap Groups");
    groupSheet.addRow([
      "SiteMap Name",
      "SiteMap Id",
      "Area Id",
      "Group Id",
      "Type",
      ...this.outputLangs.map((lang) => lang.code),
    ]);
    const subareaSheet = workbook.addWorksheet("SiteMap SubAreas");
    subareaSheet.addRow([
      "SiteMap Name",
      "SiteMap Id",
      "Area Id",
      "Group Id",
      "SubArea Id",
      "Type",
      ...this.outputLangs.map((lang) => lang.code),
    ]);

    this.vm.siteMaps = await this.dvSvc.getSiteMaps(this.vm.solution?.solutionId ?? "");

    if (this.vm.siteMaps.length === 0) {
      this.onLog("No site map found for the selected solution.", "warning");
      return;
    }

    this.vm.siteMaps.forEach((sm) => {
      const xmlDoc = new DOMParser().parseFromString(sm.props?.sitemapXml, "application/xml");
      const areas = xmlDoc.getElementsByTagName("Area");
      Array.from(areas).forEach((area) => {
        const areaId = area.getAttribute("Id") ?? "";
        const areaTitles = area.querySelector("Titles");
        const areaDescriptions = area.querySelector("Descriptions");
        this.vm.siteAreas.push({
          id: areaId,
          name: area.getAttribute("title") ?? "",
          langProps: [
            {
              name: "Title",
              translation: areaTitles?.children
                ? Array.from(areaTitles.children).map((label) => ({
                    code: label.getAttribute("LCID") ?? "",
                    translation: label.getAttribute("Title") ?? "",
                  }))
                : [],
            },
            {
              name: "Description",
              translation: areaDescriptions?.children
                ? Array.from(areaDescriptions.children).map((label) => ({
                    code: label.getAttribute("LCID") ?? "",
                    translation: label.getAttribute("Description") ?? "",
                  }))
                : [],
            },
          ],
          props: { uniqueName: sm.props?.uniqueName ?? "", siteName: sm.name ?? "", siteId: sm.id ?? "" },
        });

        const groups = area.getElementsByTagName("Group");
        Array.from(groups).forEach((group) => {
          const groupId = group.getAttribute("Id") ?? "";
          let groupTitles = group.querySelector("Titles");
          if (groupTitles?.parentElement != group) groupTitles = null;
          const groupDescriptions = group.querySelector("Descriptions");
          this.vm.siteGroups.push({
            id: groupId,
            name: group.getAttribute("title") ?? "",
            langProps: [
              {
                name: "Title",
                translation: groupTitles?.children
                  ? Array.from(groupTitles.children).map((label) => ({
                      code: label.getAttribute("LCID") ?? "",
                      translation: label.getAttribute("Title") ?? "",
                    }))
                  : [],
              },
              {
                name: "Description",
                translation: groupDescriptions?.children
                  ? Array.from(groupDescriptions.children).map((label) => ({
                      code: label.getAttribute("LCID") ?? "",
                      translation: label.getAttribute("Description") ?? "",
                    }))
                  : [],
              },
            ],
            props: { uniqueName: sm.props?.uniqueName ?? "", siteName: sm.name ?? "", siteId: sm.id ?? "" },
          });

          const subareas = group.getElementsByTagName("SubArea");
          Array.from(subareas).forEach((subarea) => {
            const subareaId = subarea.getAttribute("Id") ?? "";
            const subareaTitles = subarea.querySelector("Titles");
            const subareaDescriptions = subarea.querySelector("Descriptions");
            this.vm.siteSubAreas.push({
              id: subareaId,
              name: subarea.getAttribute("title") ?? "",
              langProps: [
                {
                  name: "Title",
                  translation: subareaTitles?.children
                    ? Array.from(subareaTitles.children).map((label) => ({
                        code: label.getAttribute("LCID") ?? "",
                        translation: label.getAttribute("Title") ?? "",
                      }))
                    : [],
                },
                {
                  name: "Description",
                  translation: subareaDescriptions?.children
                    ? Array.from(subareaDescriptions.children).map((label) => ({
                        code: label.getAttribute("LCID") ?? "",
                        translation: label.getAttribute("Description") ?? "",
                      }))
                    : [],
                },
              ],
              props: {
                uniqueName: sm.props?.uniqueName ?? "",
                siteName: sm.name ?? "",
                siteId: sm.id ?? "",
                groupId: groupId,
              },
            });
          });
        });
      });
    });
    this.onLog("Areas:" + this.vm.siteAreas.length, "success");

    for (const area of this.vm.siteAreas) {
      for (const langprop of area.langProps) {
        areaSheet.addRow([
          area.props?.siteName ?? "",
          area.props?.siteId ?? "",
          area.id,
          langprop.name,
          ...langprop.translation.flatMap((trans) => trans.translation),
        ]);
      }
    }

    for (const group of this.vm.siteGroups) {
      for (const langprop of group.langProps) {
        groupSheet.addRow([
          group.props?.siteName ?? "",
          group.props?.siteId ?? "",
          group.props?.uniqueName ?? "",
          group.id,
          langprop.name,
          ...langprop.translation.flatMap((trans) => trans.translation),
        ]);
      }
    }

    for (const subarea of this.vm.siteSubAreas) {
      for (const langprop of subarea.langProps) {
        subareaSheet.addRow([
          subarea.props?.siteName ?? "",
          subarea.props?.siteId ?? "",
          subarea.props?.uniqueName ?? "",
          subarea.props?.groupId ?? "",
          subarea.id,
          langprop.name,
          ...langprop.translation.flatMap((trans) => trans.translation),
        ]);
      }
    }

    this.styleSheet(areaSheet);
    this.styleSheet(groupSheet);
    this.styleSheet(subareaSheet);
  }

  private async exportDashboards(workbook: ExcelJS.Workbook) {
    let wsheet = workbook.addWorksheet("Dashboards");
    wsheet.addRow(["Form Unique Id", "Form Id", "Type", ...this.outputLangs.map((lang) => lang.code)]);

    if (!this.vm.options.dashboards) {
      this.styleSheet(wsheet);
      return;
    }
    this.onLog("Fetching dashboards...", "info");

    if (this.vm.dashboards.length === 0) {
      this.onLog("No dashboards found for the selected solution.", "info");
      return;
    }

    for (const dashboard of this.vm.dashboards.filter((d) => d.props?.base)) {
      if (this.vm.options.labels()) {
        const names = await this.dvSvc.getLocLabels("systemforms", dashboard.id, "name");
        dashboard.langProps.push({ name: "name", translation: names });
      }
      if (this.vm.options.descriptions()) {
        const descriptions = await this.dvSvc.getLocLabels("systemforms", dashboard.id, "description");
        dashboard.langProps.push({ name: "description", translation: descriptions });
      }
      dashboard.langProps.forEach((langProp) => {
        const row = [
          `{${dashboard.props?.uniqueName ?? ""}}`,
          `{${dashboard.id}}`,
          langProp.name,
          ...langProp.translation.map((trans) => trans.translation),
        ];
        wsheet.addRow(row);
      });
    }

    for (const dashboard of this.vm.dashboards) {
      const xmlDoc = new DOMParser().parseFromString(dashboard.props?.formXml ?? "", "application/xml");
      const tabs = xmlDoc.getElementsByTagName("tab");
      Array.from(tabs).forEach((tab) => {
        const labels = tab.querySelector("labels");
        const tabLabel = labels?.firstElementChild;
        const tabName = tabLabel?.getAttribute("description") ?? "";
        const existingTab = this.vm.dashboardTabs.find((t) => t.id === tab.getAttribute("id"));
        if (existingTab) {
          const labelProp = existingTab.langProps.find((lp) => lp.name === "Label");
          if (labelProp) {
            labelProp.translation.push({
              code: dashboard.props?.lang ?? "",
              translation: tabLabel?.getAttribute("description") ?? "",
            });
          }
        } else {
          this.vm.dashboardTabs.push({
            id: tab.getAttribute("id") ?? "",
            name: tabName,
            props: { formUniqueName: dashboard.props?.uniqueName ?? "", formId: dashboard.props?.formId ?? "" },
            langProps: [
              {
                name: "Label",
                translation: [
                  { code: dashboard.props?.lang ?? "", translation: tabLabel?.getAttribute("description") ?? "" },
                ],
              },
            ],
          });
        }
        const sections = tab.getElementsByTagName("section");
        Array.from(sections).forEach((section) => {
          const labels = section.querySelector("labels");
          const sectionLabel = labels?.firstElementChild;
          const sectionName = sectionLabel?.getAttribute("description") ?? "";
          const existingSection = this.vm.dashboardSections.find((s) => s.id === section.getAttribute("id"));
          if (existingSection) {
            const labelProp = existingSection.langProps.find((lp) => lp.name === "Label");
            if (labelProp) {
              labelProp.translation.push({
                code: dashboard.props?.lang ?? "",
                translation: sectionLabel?.getAttribute("description") ?? "",
              });
            }
          } else {
            this.vm.dashboardSections.push({
              id: section.getAttribute("id") ?? "",
              name: dashboard.name,
              props: {
                formUniqueName: dashboard.props?.uniqueName ?? "",
                formId: dashboard.props?.formId ?? "",
                tabName: tabName,
                sectionName: sectionName,
              },
              langProps: [
                {
                  name: "Label",
                  translation: [
                    { code: dashboard.props?.lang ?? "", translation: sectionLabel?.getAttribute("description") ?? "" },
                  ],
                },
              ],
            });
          }
          const cells = section.getElementsByTagName("cell");
          Array.from(cells).forEach((cell) => {
            const labels = cell.querySelector("labels");
            const cellLabel = labels?.firstElementChild;

            const control = cell.querySelector("control");
            if (!control) return;
            const attributeName = control?.getAttribute("id") ?? "";
            const displayName = this.vm.selectedTables
              .flatMap((t) => t.fields)
              .find((f) => f.name === attributeName)
              ?.langProps.find((lp) => lp.name === "DisplayName")?.translation[0].translation;
            const existingField = this.vm.dashboardFields.find((f) => f.id === cell.getAttribute("id"));
            if (existingField) {
              const labelProp = existingField.langProps.find((lp) => lp.name === "Label");
              if (labelProp) {
                labelProp.translation.push({
                  code: dashboard.props?.lang ?? "",
                  translation: cellLabel?.getAttribute("description") ?? "",
                });
              }
            } else {
              this.vm.dashboardFields.push({
                id: cell.getAttribute("id") ?? "",
                name: dashboard.name,
                props: {
                  formUniqueName: dashboard.props?.uniqueName ?? "",
                  formId: dashboard.props?.formId ?? "",
                  tabName: tabName,
                  sectionName: sectionName,
                  attributeName: attributeName,
                  displayName: displayName ?? "",
                  lang: dashboard.props?.lang ?? "",
                },
                langProps: [
                  {
                    name: "Label",
                    translation: [
                      { code: dashboard.props?.lang ?? "", translation: cellLabel?.getAttribute("description") ?? "" },
                    ],
                  },
                ],
              });
            }
          });
        });
      });
    }

    wsheet = workbook.addWorksheet("Dashboards Tabs");
    wsheet.addRow([
      "Tab Id",
      "Form Name test",
      "Form Unique Id",
      "Form Id",
      ...this.outputLangs.map((lang) => lang.code),
    ]);
    for (const tab of this.vm.dashboardTabs) {
      const row = [
        `${tab.id}`,
        tab.name,
        `{${tab.props?.formUniqueName ?? ""}}`,
        `{${tab.props?.formId ?? ""}}`,
        ...tab.langProps.flatMap((lp) => lp.translation.map((trans) => trans.translation)),
      ];
      wsheet.addRow(row);
    }

    wsheet = workbook.addWorksheet("Dashboards Sections");
    wsheet.addRow([
      "Section Id",
      "Form Name",
      "Form Unique Id",
      "Form Id",
      "Tab Name",
      ...this.outputLangs.map((lang) => lang.code),
    ]);
    for (const section of this.vm.dashboardSections) {
      const row = [
        `${section.id}`,
        section.name,
        `{${section.props?.formUniqueName ?? ""}}`,
        `{${section.props?.formId ?? ""}}`,
        section.props?.tabName ?? "",
        ...section.langProps
          .filter((lp) => lp.name === "Label")
          .flatMap((lp) => lp.translation.map((trans) => trans.translation)),
      ];
      wsheet.addRow(row);
    }
    wsheet = workbook.addWorksheet("Dashboards Fields");
    wsheet.addRow([
      "Label Id",
      "Form Name",
      "Form Unique Id",
      "Form Id",
      "Tab Name",
      "Section Name",
      ...this.outputLangs.map((lang) => lang.code),
    ]);
    for (const field of this.vm.dashboardFields) {
      const row = [
        `${field.id}`,
        field.name,
        field.props?.formUniqueName ?? "",
        field.props?.formId ?? "",
        field.props?.tabName ?? "",
        field.props?.sectionName ?? "",
        ...field.langProps
          .filter((lp) => lp.name === "Label")
          .flatMap((lp) => lp.translation.map((trans) => trans.translation)),
      ];
      wsheet.addRow(row);
    }
    this.styleSheet(wsheet);
  }

  private styleSheet(sheet: ExcelJS.Worksheet): void {
    this.styleHeaderRow(sheet.getRow(1));
    this.applyAlternatingRowFill(sheet, 1);
  }

  private styleHeaderRow(row: ExcelJS.Row): void {
    row.font = { bold: true, color: { argb: "FFFFFFFF" } };
    row.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF0078D4" },
    };
    row.alignment = { vertical: "middle", horizontal: "left" };
    row.height = 20;
  }

  private applyAlternatingRowFill(worksheet: ExcelJS.Worksheet, headerRowIndex = 1): void {
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber > headerRowIndex && rowNumber % 2 === 0) {
        row.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFF3F4F6" },
        };
      }
    });
  }

  private downloadBlob(blob: Blob, fileName: string): void {
    const link = document.createElement("a");
    const url = URL.createObjectURL(blob);

    link.setAttribute("href", url);
    link.setAttribute("download", fileName);
    link.style.visibility = "hidden";

    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    URL.revokeObjectURL(url);
  }
}
