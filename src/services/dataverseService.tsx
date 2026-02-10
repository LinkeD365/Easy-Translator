import { Solution } from "../model/solution";
import { LangTranslation, LanguageDef, Table, OptionSetDef } from "../model/viewModel";
import * as locale from "locale-codes";

interface dvServiceProps {
  connection: ToolBoxAPI.DataverseConnection | null;
  dvApi: DataverseAPI.API;
  onLog: (message: string, type?: "info" | "success" | "warning" | "error") => void;
}

export class dvService {
  connection: ToolBoxAPI.DataverseConnection | null;
  dvApi: DataverseAPI.API;
  onLog: (message: string, type?: "info" | "success" | "warning" | "error") => void;
  batchSize = 2;
  constructor(props: dvServiceProps) {
    this.connection = props.connection;
    this.dvApi = props.dvApi;
    this.onLog = props.onLog;
  }

  async getSolutions(): Promise<Solution[]> {
    try {
      const solutionsData = await this.dvApi.queryData(
        "solutions?$filter=(isvisible eq true)&$select=friendlyname,uniquename&$orderby=createdon desc",
      );
      const solutions: Solution[] = (solutionsData.value as any[]).map((sol: any) => {
        const solution = new Solution();
        solution.name = sol.friendlyname;
        solution.uniqueName = sol.uniquename;
        solution.solutionId = sol.solutionid;
        return solution;
      });
      this.onLog("Solutions fetched successfully", "success");
      return solutions;
    } catch (error: any) {
      this.onLog(`Error fetching solutions: ${error.message || error}`, "error");
      throw error;
    }
  }

  async getSolutionTables(solutionId: string): Promise<Table[]> {
    this.onLog(`Fetching tables for solution: ${solutionId}`, "info");
    if (!this.connection) {
      throw new Error("No connection available");
    }

    var fetchXml = [
      "<fetch>",
      "  <entity name='entity'>",
      "    <attribute name='entityid'/>",
      "    <attribute name='entitysetname'/>",
      "    <attribute name='logicalname'/>",
      "    <attribute name='name'/>",
      "    <link-entity name='solutioncomponent' from='objectid' to='entityid' alias='sc'>",
      "      <filter>",
      "        <condition attribute='solutionid' operator='eq' value='",
      solutionId,
      "'/>",
      "      </filter>",
      "    </link-entity>",
      "  </entity>",
      "</fetch>",
    ].join("");

    try {
      const componentsData = await this.dvApi.fetchXmlQuery(fetchXml);
      const componentArray = Array.isArray((componentsData as any).value)
        ? ((componentsData as any).value as Record<string, any>[])
        : [];
      const tablePromises = componentArray.map(async (comp) => {
        if (!comp.entityid) return null;
        try {
          const entityMeta = await this.dvApi.queryData(`EntityDefinitions(${comp.entityid})`);
          const src: any = Array.isArray((entityMeta as any)?.value)
            ? (entityMeta as any).value[0]
            : (entityMeta as any);

          const tm = new Table(
            src?.DisplayName?.LocalizedLabels?.[0]?.Label || src?.LogicalName || "",
            src?.EntitySetName || "",
            src?.LogicalName || "",
            src?.MetadataId || "",
            src?.ObjectTypeCode || "",
          );

          return tm;
        } catch (err) {
          this.onLog(`Failed to fetch entity metadata for id ${comp.entityid}: ${(err as Error).message}`, "warning");
        }
      });

      const tables = await Promise.all(tablePromises);
      const filteredTables = tables.filter((table): table is Table => table !== null && table !== undefined);
      this.onLog(`Fetched ${filteredTables.length} tables for solution: ${solutionId}`, "success");
      return filteredTables.sort((a, b) => a.label.localeCompare(b.label));
    } catch (err) {
      this.onLog(`Error fetching solution tables for ${solutionId}: ${(err as Error).message}`, "error");
      throw err;
    }
  }

  async getLanguages(): Promise<LanguageDef[]> {
    this.onLog("Fetching languages");
    try {
      const langData = await this.dvApi.execute({
        operationName: "RetrieveAvailableLanguages",
        operationType: "function",
      });
      const returnLangs = (langData.LocaleIds as any[]).map((lang: any) => {
        const llang = locale.getByLCID(lang);
        // Map the language data to LanguageDef objects or your desired structure
        return {
          code: lang,
          name: llang.name + " (" + llang.location + ")",
        } as LanguageDef;
      });
      return returnLangs;
    } catch (error) {
      this.onLog("Failed to fetch languages: " + (error as Error).message, "error");
      throw error;
    }
  }

  async getBaseLanguage(): Promise<LanguageDef> {
    try {
      const fetchXml = [
        "<fetch>",
        "  <entity name='organization'>",
        "    <attribute name='languagecode'/>",
        "  </entity>",
        "</fetch>",
      ].join("");
      const queryResult = await this.dvApi.fetchXmlQuery(fetchXml);
      const languageCode = queryResult?.value?.[0]?.languagecode;
      if (!languageCode) {
        throw new Error("Base language code not found");
      }
      const llang = locale.getByLCID(Number(languageCode));
      // Map the language data to LanguageDef objects or your desired structure
      return {
        code: languageCode,
        name: llang.name + " (" + llang.location + ")",
      } as LanguageDef;
    } catch (error) {
      this.onLog("Failed to fetch base language: " + (error as Error).message, "error");
      throw error;
    }
  }

  async getTableMeta(table: Table): Promise<boolean> {
    return new Promise<boolean>(async (resolve, reject) => {
      try {
        if (table.langProps.length > 0) {
          resolve(true);
          return;
        }
        const entityMeta = await this.dvApi.getEntityMetadata(table.id, false, [
          "DisplayName",
          "DisplayCollectionName",
          "Description",
          "SchemaName",
          "LogicalName",
          "ObjectTypeCode",
        ]);
        const displayNameLabels = ((entityMeta as any)?.DisplayName?.LocalizedLabels as any[]) || [];
        table.langProps.push({
          name: "DisplayName",
          translation: [...displayNameLabels.map((label: any) => new LangTranslation(label.LanguageCode, label.Label))],
        });

        const displayCollectionLabels = ((entityMeta as any)?.DisplayCollectionName?.LocalizedLabels as any[]) || [];
        table.langProps.push({
          name: "DisplayCollectionName",
          translation: [
            ...displayCollectionLabels.map((label: any) => new LangTranslation(label.LanguageCode, label.Label)),
          ],
        });

        table.langProps.push({
          name: "Description",
          translation: [
            ...((entityMeta as any)?.Description?.LocalizedLabels as any[]).map(
              (label: any) => new LangTranslation(label.LanguageCode, label.Label),
            ),
          ],
        });
        resolve(true);
      } catch (error) {
        reject(error);
      }
    });
  }

  async getTableFields(table: Table): Promise<boolean> {
    return new Promise<boolean>(async (resolve, reject) => {
      try {
        if (table.fields.length > 0) {
          resolve(true);
          return;
        }
        const entityMeta = await this.dvApi.getEntityRelatedMetadata(table.logicalName, "Attributes", [
          "DisplayName",
          "Description",
          "LogicalName",
        ]);
        entityMeta.value.forEach((fld: any) => {
          table.fields.push({
            id: fld.MetadataId,
            name: fld.LogicalName,
            langProps: [
              {
                name: "DisplayName",
                translation: [
                  ...((fld as any)?.DisplayName?.LocalizedLabels as any[]).map(
                    (label: any) => new LangTranslation(label.LanguageCode, label.Label),
                  ),
                ],
              },
              {
                name: "Description",
                translation: [
                  ...((fld as any)?.Description?.LocalizedLabels as any[]).map(
                    (label: any) => new LangTranslation(label.LanguageCode, label.Label),
                  ),
                ],
              },
            ],
          });
        });
        resolve(true);
      } catch (error) {
        reject(error);
      }
    });
  }

  async getRelationships(table: Table): Promise<boolean> {
    return new Promise<boolean>(async (resolve, reject) => {
      try {
        if (table.relationships.length > 0) {
          resolve(true);
          return;
        }

        const relList = ["OneToManyRelationships", "ManyToOneRelationships"];

        for await (const relType of relList) {
          const entityMeta = await this.dvApi.getEntityRelatedMetadata(
            table.logicalName,
            relType as unknown as DataverseAPI.EntityRelatedMetadataPath,
            ["AssociatedMenuConfiguration", "ReferencedEntity", "SchemaName", "ReferencingEntity"],
          );
          const relMeta = entityMeta as { value?: any[] };

          (relMeta.value ?? [])
            .filter(
              (rel: any) =>
                rel?.AssociatedMenuConfiguration?.Behavior == "UseLabel" &&
                !table.relationships.some((r) => r.id === rel?.MetadataId),
            )
            .forEach((rel: any) => {
              table.relationships.push({
                id: rel.MetadataId,
                name: rel.LogicalName,
                type: relType,
                props: {
                  SchemaName: rel.SchemaName,
                  ReferencingEntity: rel.ReferencingEntity,
                },
                langProps: [
                  {
                    name: "DisplayName",
                    translation: [
                      ...((rel as any)?.AssociatedMenuConfiguration?.Label?.LocalizedLabels as any[]).map(
                        (label: any) => new LangTranslation(label.LanguageCode, label.Label),
                      ),
                    ],
                  },
                ],
              });
            });
        }
        this.onLog("O2M and M2O Relationships loaded");

        const entityMeta = await this.dvApi.getEntityRelatedMetadata(table.logicalName, "ManyToManyRelationships", [
          "Entity1AssociatedMenuConfiguration",
          "Entity2AssociatedMenuConfiguration",
          "IntersectEntityName",
          "SchemaName",
          "Entity1LogicalName",
          "Entity2LogicalName",
        ]);

        const relMeta = entityMeta as { value?: any[] };
        (relMeta.value ?? [])
          .filter(
            (rel: any) =>
              rel?.Entity1AssociatedMenuConfiguration?.Behavior == "UseLabel" ||
              rel?.Entity2AssociatedMenuConfiguration?.Behavior == "UseLabel",
          )
          .forEach((rel: any) => {
            const amc =
              rel.Entity1LogicalName == table.logicalName
                ? rel.Entity1AssociatedMenuConfiguration
                : rel.Entity2AssociatedMenuConfiguration;

            table.relationships.push({
              id: rel.MetadataId,
              name: rel.SchemaName,
              type: "ManyToManyRelationship",
              props: {
                IntersectEntityName: rel.IntersectEntityName,
              },
              langProps: [
                {
                  name: "DisplayName",
                  translation: [
                    ...((amc as any)?.Label?.LocalizedLabels as any[]).map(
                      (label: any) => new LangTranslation(label.LanguageCode, label.Label),
                    ),
                  ],
                },
              ],
            });
          });
        this.onLog("ManyToMany Relationships loaded");
        resolve(true);
      } catch (error) {
        reject(error);
      }
    });
  }

  async getOptionSets(table: Table): Promise<boolean> {
    return new Promise<boolean>(async (resolve, reject) => {
      try {
        if (table.optionSets.length > 0) resolve(true);

        const em = await this.dvApi.queryData(
          `EntityDefinitions(LogicalName='${table.logicalName}')/Attributes/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?$select=LogicalName,AttributeType&$expand=OptionSet,GlobalOptionSet`,
        );

        em.value.forEach((attr: any) => {
          const optionSet = attr.OptionSet;
          if (optionSet && optionSet.Options) {
            optionSet.Options.forEach((option: any) => {
              const labelLangProps = {
                name: "Label",
                translation:
                  option.Label?.LocalizedLabels?.map(
                    (label: any) => new LangTranslation(label.LanguageCode, label.Label),
                  ) || [],
              };

              const descLangProps = {
                name: "Description",
                translation:
                  option.Description?.LocalizedLabels?.map(
                    (label: any) => new LangTranslation(label.LanguageCode, label.Label),
                  ) || [],
              };

              // Create separate entries for Label and Description
              const optionDefLabel = new OptionSetDef(
                attr.MetadataId,
                attr.LogicalName,
                [labelLangProps, descLangProps],
                false,
              );
              optionDefLabel.attributeLogicalName = attr.LogicalName;

              optionDefLabel.optionValue = option.Value;
              optionDefLabel.type = attr.AttributeType;
              optionDefLabel.isGlobal = optionSet.IsGlobal;
              table.optionSets.push(optionDefLabel);
            });
          }
        });
        this.onLog(`Loaded ${table.optionSets.length} optionset entries for ${table.logicalName}`);
        resolve(true);
      } catch (error) {
        this.onLog(`Error loading optionsets: ${(error as Error).message}`, "error");
        reject(error);
      }
    });
  }

  async getBooleans(table: Table): Promise<boolean> {
    return new Promise<boolean>(async (resolve, reject) => {
      try {
        if (table.optionSets.filter((os) => os.type === "Boolean").length > 0) {
          resolve(true);
          return;
        }

        const em = await this.dvApi.queryData(
          `EntityDefinitions(LogicalName='${table.logicalName}')/Attributes/Microsoft.Dynamics.CRM.BooleanAttributeMetadata?$select=LogicalName,AttributeType&$expand=OptionSet,GlobalOptionSet`,
        );
        em.value.forEach((attr: any) => {
          const optionSet = attr.OptionSet;
          if (optionSet && optionSet.TrueOption) {
            const trueOption = optionSet.TrueOption;

            const labelLangProps = {
              name: "Label",
              translation:
                trueOption.Label?.LocalizedLabels?.map(
                  (label: any) => new LangTranslation(label.LanguageCode, label.Label),
                ) || [],
            };

            const descLangProps = {
              name: "Description",
              translation:
                trueOption.Description?.LocalizedLabels?.map(
                  (label: any) => new LangTranslation(label.LanguageCode, label.Label),
                ) || [],
            };

            // Create separate entries for Label and Description
            const optionDefLabel = new OptionSetDef(
              attr.MetadataId,
              attr.LogicalName,
              [labelLangProps, descLangProps],
              false,
            );
            optionDefLabel.attributeLogicalName = attr.LogicalName;
            optionDefLabel.optionValue = trueOption.Value;
            optionDefLabel.type = attr.AttributeType;
            optionDefLabel.isGlobal = optionSet.IsGlobal;
            table.optionSets.push(optionDefLabel);
          }
          if (optionSet && optionSet.FalseOption) {
            const falseOption = optionSet.FalseOption;

            const labelLangProps = {
              name: "Label",
              translation:
                falseOption.Label?.LocalizedLabels?.map(
                  (label: any) => new LangTranslation(label.LanguageCode, label.Label),
                ) || [],
            };

            const descLangProps = {
              name: "Description",
              translation:
                falseOption.Description?.LocalizedLabels?.map(
                  (label: any) => new LangTranslation(label.LanguageCode, label.Label),
                ) || [],
            };

            // Create separate entries for Label and Description
            const optionDefLabel = new OptionSetDef(
              attr.MetadataId,
              attr.LogicalName,
              [labelLangProps, descLangProps],
              false,
            );
            optionDefLabel.attributeLogicalName = attr.LogicalName;
            optionDefLabel.optionValue = falseOption.Value;
            optionDefLabel.type = attr.AttributeType;
            optionDefLabel.isGlobal = optionSet.IsGlobal;
            table.optionSets.push(optionDefLabel);
          }
        });

        this.onLog(`Loaded  ${table.optionSets.length} optionset entries for ${table.logicalName}`);
        resolve(true);
      } catch (error) {
        this.onLog(`Error loading optionsets: ${(error as Error).message}`, "error");
        reject(error);
      }
    });
  }

  async getViews(table: Table): Promise<boolean> {
    return new Promise<boolean>(async (resolve, reject) => {
      try {
        var fetchXml = [
          "<fetch>",
          "  <entity name='savedquery'>",
          "    <attribute name='savedqueryid'/>",
          "    <attribute name='name'/>",
          "    <attribute name='querytype'/>",
          "    <filter>",
          "      <condition attribute='returnedtypecode' operator='eq' value='",
          table.code,
          "'/>",
          "    </filter>",
          "  </entity>",
          "</fetch>",
        ].join("");

        const viewsData = await this.dvApi.fetchXmlQuery(fetchXml);

        await Promise.all(
          (viewsData.value as any[]).map(async (view: any) => {
            return new Promise<void>(async (resolve) => {
              const label = await this.getLocLabels("savedqueries", view.savedqueryid, "name");
              const decription = await this.getLocLabels("savedqueries", view.savedqueryid, "description");
              const queryTypeMap: Record<number, string> = {
                0: "Public View",
                2: "Associated View",
                1: "Advanced Search View",
                4: "Quick Find Search View",
                64: "Lookup view",
                2048: "Saved query used for workflow templates and email templates",
                8192: "Outlook offine template",
              };
              table.views.push({
                id: view.savedqueryid,
                name: view.name,
                type: queryTypeMap[view.querytype] || String(view.querytype),
                langProps: [
                  { name: "Label", translation: label },
                  { name: "Description", translation: decription },
                ],
              });
              resolve();
            });
          }),
        );
        resolve(true);
      } catch (error) {
        reject(error);
      }
    });
  }

  async getCharts(table: Table): Promise<boolean> {
    return new Promise<boolean>(async (resolve, reject) => {
      try {
        var fetchXml = [
          "<fetch>",
          "  <entity name='savedqueryvisualization'>",
          "    <attribute name='savedqueryvisualizationid'/>",
          "    <attribute name='name'/>",
          "    <filter>",
          "      <condition attribute='primaryentitytypecode' operator='eq' value='",
          table.code,
          "'/>",
          "    </filter>",
          "  </entity>",
          "</fetch>",
        ].join("");

        const chartsData = await this.dvApi.fetchXmlQuery(fetchXml);
        await Promise.all(
          (chartsData.value as any[]).map(async (chart: any) => {
            return new Promise<void>(async (resolve) => {
              const label = await this.getLocLabels(
                "savedqueryvisualizations",
                chart.savedqueryvisualizationid,
                "name",
              );
              const decription = await this.getLocLabels(
                "savedqueryvisualizations",
                chart.savedqueryvisualizationid,
                "description",
              );
              table.charts.push({
                id: chart.savedqueryvisualizationid,
                name: chart.name,
                langProps: [
                  { name: "Label", translation: label },
                  { name: "Description", translation: decription },
                ],
              });
              resolve();
            });
          }),
        );
        resolve(true);
      } catch (error) {
        reject(error);
      }
    });
  }

  async getForms(table: Table, lang: LanguageDef, base: boolean ): Promise<boolean> {
    return new Promise<boolean>(async (resolve, reject) => {
      try {
        var fetchXml = [
          "<fetch>",
          "  <entity name='systemform'>",
          "    <attribute name='formid'/>",
          "    <attribute name='name'/>",
          "    <attribute name='formxml'/>",
          "   <attribute name='type'/>",
          "    <attribute name='formidunique'/>",
          "    <filter>",
          "      <condition attribute='objecttypecode' operator='eq' value='",
          table.code,
          "'/>",
          "    </filter>",
          "  </entity>",
          "</fetch>",
        ].join("");

        const formsData = await this.dvApi.fetchXmlQuery(fetchXml);
        
        await Promise.all(
          (formsData.value as any[]).map(async (form: any) => {
            return new Promise<void>(async (resolve) => {
              const label = await this.getLocLabels("systemforms", form.formid, "name");
              const decription = await this.getLocLabels("systemforms", form.formid, "description");
              const formTypeMap: Record<number, string> = {
                6: "Quick View Form",
                2: "Main",
                7: "Quick Create Form",
              };
              table.forms.push({
                id: form.formid,
                name: form.name,
                type: formTypeMap[form.type] || String(form.type),
                props: { formXml: form.formxml, uniqueName: form.formidunique, lang: lang.code, base: base },
                langProps: [
                  { name: "Label", translation: label },
                  { name: "Description", translation: decription },
                ],
              });
              resolve();
            });
          }),
        );
        resolve(true);
      } catch (error) {
        reject(error);
      }
    });
  }

  async getLocLabels(tableClassName: string, objectId: string, attributeName: string): Promise<LangTranslation[]> {
    return new Promise<LangTranslation[]>(async (resolve, reject) => {
      try {
        const odata = `RetrieveLocLabels(EntityMoniker=@p1,AttributeName=@p2,IncludeUnpublished=false)?@p1={'@odata.id':'${tableClassName}(${objectId})'}&@p2='${attributeName}'`;
        const result = await this.dvApi.queryData(odata);
        const returnTranslations: LangTranslation[] = [];
        (((result as any).Label?.LocalizedLabels as any[]) || []).forEach((label: any) => {
          returnTranslations.push(new LangTranslation(label.LanguageCode, label.Label));
        });

        resolve(returnTranslations); // Map the result to LangTranslation objects as needed
      } catch (error) {
        reject(error);
      }
    });
  }

  async updateLanguage(lang: string, userId: string): Promise<void> {
    return new Promise<void>(async (resolve, reject) => {
      try {
        await this.dvApi.update("usersettingscollection", userId, {
          uilanguageid: lang,
          localeid: lang,
        });

        await new Promise(r => setTimeout(r, 2000));
        resolve();
      } catch (error) {
        this.onLog("Failed to update language: " + (error as Error).message, "error");
        reject(error);
      }
    });
  }

  async getUserLanguage(): Promise<{ uiLocale: string; locale: string; userid: string }> {
    try {
      const user = await this.dvApi.execute({
        operationName: "WhoAmI",
        operationType: "function",
      });
      const setting = await this.dvApi.queryData(
        "usersettingscollection?$select=localeid,uilanguageid&$filter=(Microsoft.Dynamics.CRM.EqualUserId(PropertyName=%27systemuserid%27))",
      );
      return {
        uiLocale: setting.value[0].uilanguageid as string,
        locale: setting.value[0].localeid as string,
        userid: user.UserId as string,
      };
    } catch (error) {
      this.onLog("Failed to get user language: " + (error as Error).message, "error");
      throw error;
    }
  }
}
