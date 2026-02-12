import { Solution } from "../model/solution";
import { LangTranslation, LanguageDef, Table, OptionSetDef, SecInfo } from "../model/viewModel";
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
    if (!this.connection) {
      const error = "No connection available. Please connect first.";
      this.onLog(error, "error");
      throw new Error(error);
    }
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

  async getAllTables(): Promise<Table[]> {
    this.onLog(`Fetching all tables from environment`, "info");
    if (!this.connection) {
      const error = "No connection available. Please connect first.";
      this.onLog(error, "error");
      throw new Error(error);
    }

    try {
      const componentsData = await this.dvApi.getAllEntitiesMetadata([
        "DisplayName",
        "EntitySetName",
        "LogicalName",
        "MetadataId",
        "ObjectTypeCode",
      ]);
      const componentArray = Array.isArray((componentsData as any).value)
        ? ((componentsData as any).value as Record<string, any>[])
        : [];
      this.onLog(`Fetched ${componentArray.length} entities from environment`, "info");
      const tables = componentArray.map((comp) => {
        return new Table(
          comp.DisplayName?.LocalizedLabels?.[0]?.Label || comp.LogicalName || "",
          comp.EntitySetName || "",
          comp.LogicalName || "",
          comp.MetadataId || "",
          comp.ObjectTypeCode || "",
        );
      });

      const filteredTables = tables.filter(
        (table) => table.logicalName !== "solutioncomponent" && table.logicalName !== "entity",
      );
      return filteredTables.sort((a, b) => a.label.localeCompare(b.label));
    } catch (err) {
      this.onLog(`Error fetching all tables: ${(err as Error).message}`, "error");
      throw err;
    }
  }

  async getSolutionTables(solutionId: string): Promise<Table[]> {
    this.onLog(`Fetching tables for solution: ${solutionId}`, "info");
    if (!this.connection) {
      throw new Error("No connection available");
    }

    const fetchXml = [
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
    if (!this.connection) {
      const error = "No connection available. Please connect first.";
      this.onLog(error, "error");
      throw new Error(error);
    }
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
    if (!this.connection) {
      const error = "No connection available. Please connect first.";
      this.onLog(error, "error");
      throw new Error(error);
    }
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
    if (!this.connection) {
      const error = "No connection available. Please connect first.";
      this.onLog(error, "error");
      throw new Error(error);
    }

    if (table.langProps.length > 0) {
      return true;
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
        ...((entityMeta as any)?.Description?.LocalizedLabels || []).map(
          (label: any) => new LangTranslation(label.LanguageCode, label.Label),
        ),
      ],
    });
    return true;
  }

  async getTableFields(table: Table): Promise<boolean> {
    if (!this.connection) {
      const error = "No connection available. Please connect first.";
      this.onLog(error, "error");
      throw new Error(error);
    }

    if (table.fields.length > 0) {
      return true;
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
              ...((fld as any)?.DisplayName?.LocalizedLabels || []).map(
                (label: any) => new LangTranslation(label.LanguageCode, label.Label),
              ),
            ],
          },
          {
            name: "Description",
            translation: [
              ...((fld as any)?.Description?.LocalizedLabels || []).map(
                (label: any) => new LangTranslation(label.LanguageCode, label.Label),
              ),
            ],
          },
        ],
      });
    });
    return true;
  }

  async getRelationships(table: Table): Promise<boolean> {
    if (!this.connection) {
      const error = "No connection available. Please connect first.";
      this.onLog(error, "error");
      throw new Error(error);
    }

    if (table.relationships.length > 0) {
      return true;
    }

    const relList = ["OneToManyRelationships", "ManyToOneRelationships"];

    for (const relType of relList) {
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
                ...((amc as any)?.Label?.LocalizedLabels || []).map(
                  (label: any) => new LangTranslation(label.LanguageCode, label.Label),
                ),
              ],
            },
          ],
        });
      });
    this.onLog("ManyToMany Relationships loaded");
    return true;
  }

  async getOptionSets(table: Table): Promise<boolean> {
    if (!this.connection) {
      const error = "No connection available. Please connect first.";
      this.onLog(error, "error");
      throw new Error(error);
    }

    if (table.optionSets.length > 0) {
      return true;
    }

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
    return true;
  }

  async getBooleans(table: Table): Promise<boolean> {
    if (!this.connection) {
      const error = "No connection available. Please connect first.";
      this.onLog(error, "error");
      throw new Error(error);
    }

    if (table.optionSets.filter((os) => os.type === "Boolean").length > 0) {
      return true;
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
    return true;
  }

  async getViews(table: Table): Promise<boolean> {
    if (!this.connection) {
      const error = "No connection available. Please connect first.";
      this.onLog(error, "error");
      throw new Error(error);
    }

    const fetchXml = [
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
    const queryTypeMap: Record<number, string> = {
      0: "Public View",
      2: "Associated View",
      1: "Advanced Search View",
      4: "Quick Find Search View",
      64: "Lookup view",
      2048: "Saved query used for workflow templates and email templates",
      8192: "Outlook offline template",
    };

    await Promise.all(
      (viewsData.value as any[]).map(async (view: any) => {
        const label = await this.getLocLabels("savedqueries", view.savedqueryid, "name");
        const description = await this.getLocLabels("savedqueries", view.savedqueryid, "description");
        table.views.push({
          id: view.savedqueryid,
          name: view.name,
          type: queryTypeMap[view.querytype] || String(view.querytype),
          langProps: [
            { name: "Label", translation: label },
            { name: "Description", translation: description },
          ],
        });
      }),
    );
    return true;
  }

  async getCharts(table: Table): Promise<boolean> {
    if (!this.connection) {
      const error = "No connection available. Please connect first.";
      this.onLog(error, "error");
      throw new Error(error);
    }

    const fetchXml = [
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
        const label = await this.getLocLabels("savedqueryvisualizations", chart.savedqueryvisualizationid, "name");
        const description = await this.getLocLabels(
          "savedqueryvisualizations",
          chart.savedqueryvisualizationid,
          "description",
        );
        table.charts.push({
          id: chart.savedqueryvisualizationid,
          name: chart.name,
          langProps: [
            { name: "Label", translation: label },
            { name: "Description", translation: description },
          ],
        });
      }),
    );
    return true;
  }

  async getForms(table: Table, lang: LanguageDef, base: boolean): Promise<boolean> {
    if (!this.connection) {
      const error = "No connection available. Please connect first.";
      this.onLog(error, "error");
      throw new Error(error);
    }

    const fetchXml = [
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
    const formTypeMap: Record<number, string> = {
      6: "Quick View Form",
      2: "Main",
      7: "Quick Create Form",
    };

    await Promise.all(
      (formsData.value as any[]).map(async (form: any) => {
        const label = await this.getLocLabels("systemforms", form.formid, "name");
        const description = await this.getLocLabels("systemforms", form.formid, "description");
        table.forms.push({
          id: form.formid,
          name: form.name,
          type: formTypeMap[form.type] || String(form.type),
          props: { formXml: form.formxml, uniqueName: form.formidunique, lang: lang.code, base: base },
          langProps: [
            { name: "Label", translation: label },
            { name: "Description", translation: description },
          ],
        });
      }),
    );
    return true;
  }

  async getLocLabels(tableClassName: string, objectId: string, attributeName: string): Promise<LangTranslation[]> {
    if (!this.connection) {
      const error = "No connection available. Please connect first.";
      this.onLog(error, "error");
      throw new Error(error);
    }

    const odata = `RetrieveLocLabels(EntityMoniker=@p1,AttributeName=@p2,IncludeUnpublished=false)?@p1={'@odata.id':'${tableClassName}(${objectId})'}&@p2='${attributeName}'`;
    const result = await this.dvApi.queryData(odata);
    const returnTranslations: LangTranslation[] = [];
    (((result as any).Label?.LocalizedLabels as any[]) || []).forEach((label: any) => {
      returnTranslations.push(new LangTranslation(label.LanguageCode, label.Label));
    });

    return returnTranslations;
  }

  async updateLanguage(lang: string, userId: string): Promise<void> {
    if (!this.connection) {
      const error = "No connection available. Please connect first.";
      this.onLog(error, "error");
      throw new Error(error);
    }

    await this.dvApi.update("usersettingscollection", userId, {
      uilanguageid: lang,
      localeid: lang,
    });

    await new Promise((r) => setTimeout(r, 2000));
  }

  async getUserLanguage(): Promise<{ uiLocale: string; locale: string; userid: string }> {
    if (!this.connection) {
      const error = "No connection available. Please connect first.";
      this.onLog(error, "error");
      throw new Error(error);
    }
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

  async getSiteMaps(solutionId: string): Promise<SecInfo[]> {
    if (!this.connection) {
      const error = "No connection available. Please connect first.";
      this.onLog(error, "error");
      throw new Error(error);
    }

    let siteMaps: SecInfo[] = [];
    if (solutionId !== "") {
      const fetchXml = [
        "<fetch>",
        "  <entity name='solutioncomponent'>",
        "    <attribute name='objectid'/>",
        "    <filter>",
        "      <condition attribute='solutionid' operator='eq' value='",
        solutionId,
        "' uitype='solution'/>",
        "      <filter>",
        "        <condition attribute='componenttype' operator='eq' value='62'/>",
        "      </filter>",
        "    </filter>",
        "  </entity>",
        "</fetch>",
      ].join("");

      const data = await this.dvApi.fetchXmlQuery(fetchXml);
      if (!data.value || (data.value as any[]).length === 0) {
        this.onLog("No site maps found for the selected solution.", "warning");
        return [];
      }
      const siteMapXml = [
        "<fetch>",
        "  <entity name='sitemap'>",
        "    <attribute name='sitemapxml'/>",
        "    <attribute name='sitemapnameunique'/>",
        "    <attribute name='sitemapname'/>",
        "    <filter>",
        "      <condition attribute='sitemapid' operator='in'>",
        data.value.map((sm) => `<value>${sm.objectid}</value>`).join(""),
        "      </condition>",
        "    </filter>",
        "  </entity>",
        "</fetch>",
      ].join("");
      const siteMapData = await this.dvApi.fetchXmlQuery(siteMapXml);
      siteMaps = (siteMapData.value as any[]).map((sm: any) => {
        return {
          id: sm.sitemapid,
          name: sm.sitemapname,
          langProps: [],
          props: { sitemapXml: sm.sitemapxml, uniqueName: sm.sitemapnameunique },
        } as SecInfo;
      });
    } else {
      const url =
        "sitemaps/Microsoft.Dynamics.CRM.RetrieveUnpublishedMultiple()?$select=sitemapxml,sitemapnameunique,sitemapname";
      const data = await this.dvApi.queryData(url);
      siteMaps = (data.value as any[]).map((sm: any) => {
        return {
          id: sm.SitemapId,
          name: sm.SitemapName,
          langProps: [],
          props: { sitemapXml: sm.SitemapXml, uniqueName: sm.SitemapNameUnique },
        } as SecInfo;
      });
    }
    if (siteMaps.length === 0) {
      this.onLog("No site maps found for the selected solution.", "warning");
      return [];
    }
    this.onLog(`Fetched ${siteMaps.length} site maps for solution: ${solutionId}`, "success");
    return siteMaps;
  }

  async getDashboards(solutionId: string): Promise<SecInfo[]> {
    if (!this.connection) {
      const error = "No connection available. Please connect first.";
      this.onLog(error, "error");
      throw new Error(error);
    }

    let dashboards: SecInfo[] = [];
    if (solutionId !== "") {
      const fetchXml = [
        "<fetch>",
        "  <entity name='solutioncomponent'>",
        "    <attribute name='objectid'/>",
        "    <filter>",
        "      <condition attribute='solutionid' operator='eq' value='",
        solutionId,
        "' uitype='solution'/>",
        "      <filter>",
        "        <condition attribute='componenttype' operator='eq' value='60'/>",
        "      </filter>",
        "    </filter>",
        "  </entity>",
        "</fetch>",
      ].join("");

      const data = await this.dvApi.fetchXmlQuery(fetchXml);
      if (!data.value || (data.value as any[]).length === 0) {
        this.onLog("No dashboards found for the selected solution.", "warning");
        return [];
      }
      const dashboardsXml = [
        "<fetch>",
        "  <entity name='systemform'>",
        "    <attribute name='formxml'/>",
        "    <attribute name='name'/>",
        "    <attribute name='formidunique'/>",
        "    <attribute name='formid'/>",
        "    <filter>",
        "      <condition attribute='formid' operator='in'>",
        data.value.map((db) => `<value>${db.objectid}</value>`).join(""),
        "      </condition>",
        "      <condition attribute='type' operator='eq' value='0'/>",
        "    </filter>",
        "  </entity>",
        "</fetch>",
      ].join("");
      const dashboardData = await this.dvApi.fetchXmlQuery(dashboardsXml);
      dashboards = (dashboardData.value as any[]).map((db: any) => {
        return {
          id: db.formid,
          name: db.name,
          langProps: [],
          props: { formXml: db.formxml, uniqueName: db.formidunique, formId: db.formid },
        } as SecInfo;
      });
    } else {
      const dashboardsXml = [
        "<fetch>",
        "  <entity name='systemform'>",
        "    <attribute name='formid'/>",
        "    <attribute name='formidunique'/>",
        "    <attribute name='formxml'/>",
        "    <attribute name='name'/>",
        "    <filter>",
        "      <condition attribute='type' operator='eq' value='0'/>",
        "    </filter>",
        "  </entity>",
        "</fetch>",
      ].join("");
      const dashboardData = await this.dvApi.fetchXmlQuery(dashboardsXml);
      dashboards = (dashboardData.value as any[]).map((db: any) => {
        return {
          id: db.formid,
          name: db.name,
          langProps: [],
          props: { formXml: db.formxml, uniqueName: db.formidunique, formId: db.formid },
        } as SecInfo;
      });
    }
    return dashboards;
  }

  /**
   * Update entity metadata (DisplayName, DisplayCollectionName, Description)
   */
  async updateEntityMetadata(
    entityLogicalName: string,
    displayName: Record<number, string> | null,
    displayCollectionName: Record<number, string> | null,
    description: Record<number, string> | null,
  ): Promise<void> {
    if (!this.connection) {
      const error = "No connection available. Please connect first.";
      this.onLog(error, "error");
      throw new Error(error);
    }
    const body: Record<string, any> = {};
    if (displayName) {
      body.DisplayName = {
        "@odata.type": "Microsoft.Dynamics.CRM.Label",
        LocalizedLabels: Object.entries(displayName).map(([lcid, label]) => ({
          "@odata.type": "Microsoft.Dynamics.CRM.LocalizedLabel",
          Label: label,
          LanguageCode: Number(lcid),
        })),
      };
    }
    if (displayCollectionName) {
      body.DisplayCollectionName = {
        "@odata.type": "Microsoft.Dynamics.CRM.Label",
        LocalizedLabels: Object.entries(displayCollectionName).map(([lcid, label]) => ({
          "@odata.type": "Microsoft.Dynamics.CRM.LocalizedLabel",
          Label: label,
          LanguageCode: Number(lcid),
        })),
      };
    }
    if (description) {
      body.Description = {
        "@odata.type": "Microsoft.Dynamics.CRM.Label",
        LocalizedLabels: Object.entries(description).map(([lcid, label]) => ({
          "@odata.type": "Microsoft.Dynamics.CRM.LocalizedLabel",
          Label: label,
          LanguageCode: Number(lcid),
        })),
      };
    }
    await this.dvApi.updateEntityDefinition(entityLogicalName, body).catch((error) => {
      this.onLog(
        `Failed to update entity metadata for ${entityLogicalName} error: ${(error as Error).message}`,
        "warning",
      );
    });
  }

  /**
   * Update attribute metadata (DisplayName, Description) via Dataverse Web API
   */
  async updateAttributeMetadata(
    entityLogicalName: string,
    attributeLogicalName: string,
    displayName: Record<number, string> | null,
    description: Record<number, string> | null,
  ): Promise<void> {
    if (!this.connection) {
      const error = "No connection available. Please connect first.";
      this.onLog(error, "error");
      throw new Error(error);
    }
    const body: Record<string, any> = { LogicalName: attributeLogicalName };
    if (displayName) {
      body.DisplayName = {
        "@odata.type": "Microsoft.Dynamics.CRM.Label",
        LocalizedLabels: Object.entries(displayName).map(([lcid, label]) => ({
          "@odata.type": "Microsoft.Dynamics.CRM.LocalizedLabel",
          Label: label,
          LanguageCode: Number(lcid),
        })),
      };
    }
    if (description) {
      body.Description = {
        "@odata.type": "Microsoft.Dynamics.CRM.Label",
        LocalizedLabels: Object.entries(description).map(([lcid, label]) => ({
          "@odata.type": "Microsoft.Dynamics.CRM.LocalizedLabel",
          Label: label,
          LanguageCode: Number(lcid),
        })),
      };
    }

    await this.dvApi.updateAttribute(entityLogicalName, attributeLogicalName, body).catch((error) => {
      this.onLog(
        `Failed to update attribute metadata for ${entityLogicalName}.${attributeLogicalName} error: ${(error as Error).message}`,
        "warning",
      );
    });
  }

  /**
   * Update option set value label via SetLocLabels
   */
  async setLocLabels(
    entitySetName: string,
    objectId: string,
    attributeName: string,
    labels: Record<number, string>,
  ): Promise<void> {
    if (!this.connection) {
      const error = "No connection available. Please connect first.";
      this.onLog(error, "error");
      throw new Error(error);
    }
    const locLabels = Object.entries(labels).map(([lcid, label]) => ({
      "@odata.type": "Microsoft.Dynamics.CRM.LocalizedLabel",
      Label: label,
      LanguageCode: Number(lcid),
    }));

    await this.dvApi.execute({
      operationName: "SetLocLabels",
      operationType: "action",
      parameters: {
        EntityMoniker: {
          "@odata.type": "Microsoft.Dynamics.CRM.crmbaseentity",
          "@odata.id": `${entitySetName}(${objectId})`,
        },
        AttributeName: attributeName,
        Labels: locLabels,
      },
    });
  }

  /**
   * Update an option value label for local/global option set
   */
  async updateOptionValue(
    entityLogicalName: string | null,
    attributeLogicalName: string | null,
    optionSetName: string | null,
    value: number,
    labels: Record<number, string>,
    isDescription: boolean = false,
  ): Promise<void> {
    if (!this.connection) {
      const error = "No connection available. Please connect first.";
      this.onLog(error, "error");
      throw new Error(error);
    }
    const locLabels = Object.entries(labels).map(([lcid, label]) => ({
      "@odata.type": "Microsoft.Dynamics.CRM.LocalizedLabel",
      Label: label,
      LanguageCode: Number(lcid),
    }));

    const params: Record<string, any> = {
      Value: value,
      MergeLabels: true,
    };

    if (isDescription) {
      params.Description = {
        "@odata.type": "Microsoft.Dynamics.CRM.Label",
        LocalizedLabels: locLabels,
      };
    } else {
      params.Label = {
        "@odata.type": "Microsoft.Dynamics.CRM.Label",
        LocalizedLabels: locLabels,
      };
    }

    if (optionSetName) {
      params.OptionSetName = optionSetName;
    } else {
      params.EntityLogicalName = entityLogicalName;
      params.AttributeLogicalName = attributeLogicalName;
    }

    await this.dvApi.updateOptionValue(params).catch((error) => {
      this.onLog(
        `Failed to update option value label for ${entityLogicalName}.${attributeLogicalName} value: ${value} error: ${(error as Error).message}`,
        "warning",
      );
    });
  }

  /**
   * Update relationship metadata
   */
  async updateRelationshipLabel(
    schemaName: string,
    labels: Record<number, string>,
    relationType: string,
  ): Promise<void> {
    if (!this.connection) {
      const error = "No connection available. Please connect first.";
      this.onLog(error, "error");
      throw new Error(error);
    }
    this.onLog("Updating relationship label", "info");
    const relationship = await this.dvApi.queryData(`RelationshipDefinitions?$filter=SchemaName eq '${schemaName}'`);

    if (!relationship?.value || !Array.isArray(relationship.value) || relationship.value.length === 0) {
      const error = `Relationship with SchemaName '${schemaName}' not found`;
      this.onLog(error, "error");
      throw new Error(error);
    }

    const relTypeName =
      relationType === "ManyToManyRelationship"
        ? "Microsoft.Dynamics.CRM.ManyToManyRelationshipMetadata"
        : "Microsoft.Dynamics.CRM.OneToManyRelationshipMetadata";

    const rel = relationship.value[0];

    // Select the appropriate associated menu configuration object(s) based on relationship type
    // Note: Using 'any' type here because Dataverse API relationship metadata structure is complex
    // and varies between ManyToMany and OneToMany relationships
    const menuConfigs: any[] = [];
    if (relationType === "ManyToManyRelationship") {
      if (rel?.Entity1AssociatedMenuConfiguration) {
        menuConfigs.push(rel.Entity1AssociatedMenuConfiguration);
      }
      if (rel?.Entity2AssociatedMenuConfiguration) {
        menuConfigs.push(rel.Entity2AssociatedMenuConfiguration);
      }
    } else if (rel?.AssociatedMenuConfiguration) {
      menuConfigs.push(rel.AssociatedMenuConfiguration);
    }

    // Update labels with the provided translations on all relevant menu configurations
    menuConfigs.forEach((amc) => {
      if (amc?.Label?.LocalizedLabels) {
        // Update existing labels or add new ones
        Object.entries(labels).forEach(([lcid, labelText]) => {
          const languageCode = Number(lcid);
          const existingLabel = amc.Label.LocalizedLabels.find((lbl: any) => lbl.LanguageCode === languageCode);

          if (existingLabel) {
            existingLabel.Label = labelText;
          } else {
            amc.Label.LocalizedLabels.push({
              "@odata.type": "Microsoft.Dynamics.CRM.LocalizedLabel",
              Label: labelText,
              LanguageCode: languageCode,
            });
          }
        });
      }
    });

    await this.dvApi.updateRelationship(relTypeName, rel).catch((error) => {
      this.onLog(
        `Failed to update relationship metadata for ${schemaName} error: ${(error as Error).message}`,
        "warning",
      );
    });
  }

  /**
   * Update a form record (for form name/description or form XML content)
   */
  async updateForm(formId: string, updates: Record<string, unknown>): Promise<void> {
    if (!this.connection) {
      const error = "No connection available. Please connect first.";
      this.onLog(error, "error");
      throw new Error(error);
    }
    await this.dvApi.update("systemform", formId, updates);
  }

  /**
   * Update a view record
   */
  async updateView(viewId: string, updates: Record<string, unknown>): Promise<void> {
    if (!this.connection) {
      const error = "No connection available. Please connect first.";
      this.onLog(error, "error");
      throw new Error(error);
    }
    await this.dvApi.update("savedquery", viewId, updates);
  }

  /**
   * Update a chart record
   */
  async updateChart(chartId: string, updates: Record<string, unknown>): Promise<void> {
    if (!this.connection) {
      const error = "No connection available. Please connect first.";
      this.onLog(error, "error");
      throw new Error(error);
    }
    await this.dvApi.update("savedqueryvisualization", chartId, updates);
  }

  /**
   * Update a sitemap record
   */
  async updateSiteMap(siteMapId: string, sitemapXml: string): Promise<void> {
    if (!this.connection) {
      const error = "No connection available. Please connect first.";
      this.onLog(error, "error");
      throw new Error(error);
    }
    await this.dvApi.update("sitemap", siteMapId, { sitemapxml: sitemapXml });
  }

  /**
   * Retrieve a form by ID
   */
  async getFormById(formId: string): Promise<Record<string, unknown>> {
    if (!this.connection) {
      const error = "No connection available. Please connect first.";
      this.onLog(error, "error");
      throw new Error(error);
    }
    return await this.dvApi.retrieve("systemform", formId, ["formid", "formxml", "name", "formidunique", "type"]);
  }

  /**
   * Retrieve a sitemap by ID
   */
  async getSiteMapById(siteMapId: string): Promise<Record<string, unknown>> {
    if (!this.connection) {
      const error = "No connection available. Please connect first.";
      this.onLog(error, "error");
      throw new Error(error);
    }
    return await this.dvApi.retrieve("sitemap", siteMapId, [
      "sitemapid",
      "sitemapxml",
      "sitemapname",
      "sitemapnameunique",
    ]);
  }

  /**
   * Publish all customizations
   */
  async publishAllCustomizations(): Promise<void> {
    if (!this.connection) {
      const error = "No connection available. Please connect first.";
      this.onLog(error, "error");
      throw new Error(error);
    }
    await this.dvApi.publishCustomizations();
  }
}
