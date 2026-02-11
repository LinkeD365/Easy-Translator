import { makeAutoObservable } from "mobx";
import { Solution } from "./solution";

export class ViewModel {
  exporting: boolean = false;
  message: string = "";
  exportpercentage: number = 0;
  options: Options = new Options();
  selectedLanguage?: LanguageDef;
  allLanguages: LanguageDef[] = [];
  selectedTables: Table[] = [];
  uiLocale: string = "";
  locale: string = "";
  userId: string = "";
  solution?: Solution;
  siteMaps: SecInfo[] = [];
  siteAreas: SecInfo[] = [];
  siteGroups: SecInfo[] = [];
  siteSubAreas: SecInfo[] = [];
  dashboards: SecInfo[] = [];
  dashboardTabs: SecInfo[] = [];
  dashboardSections: SecInfo[] = [];
  dashboardFields: SecInfo[] = [];

  constructor() {
    makeAutoObservable(this);
  }
}

export class Table {
  label: string;
  setName: string;
  logicalName: string;
  id: string;
  langProps: LangProp[] = [];
  fields: SecInfo[] = [];
  relationships: SecInfo[] = [];
  optionSets: OptionSetDef[] = [];
  views: SecInfo[] = [];
  charts: SecInfo[] = [];
  forms: SecInfo[] = [];
  tabs: SecInfo[] = [];
  sections: SecInfo[] = [];
  formFields: SecInfo[] = [];
  code: string;

  constructor(label: string, setName: string, logicalName: string, id: string, code: string) {
    this.label = label;
    this.setName = setName;
    this.logicalName = logicalName;
    this.id = id;
    this.code = code;
    makeAutoObservable(this);
  }
}

export class Options {
  optionSets: boolean = true;
  siteMaps: boolean = true;
  dashboards: boolean = true;

  table: boolean = true;
  fields: boolean = true;
  localOptionSets: boolean = true;
  globalOptionSets: boolean = true;
  booleanOptions: boolean = true;
  views: boolean = true;
  charts: boolean = true;
  forms: boolean = true;
  formTabs: boolean = true;
  formSections: boolean = true;
  formFields: boolean = true;
  relationships: boolean = true;
  labelOptions: LabelOptions = LabelOptions.both;

  exportAllLanguages: boolean = true;

  constructor() {
    makeAutoObservable(this);
  }
  labels() {
    return this.labelOptions !== LabelOptions.descriptions;
  }
  descriptions() {
    return this.labelOptions !== LabelOptions.names;
  }
}

export enum LabelOptions {
  both,
  names,
  descriptions,
}

export class LanguageDef {
  code: string = "";
  name: string = "";

  constructor() {
    makeAutoObservable(this);
  }
}

export class SecInfo {
  id: string;
  name: string;
  type?: string;
  props?: Record<string, any>;
  langProps: LangProp[];

  constructor(id: string, name: string, langProps: LangProp[]) {
    this.id = id;
    this.name = name;
    this.langProps = langProps;
  }
}

export class OptionSetDef extends SecInfo {
  isGlobal: boolean;
  attributeLogicalName?: string;
  optionValue?: number;

  constructor(id: string, name: string, langProps: LangProp[], isGlobal: boolean) {
    super(id, name, langProps);
    this.isGlobal = isGlobal;
  }
}

export class LangProp {
  name: string;
  translation: LangTranslation[];

  constructor(name: string, translation: LangTranslation[]) {
    this.name = name;
    this.translation = translation;
  }
}

export class LangTranslation {
  code: string;
  translation: string;

  constructor(code: string, translation: string) {
    this.code = code;
    this.translation = translation;
  }
}
