import { makeAutoObservable } from "mobx";

export class Solution {
  solutionId: string;
  name: string;
  uniqueName: string;
  description?: string;
  version?: string;
  isManaged?: boolean;
  subcomponents?: boolean;

  attributes: { attributeName: string; attributeValue: string }[];

  constructor() {
    this.solutionId = "";
    this.name = "";
    this.uniqueName = "";
    this.attributes = [];
    makeAutoObservable(this);
  }
}
