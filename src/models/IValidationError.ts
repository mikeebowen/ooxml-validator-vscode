/* eslint-disable @typescript-eslint/naming-convention */
export interface IValidationError {
  Description?: string;
  Path?: {
    NamespacesDefinitions?: string[];
    Namespaces: any;
    XPath?: string;
    PartUri?: string;
  };
  Id?: string;
  ErrorType?: number;
}