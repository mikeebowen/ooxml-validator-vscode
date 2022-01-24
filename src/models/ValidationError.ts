/* eslint-disable @typescript-eslint/naming-convention */
import { IValidationError } from './IValidationError';

export class ValidationError {
  Description?: string;
  NamespacesDefinitions?: string[] | undefined;
  Namespaces: any;
  XPath?: string;
  PartUri?: string;
  Id?: string;
  ErrorType?: number;

  constructor(options: IValidationError) {
    this.Id = options.Id;
    this.Description = options.Description;
    this.Namespaces = options.Path?.Namespaces;
    this.NamespacesDefinitions = options.Path?.NamespacesDefinitions;
    this.XPath = options.Path?.XPath;
    this.PartUri = options.Path?.PartUri;
    this.ErrorType = options.ErrorType;
  }
}