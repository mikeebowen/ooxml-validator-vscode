import { commands, ExtensionContext, Uri } from 'vscode';
import OOXMLValidator from './ooxml-validator';

export function activate(context: ExtensionContext) {
  const validateOOXML = commands.registerCommand('ooxml-validator-vscode.validateOOXML', (file: Uri) => OOXMLValidator.validate(file));

  context.subscriptions.push(validateOOXML);
}

export function deactivate() {}
