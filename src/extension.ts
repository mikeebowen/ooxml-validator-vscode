import { commands, ExtensionContext } from 'vscode';
import OOXMLValidator from './ooxm-validator';

export function activate(context: ExtensionContext) {
  const validateOOXML = commands.registerCommand('ooxml-validator-vscode.validateOOXML', file => OOXMLValidator.validate(file));

  context.subscriptions.push(validateOOXML);
}

export function deactivate() {}
