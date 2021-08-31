import { commands, ExtensionContext, Uri, extensions } from 'vscode';
import OOXMLValidator from './ooxml-validator';

export async function activate(context: ExtensionContext) {
  const runtimeExtension = extensions.getExtension('ms-dotnettools.vscode-dotnet-runtime');
  await runtimeExtension?.activate();

  const validateOOXML = commands.registerCommand('ooxml-validator-vscode.validateOOXML', (file: Uri) => OOXMLValidator.validate(file));

  context.subscriptions.push(validateOOXML);
}

export function deactivate() {}
