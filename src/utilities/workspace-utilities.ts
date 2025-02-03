import { FileSystemError, Uri, workspace } from 'vscode';
import logger from './logger';

export class WorkspaceUtilities {
  /**
   * Creates a directory.
   *
   * @param {string} directoryPath The path to the directory to create.
   */
  static async createDirectory(directoryPath: string): Promise<void> {
    logger.trace(`Creating directory '${directoryPath}'`);
    await workspace.fs.createDirectory(Uri.file(directoryPath));
  }

  /**
     * Writes data to a file.
     *
     * @param filePath The path of the file to update.
     * @param data The data write to the file.
     * @returns A promise resolving to whether or not the file was written to successfully.
     */
  static async writeFile(filePath: string, data: Uint8Array): Promise<boolean> {
    try {
      logger.trace(`Writing file '${filePath}'`);
      await workspace.fs.writeFile(Uri.file(filePath), data);
    } catch (err) {
      if ((err as FileSystemError)?.code?.toLowerCase() === 'unknown'
      || (err as FileSystemError)?.message.toLowerCase().includes('ebusy')) {
        return false;
      }

      logger.error(`Error writing file '${filePath}'`);
      throw err;
    }

    return true;
  }

  /**
   * Gets a configuration value
   *
   * @param {string} section The configuration section.
   * @param {string} name The configuration name.
   * @param {T} defaultValue The type to return.
   * @returns {T} The configuration value or undefined.
   */
  static getConfigurationValue<T>(section: string, name: string): T | undefined {
    return workspace
      .getConfiguration(section)
      .get(name);
  }

}