import { spawnSync } from 'child_process';
import logger from './logger';

export class ExtensionUtilities {
  /**
   * Checks if a file is the .NET runtime.
   *
   * @param {string} filePath The file path to check.
   * @returns {Promise<boolean>} A promise that resolves to whether or not the file is the .NET runtime.
   */
  static async isDotNetRuntime(filePath: string): Promise<boolean> {
    try {
      const result = spawnSync(filePath, ['--info']);
      const output = result.stdout?.toString()?.trim();

      return output?.includes('.NET');
    } catch (err) {
      logger.error(`Error checking if file '${filePath}' is the .NET runtime: ${err}`);

      return false;
    }
  }
}