import { window } from 'vscode';
import logger from './logger';

export class WindowUtilities {
  /**
   * Handles an error.
   *
   * @param {unknown} err The error.
   */
  static async showError(err: unknown, modal: boolean = false): Promise<void> {
    let msg = 'unknown error';

    if (typeof err === 'string') {
      msg = err;
    } else if (err instanceof Error) {
      msg = err.message;
    }

    logger.error(msg);
    await window.showErrorMessage(msg, { modal });
  }
  /**
   * Displays a warning message.
   *
   * @param {string} message The warning message.
   */
  static async showWarning(message: string, detail: string | undefined = undefined,  modal: boolean = false): Promise<void> {
    logger.warn(message);
    await window.showWarningMessage(message, { modal, detail });
  }
}