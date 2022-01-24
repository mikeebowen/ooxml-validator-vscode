import { IHeader } from './IHeader';

export class Header {
  id: string;
  title: string;
  constructor(options: IHeader) {
    this.id = options.id;
    this.title = options.title;
  }
}