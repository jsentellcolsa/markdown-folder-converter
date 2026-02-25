// Type declarations for packages that ship without their own types.

declare module 'mammoth' {
  interface Message {
    type: string;
    message: string;
  }
  interface ConversionResult {
    value: string;
    messages: Message[];
  }
  interface ConversionInput {
    path?: string;
    buffer?: Buffer;
    arrayBuffer?: ArrayBuffer;
  }
  export function convertToHtml(input: ConversionInput): Promise<ConversionResult>;
}

declare module 'turndown-plugin-gfm' {
  import TurndownService from 'turndown';
  export function gfm(service: TurndownService): void;
  export function tables(service: TurndownService): void;
  export function strikethrough(service: TurndownService): void;
  export function taskListItems(service: TurndownService): void;
}
