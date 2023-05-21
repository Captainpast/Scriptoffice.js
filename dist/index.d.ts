import { OfficeDocumentTypes } from "./OfficeDocument";
export declare function create<T extends keyof OfficeDocumentTypes>(options: T | {
    type: T;
}): OfficeDocumentTypes[T];
/**@deprecated WIP*/
export declare function load(src: string): void;
