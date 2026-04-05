/**
 * Shared XML parser/builder options matching OOXML conventions.
 */
export declare const xmlParserOptions: {
    ignoreAttributes: boolean;
    attributeNamePrefix: string;
    removeNSPrefix: boolean;
    parseTagValue: boolean;
    parseAttributeValue: boolean;
    trimValues: boolean;
    isArray: (name: string) => boolean;
};
export declare const xmlBuilderOptions: {
    ignoreAttributes: boolean;
    attributeNamePrefix: string;
    format: boolean;
    suppressEmptyNode: boolean;
    suppressBooleanAttributes: boolean;
};
