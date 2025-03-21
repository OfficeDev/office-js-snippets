/**
 * YAML
 */
export interface RawSample {
    name: string;
    description: string;
    script: {
        content: string;
        language: string;
    };
    template: {
        content: string;
        language: string;
    };
    style: {
        content: string;
        language: string;
    };
    libraries: string;
}
