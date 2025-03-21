import yaml from "yaml";
import type { RawSample } from "./RawSample";

export function parseRawSample(data: string): RawSample {
    const rawSample = yaml.parse(data) as RawSample;
    return rawSample;
}
