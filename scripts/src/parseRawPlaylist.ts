import yaml from "yaml";
import type { RawPlaylist } from "./RawPlaylist";

export function parseRawPlaylist(data: string): RawPlaylist {
    const items = yaml.parse(data) as RawPlaylist;
    return items;
}
