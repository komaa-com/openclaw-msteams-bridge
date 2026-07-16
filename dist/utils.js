import os from "node:os";
import path from "node:path";
export function resolveUserPath(input) {
    const trimmed = input.trim();
    if (!trimmed) {
        return trimmed;
    }
    if (trimmed.startsWith("~")) {
        const expanded = trimmed.replace(/^~(?=$|[\\/])/, os.homedir());
        return path.resolve(expanded);
    }
    return path.resolve(trimmed);
}
export function readArgText(args, key) {
    if (!args || typeof args !== "object" || Array.isArray(args)) {
        return undefined;
    }
    const value = args[key];
    return typeof value === "string" && value.trim() ? value.trim() : undefined;
}
