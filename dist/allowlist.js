// Caller allowlist helpers for provider-normalized caller ids (phone or Teams aadId).
/** Normalize a phone number to digits only. */
export function normalizePhoneNumber(input) {
    if (!input) {
        return "";
    }
    return input.replace(/\D/g, "");
}
/**
 * Return true when the caller matches an allowlist entry — by phone number
 * (digits only) OR by exact caller id, case-insensitive. The id match lets
 * providers whose caller id is not a phone number (e.g. the Microsoft Teams
 * `aadId`, an AAD object id) use the allowlist without being phone-normalized.
 * `from` may be a raw caller id or an already-normalized phone string.
 */
export function isAllowlistedCaller(from, allowFrom) {
    const raw = from?.trim();
    if (!raw) {
        return false;
    }
    const idFrom = raw.toLowerCase();
    const normalizedFrom = normalizePhoneNumber(raw);
    return (allowFrom ?? []).some((entry) => {
        const trimmed = entry.trim();
        if (!trimmed) {
            return false;
        }
        // Exact caller-id match (e.g. a Teams AAD object id), case-insensitive.
        if (trimmed.toLowerCase() === idFrom) {
            return true;
        }
        // Phone-number match (digits only).
        const normalizedAllow = normalizePhoneNumber(trimmed);
        return normalizedAllow !== "" && normalizedFrom !== "" && normalizedAllow === normalizedFrom;
    });
}
/**
 * Inbound-policy decision. `from` is the caller id — a phone number, or a Teams
 * `aadId` for msteams. `isAllowlistedCaller` matches either form.
 * "disabled"/unset → reject (defensive).
 */
export function isInboundCallAllowed(inboundPolicy, allowFrom, from) {
    switch (inboundPolicy) {
        case "open":
            return true;
        case "allowlist":
        case "pairing":
            return isAllowlistedCaller(from, allowFrom);
        default:
            return false;
    }
}
