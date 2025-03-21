export function transformHtml(data: string): string {
    return data.replace(/\n\n/g, "\n").trim();
}
