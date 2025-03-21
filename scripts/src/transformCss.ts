export function transformCss(data: string): string {
    const body = `body {
    background-color: white;
}`;
    const clean = `${body}

${data}`;

    return clean;
}
