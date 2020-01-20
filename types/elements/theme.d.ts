export default class Theme {
    fontFamily?: string;
    colorScheme?: any;
    constructor(fontFamily?: string, colorScheme?: any);
    render(): string;
}
