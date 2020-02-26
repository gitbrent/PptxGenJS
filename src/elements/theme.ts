import themeXML, { colorSchemeXML } from '../templates/theme'

export default class Theme {
    fontFamily?: string
    titleFontFamily?: string
    colorScheme?

    constructor(fontFamily?: string, colorScheme?, titleFontFamily?) {
        this.fontFamily = fontFamily
        this.titleFontFamily = titleFontFamily
        this.colorScheme = colorScheme
    }

    render() {
        return themeXML(
            this.fontFamily,
            this.titleFontFamily,
            this.colorScheme && colorSchemeXML(this.colorScheme)
        )
    }
}
