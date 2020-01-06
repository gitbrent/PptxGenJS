import themeXML, { colorSchemeXML } from '../templates/theme'

export default class Theme {
    fontFamily?: string
    colorScheme?

    constructor(fontFamily?: string, colorScheme?) {
        this.fontFamily = fontFamily
        this.colorScheme = colorScheme
    }

    render() {
        return themeXML(
            this.fontFamily,
            this.colorScheme && colorSchemeXML(this.colorScheme)
        )
    }
}
