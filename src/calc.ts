export const CALC_EXPR = /^calc\((.+)\)$/

type Value = string | number

const processCalcArray = (values: Value[], calcExpr: string): number => {
    values.forEach((v, index) => {
        if (v === '-') values[index + 1] = -values[index + 1]
    })
    values = values.filter(v => v !== '-' && v !== '+')

    values.forEach((v, index) => {
        if (v === '/') {
            const nominator = values[index - 1]
            const denominator = values[index + 1]

            if (
                typeof nominator !== 'number' ||
                typeof denominator !== 'number'
            ) {
                console.warn(`Bad calc expression (division) "${calcExpr}"`)
                values[index + 1] = 0
                values[index - 1] = 0
            } else {
                values[index + 1] = nominator / denominator
                values[index - 1] = 0
            }
        }
        if (v === '*') {
            const firstVal = values[index - 1]
            const secondVal = values[index + 1]

            if (typeof firstVal !== 'number' || typeof secondVal !== 'number') {
                console.warn(
                    `Bad calc expression (multiplication) "${calcExpr}"`
                )
                values[index + 1] = 0
                values[index - 1] = 0
            } else {
                values[index + 1] = firstVal * secondVal
                values[index - 1] = 0
            }
        }
    })
    const result = values
        .filter(v => v !== '*' && v !== '/')
        .map(n => Number(n))
        .reduce((x, y) => x + y, 0)
    return result
}

const processParentheses = (valueArray: Value[], calcExpr: string): number => {
    const withParenthesesDone = []
    let parenLevel = 0
    let currentParen = []

    for (let value of valueArray) {
        if (value === ')' && parenLevel === 0) {
            console.warn(`parenthesis mismatch in ${calcExpr}`)
            return 0
        } else if (value === ')' && parenLevel === 1) {
            withParenthesesDone.push(processParentheses(currentParen, calcExpr))
            currentParen = []
            parenLevel = 0
        } else if (value === ')' && parenLevel > 1) {
            parenLevel -= 1
        } else if (parenLevel === 0 && value !== '(' && value !== ')') {
            withParenthesesDone.push(value)
        } else if (parenLevel > 0 && value !== '(' && value !== ')') {
            currentParen.push(value)
        } else if (value === '(') {
            parenLevel += 1
        }
    }

    if (parenLevel > 0) {
        console.warn(`parenthesis mismatch in "${calcExpr}"`)
        return 0
    }

    return processCalcArray(withParenthesesDone, calcExpr)
}

export default (calcExpr: string, parseValue: (n: string) => number) => {
    if (!CALC_EXPR.test(calcExpr)) {
        console.warn(`Not a valid calc expression "${calcExpr}"`)
        return 0
    }
    const [, calc] = calcExpr.match(CALC_EXPR)
    const values: string[] = calc
        .replace(/calc/g, '')
        .replace(/\(/g, ' ( ')
        .replace(/\)/g, ' ) ')
        .replace(/\+/g, ' + ')
        .replace(/-/g, ' - ')
        .replace(/\*/g, ' * ')
        .replace(/\//g, ' / ')
        .split(/\s/)
        .filter(v => v)

    const parsedValues: Value[] = values.map(v => {
        if (
            v === '+' ||
            v === '-' ||
            v === '/' ||
            v === '*' ||
            v === '(' ||
            v === ')'
        )
            return v

        const scalar = Number(v)
        if (!Number.isNaN(scalar)) return scalar
        return parseValue(v)
    })

    return processParentheses(parsedValues, calcExpr)
}
