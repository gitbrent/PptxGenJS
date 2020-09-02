---
id: types
title: Type Interfaces
---

PptxGenJS Type Interfaces.

## Position Props (`PositionProps`)

| Name | Type   | Default | Description            | Possible Values                              |
| :--- | :----- | :------ | :--------------------- | :------------------------------------------- |
| `x`  | number | `1.0`   | hor location (inches)  | 0-n                                          |
| `x`  | string |         | hor location (percent) | 'n%'. (Ex: `{x:'50%'}` middle of the Slide)  |
| `y`  | number | `1.0`   | ver location (inches)  | 0-n                                          |
| `y`  | string |         | ver location (percent) | 'n%'. (Ex: `{y:'50%'}` middle of the Slide)  |
| `w`  | number | `1.0`   | width (inches)         | 0-n                                          |
| `w`  | string |         | width (percent)        | 'n%'. (Ex: `{w:'50%'}` 50% the Slide width)  |
| `h`  | number | `1.0`   | height (inches)        | 0-n                                          |
| `h`  | string |         | height (percent)       | 'n%'. (Ex: `{h:'50%'}` 50% the Slide height) |

## Data/Path Props (`DataOrPathProps`)

| Name   | Type   | Description         | Possible Values                                                            |
| :----- | :----- | :------------------ | :------------------------------------------------------------------------- |
| `data` | string | image data (base64) | base64-encoded image string. (either `data` or `path` is required)         |
| `path` | string | image path          | Same as used in an (img src="") tag. (either `data` or `path` is required) |

## Hyperlink Props (`HyperlinkProps`)

| Name      | Type   | Description           | Possible Values                |
| :-------- | :----- | :-------------------- | :----------------------------- |
| `slide`   | number | link to a given slide | Ex: `2`                        |
| `tooltip` | string | link tooltip text     | Ex: `Click to visit home page` |
| `url`     | string | target URL            | Ex: `https://wikipedia.org`    |

## Shadow Props (`ShadowProps`)

| Name      | Type   | Default  | Description            | Possible Values          |
| :-------- | :----- | :------- | :--------------------- | :----------------------- |
| `type`    | string | `none`   | shadow type            | `outer`, `inner`, `none` |
| `angle`   | number | `0`      | blue degrees           | `0`-`359`                |
| `blur`    | number | `0`      | blur range (points)    | `0`-`100`                |
| `color`   | string | `000000` | color                  | hex color code           |
| `offset`  | number | `0`      | shadow offset (points) | `0`-`200`                |
| `opacity` | number | `0`      | opacity percentage     | `0.0`-`1.0`              |

## Shape Fill Props (`ShapeFillProps`)

| Name           | Type   | Default  | Description            | Possible Values                                                       |
| :------------- | :----- | :------- | :--------------------- | :-------------------------------------------------------------------- |
| `color`        | string | `000000` | `ShapeFillProps` color | hex color or [scheme color](/PptxGenJS/docs/shapes-and-schemes.html). |
| `transparency` | number | `0`      | `ShapeFillProps` trans | transparency percentage: 0-100                                        |

## Shape Line Props (`ShapeLineProps`)

| Name             | Type   | Default | Description         | Possible Values                                                                                           |
| :--------------- | :----- | :------ | :------------------ | :-------------------------------------------------------------------------------------------------------- |
| `beginArrowType` | string |         | line ending         | `arrow`, `diamond`, `oval`, `stealth`, `triangle` or `none`                                               |
| `color`          | string |         | line color          | hex color code or [scheme color constant](/PptxGenJS/docs/shapes-and-schemes.html). Ex: `{line:'0088CC'}` |
| `dashType`       | string | `solid` | line dash style     | `dash`, `dashDot`, `lgDash`, `lgDashDot`, `lgDashDotDot`, `solid`, `sysDash` or `sysDot`                  |
| `endArrowType`   | string |         | line heading        | `arrow`, `diamond`, `oval`, `stealth`, `triangle` or `none`                                               |
| `transparency`   | number | `0`     | line transparency   | Percentage: 0-100                                                                                         |
| `width`          | number | `1`     | line width (points) | 1-256. Ex: `{ width:4 }`                                                                                  |
