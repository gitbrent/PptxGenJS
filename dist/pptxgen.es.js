/* PptxGenJS 3.0.0-beta.7 @ 2019-11-30T17:36:43.138Z */
import * as JSZip from 'jszip';

/**
 * PptxGenJS Enums
 * NOTE: `enum` wont work for objects, so use `Object.freeze`
 */
// CONST
var EMU = 914400; // One (1) inch (OfficeXML measures in EMU (English Metric Units))
var ONEPT = 12700; // One (1) point (pt)
var CRLF = '\r\n'; // AKA: Chr(13) & Chr(10)
var LAYOUT_IDX_SERIES_BASE = 2147483649;
var REGEX_HEX_COLOR = /^[0-9a-fA-F]{6}$/;
var LINEH_MODIFIER = 1.67; // AKA: Golden Ratio Typography
var DEF_CELL_BORDER = { color: '666666' };
var DEF_CELL_MARGIN_PT = [3, 3, 3, 3]; // TRBL-style
var DEF_CHART_GRIDLINE = { color: '888888', style: 'solid', size: 1 };
var DEF_FONT_COLOR = '000000';
var DEF_FONT_SIZE = 12;
var DEF_FONT_TITLE_SIZE = 18;
var DEF_PRES_LAYOUT = 'LAYOUT_16x9';
var DEF_PRES_LAYOUT_NAME = 'DEFAULT';
var DEF_SLIDE_MARGIN_IN = [0.5, 0.5, 0.5, 0.5]; // TRBL-style
var DEF_SHAPE_SHADOW = { type: 'outer', blur: 3, offset: 23000 / 12700, angle: 90, color: '000000', opacity: 0.35, rotateWithShape: true };
var AXIS_ID_VALUE_PRIMARY = '2094734552';
var AXIS_ID_VALUE_SECONDARY = '2094734553';
var AXIS_ID_CATEGORY_PRIMARY = '2094734554';
var AXIS_ID_CATEGORY_SECONDARY = '2094734555';
var AXIS_ID_SERIES_PRIMARY = '2094734556';
var LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('');
var BARCHART_COLORS = [
    'C0504D',
    '4F81BD',
    '9BBB59',
    '8064A2',
    '4BACC6',
    'F79646',
    '628FC6',
    'C86360',
    'C0504D',
    '4F81BD',
    '9BBB59',
    '8064A2',
    '4BACC6',
    'F79646',
    '628FC6',
    'C86360',
];
var PIECHART_COLORS = [
    '5DA5DA',
    'FAA43A',
    '60BD68',
    'F17CB0',
    'B2912F',
    'B276B2',
    'DECF3F',
    'F15854',
    'A7A7A7',
    '5DA5DA',
    'FAA43A',
    '60BD68',
    'F17CB0',
    'B2912F',
    'B276B2',
    'DECF3F',
    'F15854',
    'A7A7A7',
];
var TEXT_HALIGN;
(function (TEXT_HALIGN) {
    TEXT_HALIGN["left"] = "left";
    TEXT_HALIGN["center"] = "center";
    TEXT_HALIGN["right"] = "right";
    TEXT_HALIGN["justify"] = "justify";
})(TEXT_HALIGN || (TEXT_HALIGN = {}));
var TEXT_VALIGN;
(function (TEXT_VALIGN) {
    TEXT_VALIGN["b"] = "b";
    TEXT_VALIGN["ctr"] = "ctr";
    TEXT_VALIGN["t"] = "t";
})(TEXT_VALIGN || (TEXT_VALIGN = {}));
var SLDNUMFLDID = '{F7021451-1387-4CA6-816F-3879F97B5CBC}';
// ENUM
var SCHEME_COLOR_NAMES;
(function (SCHEME_COLOR_NAMES) {
    SCHEME_COLOR_NAMES["TEXT1"] = "tx1";
    SCHEME_COLOR_NAMES["TEXT2"] = "tx2";
    SCHEME_COLOR_NAMES["BACKGROUND1"] = "bg1";
    SCHEME_COLOR_NAMES["BACKGROUND2"] = "bg2";
    SCHEME_COLOR_NAMES["ACCENT1"] = "accent1";
    SCHEME_COLOR_NAMES["ACCENT2"] = "accent2";
    SCHEME_COLOR_NAMES["ACCENT3"] = "accent3";
    SCHEME_COLOR_NAMES["ACCENT4"] = "accent4";
    SCHEME_COLOR_NAMES["ACCENT5"] = "accent5";
    SCHEME_COLOR_NAMES["ACCENT6"] = "accent6";
})(SCHEME_COLOR_NAMES || (SCHEME_COLOR_NAMES = {}));
var MASTER_OBJECTS;
(function (MASTER_OBJECTS) {
    MASTER_OBJECTS["chart"] = "chart";
    MASTER_OBJECTS["image"] = "image";
    MASTER_OBJECTS["line"] = "line";
    MASTER_OBJECTS["rect"] = "rect";
    MASTER_OBJECTS["text"] = "text";
    MASTER_OBJECTS["placeholder"] = "placeholder";
})(MASTER_OBJECTS || (MASTER_OBJECTS = {}));
var SLIDE_OBJECT_TYPES;
(function (SLIDE_OBJECT_TYPES) {
    SLIDE_OBJECT_TYPES["chart"] = "chart";
    SLIDE_OBJECT_TYPES["hyperlink"] = "hyperlink";
    SLIDE_OBJECT_TYPES["image"] = "image";
    SLIDE_OBJECT_TYPES["media"] = "media";
    SLIDE_OBJECT_TYPES["online"] = "online";
    SLIDE_OBJECT_TYPES["placeholder"] = "placeholder";
    SLIDE_OBJECT_TYPES["table"] = "table";
    SLIDE_OBJECT_TYPES["tablecell"] = "tablecell";
    SLIDE_OBJECT_TYPES["text"] = "text";
    SLIDE_OBJECT_TYPES["notes"] = "notes";
})(SLIDE_OBJECT_TYPES || (SLIDE_OBJECT_TYPES = {}));
var PLACEHOLDER_TYPES;
(function (PLACEHOLDER_TYPES) {
    PLACEHOLDER_TYPES["title"] = "title";
    PLACEHOLDER_TYPES["body"] = "body";
    PLACEHOLDER_TYPES["image"] = "pic";
    PLACEHOLDER_TYPES["chart"] = "chart";
    PLACEHOLDER_TYPES["table"] = "tbl";
    PLACEHOLDER_TYPES["media"] = "media";
})(PLACEHOLDER_TYPES || (PLACEHOLDER_TYPES = {}));
var CHART_TYPES;
(function (CHART_TYPES) {
    CHART_TYPES["AREA"] = "area";
    CHART_TYPES["BAR"] = "bar";
    CHART_TYPES["BAR3D"] = "bar3D";
    CHART_TYPES["BUBBLE"] = "bubble";
    CHART_TYPES["DOUGHNUT"] = "doughnut";
    CHART_TYPES["LINE"] = "line";
    CHART_TYPES["PIE"] = "pie";
    CHART_TYPES["RADAR"] = "radar";
    CHART_TYPES["SCATTER"] = "scatter";
})(CHART_TYPES || (CHART_TYPES = {}));
/**
 * NOTE: 20170304: BULLET_TYPES: Only default is used so far. I'd like to combine the two pieces of code that use these before implementing these as options
 * Since we close <p> within the text object bullets, its slightly more difficult than combining into a func and calling to get the paraProp
 * and i'm not sure if anyone will even use these... so, skipping for now.
 */
var BULLET_TYPES;
(function (BULLET_TYPES) {
    BULLET_TYPES["DEFAULT"] = "&#x2022;";
    BULLET_TYPES["CHECK"] = "&#x2713;";
    BULLET_TYPES["STAR"] = "&#x2605;";
    BULLET_TYPES["TRIANGLE"] = "&#x25B6;";
})(BULLET_TYPES || (BULLET_TYPES = {}));
var BASE_SHAPES = Object.freeze({
    RECTANGLE: { displayName: 'Rectangle', name: 'rect', avLst: {} },
    LINE: { displayName: 'Line', name: 'line', avLst: {} },
});
// IMAGES (base64)
var IMG_BROKEN = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAGQAAAB3CAYAAAD1oOVhAAAGAUlEQVR4Xu2dT0xcRRzHf7tAYSsc0EBSIq2xEg8mtTGebVzEqOVIolz0siRE4gGTStqKwdpWsXoyGhMuyAVJOHBgqyvLNgonDkabeCBYW/8kTUr0wsJC+Wfm0bfuvn37Znbem9mR9303mJnf/Pb7ed95M7PDI5JIJPYJV5EC7e3t1N/fT62trdqViQCIu+bVgpIHEo/Hqbe3V/sdYVKHyWSSZmZm8ilVA0oeyNjYmEnaVC2Xvr6+qg5fAOJAz4DU1dURGzFSqZRVqtMpAFIGyMjICC0vL9PExIRWKADiAYTNshYWFrRCARAOEFZcCKWtrY0GBgaUTYkBRACIE4rKZwqACALR5RQAqQCIDqcASIVAVDsFQCSAqHQKgEgCUeUUAPEBRIVTAMQnEBvK5OQkbW9vk991CoAEAMQJxc86BUACAhKUUwAkQCBBOAVAAgbi1ykAogCIH6cAiCIgsk4BEIVAZJwCIIqBVLqiBxANQFgXS0tLND4+zl08AogmIG5OSSQS1gGKwgtANAIRcQqAaAbCe6YASBWA2E6xDyeyDUl7+AKQMkDYYevm5mZHabA/Li4uUiaTsYLau8QA4gLE/hU7wajyYtv1hReDAiAOxQcHBymbzark4BkbQKom/X8dp9Npmpqasn4BIAYAYSnYp+4BBEAMUcCwNOCQsAKZnp62NtQOw8WmwT09PUo+ijaHsOMx7GppaaH6+nolH0Z10K2tLVpdXbW6UfV3mNqBdHd3U1NTk2rtlMRfW1uj2dlZAFGirkRQAJEQTWUTAFGprkRsAJEQTWUTAFGprkRsAJEQTWUTAFGprkRsAJEQTWUTAFGprkRsAJEQTWUTAFGprkRsAJEQTWUTAGHqrm8caPzQ0WC1logbeiC7X3xJm0PvUmRzh45cuki1588FAmVn9BO6P3yF9utrqGH0MtW82S8UN9RA9v/4k7InjhcJFTs/TLVXLwmJV67S7vD7tHF5pKi46fYdosdOcOOGG8j1OcqefbFEJD9Q3GCwDhqT31HklS4A8VRgfYM2Op6k3bt/BQJl58J7lPvwg5JYNccepaMry0LPqFA7hCm39+NNyp2J0172b19QysGINj5CsRtpij57musOViH0QPJQXn6J9u7dlYJSFkbrMYolrwvDAJAC+WWdEpQz7FTgECeUCpzi6YxvvqXoM6eEhqnCSgDikEzUKUE7Aw7xuHctKB5OYU3dZlNR9syQdAaAcAYTC0pXF+39c09o2Ik+3EqxVKqiB7hbYAxZkk4pbBaEM+AQofv+wTrFwylBOQNABIGwavdfe4O2pg5elO+86l99nY58/VUF0byrYsjiSFluNlXYrOHcBar7+EogUADEQ0YRGHbzoKAASBkg2+9cpM1rV0tK2QOcXW7bLEFAARAXIF4w2DrDWoeUWaf4hQIgDiA8GPZ2iNfi0Q8UACkAIgrDbrJ385eDxaPLLrEsFAB5oG6lMPJQPLZZZKAACBGVhcG2Q+bmuLu2nk55e4jqPv1IeEoceiBeX7s2zCa5MAqdstl91vfXwaEGsv/rb5TtOFk6tWXOuJGh6KmnhO9sayrMninPx103JBtXblHkice58cINZP4Hyr5wpkgkdiChEmc4FWazLzenNKa/p0jncwDiqcD6BuWePk07t1asatZGoYQzSqA4nFJ7soNiP/+EUyfc25GI2GG53dHPrKo1g/1Cw4pIXLrzO+1c+/wg7tBbFDle/EbQcjFCPWQJCau5EoBoFpzXHYDwFNJcDiCaBed1ByA8hTSXA4hmwXndAQhPIc3lAKJZcF53AMJTSHM5gGgWnNcdgPAU0lwOIJoF53UHIDyFNJcfSiCdnZ0Ui8U0SxlMd7lcjubn561gh+Y1scFIU/0o/3sgeLO12E2k7UXKYumgFoAYdg8ACIAYpoBh6cAhAGKYAoalA4cAiGEKGJYOHAIghilgWDpwCIAYpoBh6cAhAGKYAoalA4cAiGEKGJYOHAIghilgWDpwCIAYpoBh6ZQ4JB6PKzviYthnNy4d9h+1M5mMlVckkUjsG5dhiBMCEMPg/wuOfrZZ/RSywQAAAABJRU5ErkJggg==';
var IMG_PLAYBTN = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAyAAAAHCCAYAAAAXY63IAAAACXBIWXMAAAsTAAALEwEAmpwYAAAKT2lDQ1BQaG90b3Nob3AgSUNDIHByb2ZpbGUAAHjanVNnVFPpFj333vRCS4iAlEtvUhUIIFJCi4AUkSYqIQkQSoghodkVUcERRUUEG8igiAOOjoCMFVEsDIoK2AfkIaKOg6OIisr74Xuja9a89+bN/rXXPues852zzwfACAyWSDNRNYAMqUIeEeCDx8TG4eQuQIEKJHAAEAizZCFz/SMBAPh+PDwrIsAHvgABeNMLCADATZvAMByH/w/qQplcAYCEAcB0kThLCIAUAEB6jkKmAEBGAYCdmCZTAKAEAGDLY2LjAFAtAGAnf+bTAICd+Jl7AQBblCEVAaCRACATZYhEAGg7AKzPVopFAFgwABRmS8Q5ANgtADBJV2ZIALC3AMDOEAuyAAgMADBRiIUpAAR7AGDIIyN4AISZABRG8lc88SuuEOcqAAB4mbI8uSQ5RYFbCC1xB1dXLh4ozkkXKxQ2YQJhmkAuwnmZGTKBNA/g88wAAKCRFRHgg/P9eM4Ors7ONo62Dl8t6r8G/yJiYuP+5c+rcEAAAOF0ftH+LC+zGoA7BoBt/qIl7gRoXgugdfeLZrIPQLUAoOnaV/Nw+H48PEWhkLnZ2eXk5NhKxEJbYcpXff5nwl/AV/1s+X48/Pf14L7iJIEyXYFHBPjgwsz0TKUcz5IJhGLc5o9H/LcL//wd0yLESWK5WCoU41EScY5EmozzMqUiiUKSKcUl0v9k4t8s+wM+3zUAsGo+AXuRLahdYwP2SycQWHTA4vcAAPK7b8HUKAgDgGiD4c93/+8//UegJQCAZkmScQAAXkQkLlTKsz/HCAAARKCBKrBBG/TBGCzABhzBBdzBC/xgNoRCJMTCQhBCCmSAHHJgKayCQiiGzbAdKmAv1EAdNMBRaIaTcA4uwlW4Dj1wD/phCJ7BKLyBCQRByAgTYSHaiAFiilgjjggXmYX4IcFIBBKLJCDJiBRRIkuRNUgxUopUIFVIHfI9cgI5h1xGupE7yAAygvyGvEcxlIGyUT3UDLVDuag3GoRGogvQZHQxmo8WoJvQcrQaPYw2oefQq2gP2o8+Q8cwwOgYBzPEbDAuxsNCsTgsCZNjy7EirAyrxhqwVqwDu4n1Y8+xdwQSgUXACTYEd0IgYR5BSFhMWE7YSKggHCQ0EdoJNwkDhFHCJyKTqEu0JroR+cQYYjIxh1hILCPWEo8TLxB7iEPENyQSiUMyJ7mQAkmxpFTSEtJG0m5SI+ksqZs0SBojk8naZGuyBzmULCAryIXkneTD5DPkG+Qh8lsKnWJAcaT4U+IoUspqShnlEOU05QZlmDJBVaOaUt2ooVQRNY9aQq2htlKvUYeoEzR1mjnNgxZJS6WtopXTGmgXaPdpr+h0uhHdlR5Ol9BX0svpR+iX6AP0dwwNhhWDx4hnKBmbGAcYZxl3GK+YTKYZ04sZx1QwNzHrmOeZD5lvVVgqtip8FZHKCpVKlSaVGyovVKmqpqreqgtV81XLVI+pXlN9rkZVM1PjqQnUlqtVqp1Q61MbU2epO6iHqmeob1Q/pH5Z/YkGWcNMw09DpFGgsV/jvMYgC2MZs3gsIWsNq4Z1gTXEJrHN2Xx2KruY/R27iz2qqaE5QzNKM1ezUvOUZj8H45hx+Jx0TgnnKKeX836K3hTvKeIpG6Y0TLkxZVxrqpaXllirSKtRq0frvTau7aedpr1Fu1n7gQ5Bx0onXCdHZ4/OBZ3nU9lT3acKpxZNPTr1ri6qa6UbobtEd79up+6Ynr5egJ5Mb6feeb3n+hx9L/1U/W36p/VHDFgGswwkBtsMzhg8xTVxbzwdL8fb8VFDXcNAQ6VhlWGX4YSRudE8o9VGjUYPjGnGXOMk423GbcajJgYmISZLTepN7ppSTbmmKaY7TDtMx83MzaLN1pk1mz0x1zLnm+eb15vft2BaeFostqi2uGVJsuRaplnutrxuhVo5WaVYVVpds0atna0l1rutu6cRp7lOk06rntZnw7Dxtsm2qbcZsOXYBtuutm22fWFnYhdnt8Wuw+6TvZN9un2N/T0HDYfZDqsdWh1+c7RyFDpWOt6azpzuP33F9JbpL2dYzxDP2DPjthPLKcRpnVOb00dnF2e5c4PziIuJS4LLLpc+Lpsbxt3IveRKdPVxXeF60vWdm7Obwu2o26/uNu5p7ofcn8w0nymeWTNz0MPIQ+BR5dE/C5+VMGvfrH5PQ0+BZ7XnIy9jL5FXrdewt6V3qvdh7xc+9j5yn+M+4zw33jLeWV/MN8C3yLfLT8Nvnl+F30N/I/9k/3r/0QCngCUBZwOJgUGBWwL7+Hp8Ib+OPzrbZfay2e1BjKC5QRVBj4KtguXBrSFoyOyQrSH355jOkc5pDoVQfujW0Adh5mGLw34MJ4WHhVeGP45wiFga0TGXNXfR3ENz30T6RJZE3ptnMU85ry1KNSo+qi5qPNo3ujS6P8YuZlnM1VidWElsSxw5LiquNm5svt/87fOH4p3iC+N7F5gvyF1weaHOwvSFpxapLhIsOpZATIhOOJTwQRAqqBaMJfITdyWOCnnCHcJnIi/RNtGI2ENcKh5O8kgqTXqS7JG8NXkkxTOlLOW5hCepkLxMDUzdmzqeFpp2IG0yPTq9MYOSkZBxQqohTZO2Z+pn5mZ2y6xlhbL+xW6Lty8elQfJa7OQrAVZLQq2QqboVFoo1yoHsmdlV2a/zYnKOZarnivN7cyzytuQN5zvn//tEsIS4ZK2pYZLVy0dWOa9rGo5sjxxedsK4xUFK4ZWBqw8uIq2Km3VT6vtV5eufr0mek1rgV7ByoLBtQFr6wtVCuWFfevc1+1dT1gvWd+1YfqGnRs+FYmKrhTbF5cVf9go3HjlG4dvyr+Z3JS0qavEuWTPZtJm6ebeLZ5bDpaql+aXDm4N2dq0Dd9WtO319kXbL5fNKNu7g7ZDuaO/PLi8ZafJzs07P1SkVPRU+lQ27tLdtWHX+G7R7ht7vPY07NXbW7z3/T7JvttVAVVN1WbVZftJ+7P3P66Jqun4lvttXa1ObXHtxwPSA/0HIw6217nU1R3SPVRSj9Yr60cOxx++/p3vdy0NNg1VjZzG4iNwRHnk6fcJ3/ceDTradox7rOEH0x92HWcdL2pCmvKaRptTmvtbYlu6T8w+0dbq3nr8R9sfD5w0PFl5SvNUyWna6YLTk2fyz4ydlZ19fi753GDborZ752PO32oPb++6EHTh0kX/i+c7vDvOXPK4dPKy2+UTV7hXmq86X23qdOo8/pPTT8e7nLuarrlca7nuer21e2b36RueN87d9L158Rb/1tWeOT3dvfN6b/fF9/XfFt1+cif9zsu72Xcn7q28T7xf9EDtQdlD3YfVP1v+3Njv3H9qwHeg89HcR/cGhYPP/pH1jw9DBY+Zj8uGDYbrnjg+OTniP3L96fynQ89kzyaeF/6i/suuFxYvfvjV69fO0ZjRoZfyl5O/bXyl/erA6xmv28bCxh6+yXgzMV70VvvtwXfcdx3vo98PT+R8IH8o/2j5sfVT0Kf7kxmTk/8EA5jz/GMzLdsAAAAgY0hSTQAAeiUAAICDAAD5/wAAgOkAAHUwAADqYAAAOpgAABdvkl/FRgAAFRdJREFUeNrs3WFz2lbagOEnkiVLxsYQsP//z9uZZmMswJIlS3k/tPb23U3TOAUM6Lpm8qkzbXM4A7p1dI4+/etf//oWAAAAB3ARETGdTo0EAACwV1VVRWIYAACAQxEgAACAAAEAAAQIAACAAAEAAAQIAACAAAEAAAQIAAAgQAAAAAQIAAAgQAAAAAQIAAAgQAAAAAECAAAgQAAAAAECAAAgQAAAAAECAAAIEAAAAAECAAAIEAAAAAECAAAIEAAAQIAAAAAIEAAAQIAAAAAIEAAAQIAAAAACBAAAQIAAAAACBAAAQIAAAAACBAAAQIAAAAACBAAAECAAAAACBAAAECAAAAACBAAAECAAAIAAAQAAECAAAIAAAQAAECAAAIAAAQAABAgAAIAAAQAABAgAAIAAAQAABAgAACBAAAAABAgAACBAAAAABAgAACBAAAAAAQIAACBAAAAAAQIAACBAAAAAAQIAACBAAAAAAQIAAAgQAAAAAQIAAAgQAAAAAQIAAAgQAABAgAAAAAgQAABAgAAAAAgQAABAgAAAAAIEAABAgAAAAAIEAABAgAAAAAIEAAAQIAAAAAIEAAAQIAAAAAIEAAAQIAAAgAABAAAQIAAAgAABAAAQIAAAgAABAAAQIAAAgAABAAAECAAAgAABAAAECAAAgAABAAAECAAAIEAAAAAECAAAIEAAAAAECAAAIEAAAAABAgAAIEAAAAABAgAAIEAAAAABAgAACBAAAAABAgAACBAAAAABAgAACBAAAECAAAAACBAAAECAAAAACBAAAECAAAAAAgQAAECAAAAAAgQAAECAAAAAAgQAAECAAAAAAgQAABAgAAAAAgQAABAgAAAAAgQAABAgAACAAAEAABAgAACAAAEAABAgAACAAAEAAAQIAACAAAEAAAQIAACAAAEAAAQIAAAgQAAAAPbnwhAA8CuGYYiXl5fv/7hcXESSuMcFgAAB4G90XRffvn2L5+fniIho2zYiIvq+j77vf+nfmaZppGkaERF5nkdExOXlZXz69CmyLDPoAAIEgDFo2zaen5/j5eUl+r6Pruv28t/5c7y8Bs1ms3n751mWRZqmcXFxEZeXl2+RAoAAAeBEDcMQbdu+/dlXbPyKruve/n9ewyTLssjz/O2PR7oABAgAR67v+2iaJpqmeVt5OBWvUbLdbiPi90e3iqKIoijeHucCQIAAcATRsd1uo2maX96zcYxeV26qqoo0TaMoiphMJmIEQIAAcGjDMERd11HX9VE9WrXvyNput5FlWZRlGWVZekwLQIAAsE+vjyjVdT3qMei6LqqqirIsYzKZOFkLQIAAsEt1XcfT09PJ7es4xLjUdR15nsfV1VWUZWlQAAQIAP/kAnu9Xp/V3o59eN0vsl6v4+bmRogACBAAhMf+9X0fq9VKiAAIEAB+RtM0UVWV8NhhiEyn0yiKwqAACBAAXr1uqrbHY/ch8vDwEHmex3Q6tVkdQIAAjNswDLHZbN5evsd+tG0bX758iclkEtfX147vBRAgAOPTNE08Pj7GMAwG40BejzC+vb31WBaAAAEYh9f9CR63+hjDMLw9ljWfz62GAOyZb1mAD9Q0TXz58kV8HIG2beO3336LpmkMBsAeWQEB+ADDMERVVaN+g/mxfi4PDw9RlmVMp1OrIQACBOD0dV0XDw8PjtY9YnVdR9u2MZ/PnZQFsGNu7QAc+ML269ev4uME9H0fX79+tUoFsGNWQAAOZLVauZg9McMwxGq1iufn55jNZgYEQIAAnMZF7MPDg43mJ6yu6+j73ilZADvgWxRgj7qui69fv4qPM9C2rcfnAAQIwPHHR9d1BuOMPtMvX774TAEECMBxxoe3mp+fYRiEJYAAATgeryddiY/zjxAvLQQQIAAfHh+r1Up8jCRCHh4enGwGIEAAPkbTNLFarQzEyKxWKyshAAIE4LC6rovHx0cDMVKPj4/2hAAIEIDDxYc9H+NmYzqAAAEQH4gQAAECcF4XnI+Pj+IDcwJAgADs38PDg7vd/I+u6+Lh4cFAAAgQgN1ZrVbRtq2B4LvatnUiGoAAAdiNuq69+wHzBECAAOxf13VRVZWB4KdUVeUxPQABAvBrXt98bYMx5gyAAAHYu6qqou97A8G79H1v1QxAgAC8T9M0nufnl9V1HU3TGAgAAQLw9/q+j8fHx5P6f86yLMqy9OEdEe8HARAgAD9ltVqd3IXjp0+fYjabxWKxiDzPfYhH4HU/CIAAAeAvNU1z0u/7yPM8FotFzGazSBJf+R+tbVuPYgECxBAAfN8wDCf36NVfKcsy7u7u4vr62gf7wTyKBQgQAL5rs9mc1YVikiRxc3MT9/f3URSFD/gDw3az2RgIQIAA8B9d18V2uz3Lv1uapjGfz2OxWESWZT7sD7Ddbr2gEBAgAPzHGN7bkOd5LJfLmE6n9oeYYwACBOCjnPrG8/eaTCZxd3cXk8nEh39ANqQDAgSAiBjnnekkSWI6ncb9/b1je801AAECcCh1XUff96P9+6dpGovFIhaLRaRpakLsWd/3Ude1gQAECMBYrddrgxC/7w+5v7+P6+tr+0PMOQABArAPY1/9+J6bm5u4u7uLsiwNxp5YBQEECMBIuRP9Fz8USRKz2SyWy6X9IeYegAAB2AWrH38vy7JYLBYxn8/tD9kxqyCAAAEYmaenJ4Pwk4qiiOVyaX+IOQggQAB+Rdd1o3rvx05+PJIkbm5uYrlc2h+yI23bejs6IEAAxmC73RqEX5Smacxms1gsFpFlmQExFwEECMCPDMPg2fsdyPM8lstlzGYzj2X9A3VdxzAMBgIQIADnfMHH7pRlGXd3d3F9fW0wzEkAAQLgYu8APyx/7A+5v7+PoigMiDkJIEAAIn4/+tSm3/1J0zTm83ksFgvH9r5D13WOhAYECMA5suH3MPI8j/v7+5hOp/aHmJsAAgQYr6ZpDMIBTSaTuLu7i8lkYjDMTUCAAIxL3/cec/mIH50kiel0Gvf395HnuQExPwEBAjAO7jB/rDRNY7FYxHw+tz/EHAUECICLOw6jKIq4v7+P6+tr+0PMUUCAAJynYRiibVsDcURubm7i7u4uyrI0GH9o29ZLCQEBAnAuF3Yc4Q9SksRsNovlcml/iLkKCBAAF3UcRpZlsVgsYjabjX5/iLkKnKMLQwC4qOMYlWUZl5eXsd1u4+npaZSPI5mrwDmyAgKMjrefn9CPVJLEzc1NLJfLUe4PMVcBAQJw4txRPk1pmsZsNovFYhFZlpmzAAIE4DQ8Pz8bhBOW53ksl8uYzWajObbXnAXOjT0gwKi8vLwYhDPw5/0hm83GnAU4IVZAgFHp+94gnMsP2B/7Q+7v78/62F5zFhAgACfMpt7zk6ZpLBaLWCwWZ3lsrzkLCBAAF3IcoTzP4/7+PqbT6dntDzF3AQECcIK+fftmEEZgMpnE3d1dTCYTcxdAgAB8HKcJjejHLUliOp3Gcrk8i/0h5i4gQADgBGRZFovFIubz+VnuDwE4RY7hBUbDC93GqyiKKIoi1ut1PD09xTAM5i7AB7ECAsBo3NzcxN3dXZRlaTAABAjAfnmfAhG/7w+ZzWaxWCxOZn+IuQsIEAABwonL8zwWi0XMZrOj3x9i7gLnxB4QAEatLMu4vLyM7XZ7kvtDAE6NFRAA/BgmSdzc3MRyuYyiKAwIgAAB+Gfc1eZnpGka8/k8FotFZFlmDgMIEIBf8/LyYhD4aXmex3K5jNlsFkmSmMMAO2QPCAD8hT/vD9lsNgYEYAesgADAj34o/9gfcn9/fzLH9gIIEAAAgPAIFgD80DAMsdlsYrvdGgwAAQIA+/O698MJVAACBOB9X3YXvu74eW3bRlVV0XWdOQwgQADe71iOUuW49X0fVVVF0zTmMIAAAYD9GIbBUbsAAgQA9q+u61iv19H3vcEAECAAu5OmqYtM3rRtG+v1Otq2PYm5CyBAAAQIJ6jv+1iv11HX9UnNXQABAgAnZr1ex9PTk2N1AQQIwP7leX4Sj9uwe03TRFVVJ7sClue5DxEQIABw7Lqui6qqhCeAAAE4vMvLS8esjsQwDLHZbGK73Z7N3AUQIAAn5tOnTwZhBF7f53FO+zzMXUCAAJygLMsMwhlr2zZWq9VZnnRm7gICBOCEL+S6rjMQZ6Tv+1itVme7z0N8AAIE4ISlaSpAzsQwDG+PW537nAUQIACn+qV34WvvHNR1HVVVjeJ9HuYsIEAATpiTsE5b27ZRVdWoVrGcgAUIEIBT/tJzN/kk9X0fVVVF0zSj+7t7CSEgQABOWJIkNqKfkNd9Hk9PT6N43Oq/2YAOCBCAM5DnuQA5AXVdx3q9Pstjdd8zVwEECMAZXNSdyxuyz1HXdVFV1dkeqytAAAEC4KKOIzAMQ1RVFXVdGwxzFRAgAOcjSZLI89wd9iOyXq9Hu8/jR/GRJImBAAQIwDkoikKAHIGmaaKqqlHv8/jRHAUQIABndHFXVZWB+CB938dqtRKBAgQQIADjkKZppGnqzvuBDcMQm83GIQA/OT8BBAjAGSmKwoXwAW2329hsNvZ5/OTcBBAgAGdmMpkIkANo2zZWq5XVpnfOTQABAnBm0jT1VvQ96vs+qqqKpmkMxjtkWebxK0CAAJyrsiwFyI4Nw/D2uBW/NicBBAjAGV/sOQ1rd+q6jqqq7PMQIAACBOB7kiSJsiy9ffsfats2qqqymrSD+PDyQUCAAJy5q6srAfKL+r6P9Xpt/HY4FwEECMCZy/M88jz3Urx3eN3n8fT05HGrHc9DAAECMAJXV1cC5CfVdR3r9dqxunuYgwACBGAkyrJ0Uf03uq6LqqqE2h6kaWrzOSBAAMbm5uYmVquVgfgvwzBEVVX2eex57gEIEICRsQryv9brtX0ee2b1AxAgACNmFeR3bdvGarUSYweacwACBGCkxr4K0vd9rFYr+zwOxOoHIEAAGOUqyDAMsdlsYrvdmgAHnmsAAgRg5MqyjKenp9GsAmy329hsNvZ5HFie51Y/gFFKDAHA/xrDnem2bePLly9RVZX4MMcADsYKCMB3vN6dPsejZ/u+j6qqomkaH/QHKcvSW88BAQLA/zedTuP5+flsVgeGYXh73IqPkyRJTKdTAwGM93vQEAD89YXi7e3tWfxd6rqO3377TXwcgdvb20gSP7/AeFkBAfiBoigiz/OT3ZDetm2s12vH6h6JPM+jKAoDAYyaWzAAf2M2m53cHetv377FarWKf//73+LjWH5wkyRms5mBAHwfGgKAH0vT9OQexeq67iw30J+y29vbSNPUQAACxBAA/L2iKDw6g/kDIEAADscdbH7FKa6gAQgQgGP4wkySmM/nBoJ3mc/nTr0CECAAvybLMhuJ+Wmz2SyyLDMQAAIE4NeVZRllWRoIzBMAAQJwGO5s8yNWygAECMDOff78WYTw3fj4/PmzgQAQIAA7/gJNkri9vbXBGHMCQIAAHMbr3W4XnCRJYlUMQIAAiBDEB4AAATjDCJlOpwZipKbTqfgAECAAh1WWpZOPRmg2mzluF+AdLgwBwG4jJCKiqqoYhsGAnLEkSWI6nYoPgPd+fxoCgN1HiD0h5x8fnz9/Fh8AAgTgONiYfv7xYc8HgAABOMoIcaHqMwVAgAC4YOVd8jz3WQIIEIAT+KJNklgul/YLnLCyLGOxWHikDkCAAJyO2WzmmF6fG8DoOYYX4IDKsoyLi4t4eHiIvu8NyBFL0zTm87lHrgB2zAoIwIFlWRbL5TKKojAYR6ooilgul+IDYA+sgAB8gCRJYj6fR9M08fj46KWFR/S53N7eikMAAQJwnoqiiCzLYrVaRdu2BuQD5Xkes9ks0jQ1GAACBOB8pWkai8XCasgHseoBIEAARqkoisjzPKqqirquDcgBlGUZ0+nU8boAAgRgnJIkidlsFldXV7Ferz2WtSd5nsd0OrXJHECAAPB6gbxYLKKu61iv147s3ZE0TWM6nXrcCkCAAPA9ZVlGWZZCZAfhcXNz4230AAIEACEiPAAECABHHyJPT0/2iPyFPM/j6upKeAAIEAB2GSJt28bT05NTs/40LpPJxOZyAAECwD7kef52olNd11HXdXRdN6oxyLLsLcgcpwsgQAA4gCRJYjKZxGQyib7vY7vdRtM0Z7tXJE3TKIoiJpOJN5cDCBAAPvrifDqdxnQ6jb7vo2maaJrm5PeL5HkeRVFEURSiA0CAAHCsMfK6MjIMQ7Rt+/bn2B/VyrLs7RGzPM89XgUgQAA4JUmSvK0gvGrbNp6fn+Pl5SX6vv+wKMmyLNI0jYuLi7i8vIw8z31gAAIEgHPzurrwZ13Xxbdv3+L5+fktUiIi+r7/5T0laZq+PTb1+t+7vLyMT58+ObEKQIAAMGavQfB3qxDDMMTLy8v3f1wuLjwyBYAAAWB3kiTxqBQA7//9MAQAAIAAAQAABAgAAIAAAQAABAgAAIAAAQAABAgAACBAAAAABAgAACBAAAAABAgAACBAAAAAAQIAACBAAAAAAQIAACBAAAAAAQIAAAgQAAAAAQIAAAgQAAAAAQIAAAgQAABAgAAAAAgQAABAgAAAAAgQAABAgAAAAAIEAABAgAAAAAIEAABAgAAAAAIEAABAgAAAAAIEAAAQIAAAAAIEAAAQIAAAAAIEAAAQIAAAgAABAAAQIAAAgAABAAAQIAAAgAABAAAECAAAgAABAAAECAAAgAABAAAECAAAIEAAAAAECAAAIEAAAAAECAAAIEAAAAABAgAAIEAAAAABAgAAIEAAAAABAgAAIEAAAAABAgAACBAAAAABAgAACBAAAAABAgAACBAAAECAAAAACBAAAECAAAAACBAAAECAAAAAAgQAAECAAAAAAgQAAECAAAAAAgQAABAgAAAAAgQAABAgAAAAAgQAABAgAACAAAEAABAgAACAAAEAABAgAACAAAEAAASIIQAAAAQIAAAgQAAAAAQIAAAgQAAAAAQIAAAgQAAAAAECAAAgQAAAAAECAAAgQAAAAAECAAAIEAAAAAECAAAIEAAAAAECAAAIEAAAQIAAAAAIEAAAQIAAAAAIEAAAQIAAAAACBAAAQIAAAAACBAAAQIAAAAACBAAAECAAAAACBAAAECAAAAACBAAAECAAAAACBAAAECAAAIAAAQAAECAAAIAAAQAAECAAAIAAAQAABAgAAIAAAQAABAgAAIAAAQAABAgAACBAAAAAdu0iIqKqKiMBAADs3f8NAFFjCf5mB+leAAAAAElFTkSuQmCC';

/**
* PptxGenJS - All PowerPoint Shapes
*/
var PowerPointShapes = Object.freeze({
    ACTION_BUTTON_BACK_OR_PREVIOUS: {
        'displayName': 'Action Button: Back or Previous',
        'name': 'actionButtonBackPrevious',
        'avLst': {}
    },
    ACTION_BUTTON_BEGINNING: {
        'displayName': 'Action Button: Beginning',
        'name': 'actionButtonBeginning',
        'avLst': {}
    },
    ACTION_BUTTON_CUSTOM: {
        'displayName': 'Action Button: Custom',
        'name': 'actionButtonBlank',
        'avLst': {}
    },
    ACTION_BUTTON_DOCUMENT: {
        'displayName': 'Action Button: Document',
        'name': 'actionButtonDocument',
        'avLst': {}
    },
    ACTION_BUTTON_END: {
        'displayName': 'Action Button: End',
        'name': 'actionButtonEnd',
        'avLst': {}
    },
    ACTION_BUTTON_FORWARD_OR_NEXT: {
        'displayName': 'Action Button: Forward or Next',
        'name': 'actionButtonForwardNext',
        'avLst': {}
    },
    ACTION_BUTTON_HELP: {
        'displayName': 'Action Button: Help',
        'name': 'actionButtonHelp',
        'avLst': {}
    },
    ACTION_BUTTON_HOME: {
        'displayName': 'Action Button: Home',
        'name': 'actionButtonHome',
        'avLst': {}
    },
    ACTION_BUTTON_INFORMATION: {
        'displayName': 'Action Button: Information',
        'name': 'actionButtonInformation',
        'avLst': {}
    },
    ACTION_BUTTON_MOVIE: {
        'displayName': 'Action Button: Movie',
        'name': 'actionButtonMovie',
        'avLst': {}
    },
    ACTION_BUTTON_RETURN: {
        'displayName': 'Action Button: Return',
        'name': 'actionButtonReturn',
        'avLst': {}
    },
    ACTION_BUTTON_SOUND: {
        'displayName': 'Action Button: Sound',
        'name': 'actionButtonSound',
        'avLst': {}
    },
    ARC: {
        'displayName': 'Arc',
        'name': 'arc',
        'avLst': {
            'adj1': 16200000,
            'adj2': 0
        }
    },
    BALLOON: {
        'displayName': 'Rounded Rectangular Callout',
        'name': 'wedgeRoundRectCallout',
        'avLst': {
            'adj1': -20833,
            'adj2': 62500,
            'adj3': 16667
        }
    },
    BENT_ARROW: {
        'displayName': 'Bent Arrow',
        'name': 'bentArrow',
        'avLst': {
            'adj1': 25000,
            'adj2': 25000,
            'adj3': 25000,
            'adj4': 43750
        }
    },
    BENT_UP_ARROW: {
        'displayName': 'Bent-Up Arrow',
        'name': 'bentUpArrow',
        'avLst': {
            'adj1': 25000,
            'adj2': 25000,
            'adj3': 25000
        }
    },
    BEVEL: {
        'displayName': 'Bevel',
        'name': 'bevel',
        'avLst': {
            'adj': 12500
        }
    },
    BLOCK_ARC: {
        'displayName': 'Block Arc',
        'name': 'blockArc',
        'avLst': {
            'adj1': 10800000,
            'adj2': 0,
            'adj3': 25000
        }
    },
    CAN: {
        'displayName': 'Can',
        'name': 'can',
        'avLst': {
            'adj': 25000
        }
    },
    CHART_PLUS: {
        'displayName': 'Chart Plus',
        'name': 'chartPlus',
        'avLst': {}
    },
    CHART_STAR: {
        'displayName': 'Chart Star',
        'name': 'chartStar',
        'avLst': {}
    },
    CHART_X: {
        'displayName': 'Chart X',
        'name': 'chartX',
        'avLst': {}
    },
    CHEVRON: {
        'displayName': 'Chevron',
        'name': 'chevron',
        'avLst': {
            'adj': 50000
        }
    },
    CHORD: {
        'displayName': 'Chord',
        'name': 'chord',
        'avLst': {
            'adj1': 2700000,
            'adj2': 16200000
        }
    },
    CIRCULAR_ARROW: {
        'displayName': 'Circular Arrow',
        'name': 'circularArrow',
        'avLst': {
            'adj1': 12500,
            'adj2': 1142319,
            'adj3': 20457681,
            'adj4': 10800000,
            'adj5': 12500
        }
    },
    CLOUD: {
        'displayName': 'Cloud',
        'name': 'cloud',
        'avLst': {}
    },
    CLOUD_CALLOUT: {
        'displayName': 'Cloud Callout',
        'name': 'cloudCallout',
        'avLst': {
            'adj1': -20833,
            'adj2': 62500
        }
    },
    CORNER: {
        'displayName': 'Corner',
        'name': 'corner',
        'avLst': {
            'adj1': 50000,
            'adj2': 50000
        }
    },
    CORNER_TABS: {
        'displayName': 'Corner Tabs',
        'name': 'cornerTabs',
        'avLst': {}
    },
    CROSS: {
        'displayName': 'Cross',
        'name': 'plus',
        'avLst': {
            'adj': 25000
        }
    },
    CUBE: {
        'displayName': 'Cube',
        'name': 'cube',
        'avLst': {
            'adj': 25000
        }
    },
    CURVED_DOWN_ARROW: {
        'displayName': 'Curved Down Arrow',
        'name': 'curvedDownArrow',
        'avLst': {
            'adj1': 25000,
            'adj2': 50000,
            'adj3': 25000
        }
    },
    CURVED_DOWN_RIBBON: {
        'displayName': 'Curved Down Ribbon',
        'name': 'ellipseRibbon',
        'avLst': {
            'adj1': 25000,
            'adj2': 50000,
            'adj3': 12500
        }
    },
    CURVED_LEFT_ARROW: {
        'displayName': 'Curved Left Arrow',
        'name': 'curvedLeftArrow',
        'avLst': {
            'adj1': 25000,
            'adj2': 50000,
            'adj3': 25000
        }
    },
    CURVED_RIGHT_ARROW: {
        'displayName': 'Curved Right Arrow',
        'name': 'curvedRightArrow',
        'avLst': {
            'adj1': 25000,
            'adj2': 50000,
            'adj3': 25000
        }
    },
    CURVED_UP_ARROW: {
        'displayName': 'Curved Up Arrow',
        'name': 'curvedUpArrow',
        'avLst': {
            'adj1': 25000,
            'adj2': 50000,
            'adj3': 25000
        }
    },
    CURVED_UP_RIBBON: {
        'displayName': 'Curved Up Ribbon',
        'name': 'ellipseRibbon2',
        'avLst': {
            'adj1': 25000,
            'adj2': 50000,
            'adj3': 12500
        }
    },
    DECAGON: {
        'displayName': 'Decagon',
        'name': 'decagon',
        'avLst': {
            'vf': 105146
        }
    },
    DIAGONAL_STRIPE: {
        'displayName': 'Diagonal Stripe',
        'name': 'diagStripe',
        'avLst': {
            'adj': 50000
        }
    },
    DIAMOND: {
        'displayName': 'Diamond',
        'name': 'diamond',
        'avLst': {}
    },
    DODECAGON: {
        'displayName': 'Dodecagon',
        'name': 'dodecagon',
        'avLst': {}
    },
    DONUT: {
        'displayName': 'Donut',
        'name': 'donut',
        'avLst': {
            'adj': 25000
        }
    },
    DOUBLE_BRACE: {
        'displayName': 'Double Brace',
        'name': 'bracePair',
        'avLst': {
            'adj': 8333
        }
    },
    DOUBLE_BRACKET: {
        'displayName': 'Double Bracket',
        'name': 'bracketPair',
        'avLst': {
            'adj': 16667
        }
    },
    DOUBLE_WAVE: {
        'displayName': 'Double Wave',
        'name': 'doubleWave',
        'avLst': {
            'adj1': 6250,
            'adj2': 0
        }
    },
    DOWN_ARROW: {
        'displayName': 'Down Arrow',
        'name': 'downArrow',
        'avLst': {
            'adj1': 50000,
            'adj2': 50000
        }
    },
    DOWN_ARROW_CALLOUT: {
        'displayName': 'Down Arrow Callout',
        'name': 'downArrowCallout',
        'avLst': {
            'adj1': 25000,
            'adj2': 25000,
            'adj3': 25000,
            'adj4': 64977
        }
    },
    DOWN_RIBBON: {
        'displayName': 'Down Ribbon',
        'name': 'ribbon',
        'avLst': {
            'adj1': 16667,
            'adj2': 50000
        }
    },
    EXPLOSION1: {
        'displayName': 'Explosion',
        'name': 'irregularSeal1',
        'avLst': {}
    },
    EXPLOSION2: {
        'displayName': 'Explosion',
        'name': 'irregularSeal2',
        'avLst': {}
    },
    FLOWCHART_ALTERNATE_PROCESS: {
        'displayName': 'Alternate process',
        'name': 'flowChartAlternateProcess',
        'avLst': {}
    },
    FLOWCHART_CARD: {
        'displayName': 'Card',
        'name': 'flowChartPunchedCard',
        'avLst': {}
    },
    FLOWCHART_COLLATE: {
        'displayName': 'Collate',
        'name': 'flowChartCollate',
        'avLst': {}
    },
    FLOWCHART_CONNECTOR: {
        'displayName': 'Connector',
        'name': 'flowChartConnector',
        'avLst': {}
    },
    FLOWCHART_DATA: {
        'displayName': 'Data',
        'name': 'flowChartInputOutput',
        'avLst': {}
    },
    FLOWCHART_DECISION: {
        'displayName': 'Decision',
        'name': 'flowChartDecision',
        'avLst': {}
    },
    FLOWCHART_DELAY: {
        'displayName': 'Delay',
        'name': 'flowChartDelay',
        'avLst': {}
    },
    FLOWCHART_DIRECT_ACCESS_STORAGE: {
        'displayName': 'Direct Access Storage',
        'name': 'flowChartMagneticDrum',
        'avLst': {}
    },
    FLOWCHART_DISPLAY: {
        'displayName': 'Display',
        'name': 'flowChartDisplay',
        'avLst': {}
    },
    FLOWCHART_DOCUMENT: {
        'displayName': 'Document',
        'name': 'flowChartDocument',
        'avLst': {}
    },
    FLOWCHART_EXTRACT: {
        'displayName': 'Extract',
        'name': 'flowChartExtract',
        'avLst': {}
    },
    FLOWCHART_INTERNAL_STORAGE: {
        'displayName': 'Internal Storage',
        'name': 'flowChartInternalStorage',
        'avLst': {}
    },
    FLOWCHART_MAGNETIC_DISK: {
        'displayName': 'Magnetic Disk',
        'name': 'flowChartMagneticDisk',
        'avLst': {}
    },
    FLOWCHART_MANUAL_INPUT: {
        'displayName': 'Manual Input',
        'name': 'flowChartManualInput',
        'avLst': {}
    },
    FLOWCHART_MANUAL_OPERATION: {
        'displayName': 'Manual Operation',
        'name': 'flowChartManualOperation',
        'avLst': {}
    },
    FLOWCHART_MERGE: {
        'displayName': 'Merge',
        'name': 'flowChartMerge',
        'avLst': {}
    },
    FLOWCHART_MULTIDOCUMENT: {
        'displayName': 'Multidocument',
        'name': 'flowChartMultidocument',
        'avLst': {}
    },
    FLOWCHART_OFFLINE_STORAGE: {
        'displayName': 'Offline Storage',
        'name': 'flowChartOfflineStorage',
        'avLst': {}
    },
    FLOWCHART_OFFPAGE_CONNECTOR: {
        'displayName': 'Off-page Connector',
        'name': 'flowChartOffpageConnector',
        'avLst': {}
    },
    FLOWCHART_OR: {
        'displayName': 'Or',
        'name': 'flowChartOr',
        'avLst': {}
    },
    FLOWCHART_PREDEFINED_PROCESS: {
        'displayName': 'Predefined Process',
        'name': 'flowChartPredefinedProcess',
        'avLst': {}
    },
    FLOWCHART_PREPARATION: {
        'displayName': 'Preparation',
        'name': 'flowChartPreparation',
        'avLst': {}
    },
    FLOWCHART_PROCESS: {
        'displayName': 'Process',
        'name': 'flowChartProcess',
        'avLst': {}
    },
    FLOWCHART_PUNCHED_TAPE: {
        'displayName': 'Punched Tape',
        'name': 'flowChartPunchedTape',
        'avLst': {}
    },
    FLOWCHART_SEQUENTIAL_ACCESS_STORAGE: {
        'displayName': 'Sequential Access Storage',
        'name': 'flowChartMagneticTape',
        'avLst': {}
    },
    FLOWCHART_SORT: {
        'displayName': 'Sort',
        'name': 'flowChartSort',
        'avLst': {}
    },
    FLOWCHART_STORED_DATA: {
        'displayName': 'Stored Data',
        'name': 'flowChartOnlineStorage',
        'avLst': {}
    },
    FLOWCHART_SUMMING_JUNCTION: {
        'displayName': 'Summing Junction',
        'name': 'flowChartSummingJunction',
        'avLst': {}
    },
    FLOWCHART_TERMINATOR: {
        'displayName': 'Terminator',
        'name': 'flowChartTerminator',
        'avLst': {}
    },
    FOLDED_CORNER: {
        'displayName': 'Folded Corner',
        'name': 'folderCorner',
        'avLst': {}
    },
    FRAME: {
        'displayName': 'Frame',
        'name': 'frame',
        'avLst': {
            'adj1': 12500
        }
    },
    FUNNEL: {
        'displayName': 'Funnel',
        'name': 'funnel',
        'avLst': {}
    },
    GEAR_6: {
        'displayName': 'Gear 6',
        'name': 'gear6',
        'avLst': {
            'adj1': 15000,
            'adj2': 3526
        }
    },
    GEAR_9: {
        'displayName': 'Gear 9',
        'name': 'gear9',
        'avLst': {
            'adj1': 10000,
            'adj2': 1763
        }
    },
    HALF_FRAME: {
        'displayName': 'Half Frame',
        'name': 'halfFrame',
        'avLst': {
            'adj1': 33333,
            'adj2': 33333
        }
    },
    HEART: {
        'displayName': 'Heart',
        'name': 'heart',
        'avLst': {}
    },
    HEPTAGON: {
        'displayName': 'Heptagon',
        'name': 'heptagon',
        'avLst': {
            'hf': 102572,
            'vf': 105210
        }
    },
    HEXAGON: {
        'displayName': 'Hexagon',
        'name': 'hexagon',
        'avLst': {
            'adj': 25000,
            'vf': 115470
        }
    },
    HORIZONTAL_SCROLL: {
        'displayName': 'Horizontal Scroll',
        'name': 'horizontalScroll',
        'avLst': {
            'adj': 12500
        }
    },
    ISOSCELES_TRIANGLE: {
        'displayName': 'Isosceles Triangle',
        'name': 'triangle',
        'avLst': {
            'adj': 50000
        }
    },
    LEFT_ARROW: {
        'displayName': 'Left Arrow',
        'name': 'leftArrow',
        'avLst': {
            'adj1': 50000,
            'adj2': 50000
        }
    },
    LEFT_ARROW_CALLOUT: {
        'displayName': 'Left Arrow Callout',
        'name': 'leftArrowCallout',
        'avLst': {
            'adj1': 25000,
            'adj2': 25000,
            'adj3': 25000,
            'adj4': 64977
        }
    },
    LEFT_BRACE: {
        'displayName': 'Left Brace',
        'name': 'leftBrace',
        'avLst': {
            'adj1': 8333,
            'adj2': 50000
        }
    },
    LEFT_BRACKET: {
        'displayName': 'Left Bracket',
        'name': 'leftBracket',
        'avLst': {
            'adj': 8333
        }
    },
    LEFT_CIRCULAR_ARROW: {
        'displayName': 'Left Circular Arrow',
        'name': 'leftCircularArrow',
        'avLst': {
            'adj1': 12500,
            'adj2': -1142319,
            'adj3': 1142319,
            'adj4': 10800000,
            'adj5': 12500
        }
    },
    LEFT_RIGHT_ARROW: {
        'displayName': 'Left-Right Arrow',
        'name': 'leftRightArrow',
        'avLst': {
            'adj1': 50000,
            'adj2': 50000
        }
    },
    LEFT_RIGHT_ARROW_CALLOUT: {
        'displayName': 'Left-Right Arrow Callout',
        'name': 'leftRightArrowCallout',
        'avLst': {
            'adj1': 25000,
            'adj2': 25000,
            'adj3': 25000,
            'adj4': 48123
        }
    },
    LEFT_RIGHT_CIRCULAR_ARROW: {
        'displayName': 'Left Right Circular Arrow',
        'name': 'leftRightCircularArrow',
        'avLst': {
            'adj1': 12500,
            'adj2': 1142319,
            'adj3': 20457681,
            'adj4': 11942319,
            'adj5': 12500
        }
    },
    LEFT_RIGHT_RIBBON: {
        'displayName': 'Left Right Ribbon',
        'name': 'leftRightRibbon',
        'avLst': {
            'adj1': 50000,
            'adj2': 50000,
            'adj3': 16667
        }
    },
    LEFT_RIGHT_UP_ARROW: {
        'displayName': 'Left-Right-Up Arrow',
        'name': 'leftRightUpArrow',
        'avLst': {
            'adj1': 25000,
            'adj2': 25000,
            'adj3': 25000
        }
    },
    LEFT_UP_ARROW: {
        'displayName': 'Left-Up Arrow',
        'name': 'leftUpArrow',
        'avLst': {
            'adj1': 25000,
            'adj2': 25000,
            'adj3': 25000
        }
    },
    LIGHTNING_BOLT: {
        'displayName': 'Lightning Bolt',
        'name': 'lightningBolt',
        'avLst': {}
    },
    LINE_CALLOUT_1: {
        'displayName': 'Line Callout 1',
        'name': 'borderCallout1',
        'avLst': {
            'adj1': 18750,
            'adj2': -8333,
            'adj3': 112500,
            'adj4': -38333
        }
    },
    LINE_CALLOUT_1_ACCENT_BAR: {
        'displayName': 'Line Callout 1 {Accent Bar}',
        'name': 'accentCallout1',
        'avLst': {
            'adj1': 18750,
            'adj2': -8333,
            'adj3': 112500,
            'adj4': -38333
        }
    },
    LINE_CALLOUT_1_BORDER_AND_ACCENT_BAR: {
        'displayName': 'Line Callout 1 {Border and Accent Bar}',
        'name': 'accentBorderCallout1',
        'avLst': {
            'adj1': 18750,
            'adj2': -8333,
            'adj3': 112500,
            'adj4': -38333
        }
    },
    LINE_CALLOUT_1_NO_BORDER: {
        'displayName': 'Line Callout 1 {No Border}',
        'name': 'callout1',
        'avLst': {
            'adj1': 18750,
            'adj2': -8333,
            'adj3': 112500,
            'adj4': -38333
        }
    },
    LINE_CALLOUT_2: {
        'displayName': 'Line Callout 2',
        'name': 'borderCallout2',
        'avLst': {
            'adj1': 18750,
            'adj2': -8333,
            'adj3': 18750,
            'adj4': -16667,
            'adj5': 112500,
            'adj6': -46667
        }
    },
    LINE_CALLOUT_2_ACCENT_BAR: {
        'displayName': 'Line Callout 2 {Accent Bar}',
        'name': 'accentCallout2',
        'avLst': {
            'adj1': 18750,
            'adj2': -8333,
            'adj3': 18750,
            'adj4': -16667,
            'adj5': 112500,
            'adj6': -46667
        }
    },
    LINE_CALLOUT_2_BORDER_AND_ACCENT_BAR: {
        'displayName': 'Line Callout 2 {Border and Accent Bar}',
        'name': 'accentBorderCallout2',
        'avLst': {
            'adj1': 18750,
            'adj2': -8333,
            'adj3': 18750,
            'adj4': -16667,
            'adj5': 112500,
            'adj6': -46667
        }
    },
    LINE_CALLOUT_2_NO_BORDER: {
        'displayName': 'Line Callout 2 {No Border}',
        'name': 'callout2',
        'avLst': {
            'adj1': 18750,
            'adj2': -8333,
            'adj3': 18750,
            'adj4': -16667,
            'adj5': 112500,
            'adj6': -46667
        }
    },
    LINE_CALLOUT_3: {
        'displayName': 'Line Callout 3',
        'name': 'borderCallout3',
        'avLst': {
            'adj1': 18750,
            'adj2': -8333,
            'adj3': 18750,
            'adj4': -16667,
            'adj5': 100000,
            'adj6': -16667,
            'adj7': 112963,
            'adj8': -8333
        }
    },
    LINE_CALLOUT_3_ACCENT_BAR: {
        'displayName': 'Line Callout 3 {Accent Bar}',
        'name': 'accentCallout3',
        'avLst': {
            'adj1': 18750,
            'adj2': -8333,
            'adj3': 18750,
            'adj4': -16667,
            'adj5': 100000,
            'adj6': -16667,
            'adj7': 112963,
            'adj8': -8333
        }
    },
    LINE_CALLOUT_3_BORDER_AND_ACCENT_BAR: {
        'displayName': 'Line Callout 3 {Border and Accent Bar}',
        'name': 'accentBorderCallout3',
        'avLst': {
            'adj1': 18750,
            'adj2': -8333,
            'adj3': 18750,
            'adj4': -16667,
            'adj5': 100000,
            'adj6': -16667,
            'adj7': 112963,
            'adj8': -8333
        }
    },
    LINE_CALLOUT_3_NO_BORDER: {
        'displayName': 'Line Callout 3 {No Border}',
        'name': 'callout3',
        'avLst': {
            'adj1': 18750,
            'adj2': -8333,
            'adj3': 18750,
            'adj4': -16667,
            'adj5': 100000,
            'adj6': -16667,
            'adj7': 112963,
            'adj8': -8333
        }
    },
    LINE_CALLOUT_4: {
        'displayName': 'Line Callout 3',
        'name': 'borderCallout3',
        'avLst': {
            'adj1': 18750,
            'adj2': -8333,
            'adj3': 18750,
            'adj4': -16667,
            'adj5': 100000,
            'adj6': -16667,
            'adj7': 112963,
            'adj8': -8333
        }
    },
    LINE_CALLOUT_4_ACCENT_BAR: {
        'displayName': 'Line Callout 3 {Accent Bar}',
        'name': 'accentCallout3',
        'avLst': {
            'adj1': 18750,
            'adj2': -8333,
            'adj3': 18750,
            'adj4': -16667,
            'adj5': 100000,
            'adj6': -16667,
            'adj7': 112963,
            'adj8': -8333
        }
    },
    LINE_CALLOUT_4_BORDER_AND_ACCENT_BAR: {
        'displayName': 'Line Callout 3 {Border and Accent Bar}',
        'name': 'accentBorderCallout3',
        'avLst': {
            'adj1': 18750,
            'adj2': -8333,
            'adj3': 18750,
            'adj4': -16667,
            'adj5': 100000,
            'adj6': -16667,
            'adj7': 112963,
            'adj8': -8333
        }
    },
    LINE_CALLOUT_4_NO_BORDER: {
        'displayName': 'Line Callout 3 {No Border}',
        'name': 'callout3',
        'avLst': {
            'adj1': 18750,
            'adj2': -8333,
            'adj3': 18750,
            'adj4': -16667,
            'adj5': 100000,
            'adj6': -16667,
            'adj7': 112963,
            'adj8': -8333
        }
    },
    LINE: {
        'displayName': 'Line',
        'name': 'line',
        'avLst': {}
    },
    LINE_INVERSE: {
        'displayName': 'Straight Connector',
        'name': 'lineInv',
        'avLst': {}
    },
    MATH_DIVIDE: {
        'displayName': 'Division',
        'name': 'mathDivide',
        'avLst': {
            'adj1': 23520,
            'adj2': 5880,
            'adj3': 11760
        }
    },
    MATH_EQUAL: {
        'displayName': 'Equal',
        'name': 'mathEqual',
        'avLst': {
            'adj1': 23520,
            'adj2': 11760
        }
    },
    MATH_MINUS: {
        'displayName': 'Minus',
        'name': 'mathMinus',
        'avLst': {
            'adj1': 23520
        }
    },
    MATH_MULTIPLY: {
        'displayName': 'Multiply',
        'name': 'mathMultiply',
        'avLst': {
            'adj1': 23520
        }
    },
    MATH_NOT_EQUAL: {
        'displayName': 'Not Equal',
        'name': 'mathNotEqual',
        'avLst': {
            'adj1': 23520,
            'adj2': 6600000,
            'adj3': 11760
        }
    },
    MATH_PLUS: {
        'displayName': 'Plus',
        'name': 'mathPlus',
        'avLst': {
            'adj1': 23520
        }
    },
    MOON: {
        'displayName': 'Moon',
        'name': 'moon',
        'avLst': {
            'adj': 50000
        }
    },
    NON_ISOSCELES_TRAPEZOID: {
        'displayName': 'Non-isosceles Trapezoid',
        'name': 'nonIsoscelesTrapezoid',
        'avLst': {
            'adj1': 25000,
            'adj2': 25000
        }
    },
    NOTCHED_RIGHT_ARROW: {
        'displayName': 'Notched Right Arrow',
        'name': 'notchedRightArrow',
        'avLst': {
            'adj1': 50000,
            'adj2': 50000
        }
    },
    NO_SYMBOL: {
        'displayName': '"No" symbol',
        'name': 'noSmoking',
        'avLst': {
            'adj': 18750
        }
    },
    OCTAGON: {
        'displayName': 'Octagon',
        'name': 'octagon',
        'avLst': {
            'adj': 29289
        }
    },
    OVAL: {
        'displayName': 'Oval',
        'name': 'ellipse',
        'avLst': {}
    },
    OVAL_CALLOUT: {
        'displayName': 'Oval Callout',
        'name': 'wedgeEllipseCallout',
        'avLst': {
            'adj1': -20833,
            'adj2': 62500
        }
    },
    PARALLELOGRAM: {
        'displayName': 'Parallelogram',
        'name': 'parallelogram',
        'avLst': {
            'adj': 25000
        }
    },
    PENTAGON: {
        'displayName': 'Pentagon',
        'name': 'homePlate',
        'avLst': {
            'adj': 50000
        }
    },
    PIE: {
        'displayName': 'Pie',
        'name': 'pie',
        'avLst': {
            'adj1': 0,
            'adj2': 16200000
        }
    },
    PIE_WEDGE: {
        'displayName': 'Pie',
        'name': 'pieWedge',
        'avLst': {}
    },
    PLAQUE: {
        'displayName': 'Plaque',
        'name': 'plaque',
        'avLst': {
            'adj': 16667
        }
    },
    PLAQUE_TABS: {
        'displayName': 'Plaque Tabs',
        'name': 'plaqueTabs',
        'avLst': {}
    },
    QUAD_ARROW: {
        'displayName': 'Quad Arrow',
        'name': 'quadArrow',
        'avLst': {
            'adj1': 22500,
            'adj2': 22500,
            'adj3': 22500
        }
    },
    QUAD_ARROW_CALLOUT: {
        'displayName': 'Quad Arrow Callout',
        'name': 'quadArrowCallout',
        'avLst': {
            'adj1': 18515,
            'adj2': 18515,
            'adj3': 18515,
            'adj4': 48123
        }
    },
    RECTANGLE: {
        'displayName': 'Rectangle',
        'name': 'rect',
        'avLst': {}
    },
    RECTANGULAR_CALLOUT: {
        'displayName': 'Rectangular Callout',
        'name': 'wedgeRectCallout',
        'avLst': {
            'adj1': -20833,
            'adj2': 62500
        }
    },
    REGULAR_PENTAGON: {
        'displayName': 'Regular Pentagon',
        'name': 'pentagon',
        'avLst': {
            'hf': 105146,
            'vf': 110557
        }
    },
    RIGHT_ARROW: {
        'displayName': 'Right Arrow',
        'name': 'rightArrow',
        'avLst': {
            'adj1': 50000,
            'adj2': 50000
        }
    },
    RIGHT_ARROW_CALLOUT: {
        'displayName': 'Right Arrow Callout',
        'name': 'rightArrowCallout',
        'avLst': {
            'adj1': 25000,
            'adj2': 25000,
            'adj3': 25000,
            'adj4': 64977
        }
    },
    RIGHT_BRACE: {
        'displayName': 'Right Brace',
        'name': 'rightBrace',
        'avLst': {
            'adj1': 8333,
            'adj2': 50000
        }
    },
    RIGHT_BRACKET: {
        'displayName': 'Right Bracket',
        'name': 'rightBracket',
        'avLst': {
            'adj': 8333
        }
    },
    RIGHT_TRIANGLE: {
        'displayName': 'Right Triangle',
        'name': 'rtTriangle',
        'avLst': {}
    },
    ROUNDED_RECTANGLE: {
        'displayName': 'Rounded Rectangle',
        'name': 'roundRect',
        'avLst': {
            'adj': 16667
        }
    },
    ROUNDED_RECTANGULAR_CALLOUT: {
        'displayName': 'Rounded Rectangular Callout',
        'name': 'wedgeRoundRectCallout',
        'avLst': {
            'adj1': -20833,
            'adj2': 62500,
            'adj3': 16667
        }
    },
    ROUND_1_RECTANGLE: {
        'displayName': 'Round Single Corner Rectangle',
        'name': 'round1Rect',
        'avLst': {
            'adj': 16667
        }
    },
    ROUND_2_DIAG_RECTANGLE: {
        'displayName': 'Round Diagonal Corner Rectangle',
        'name': 'round2DiagRect',
        'avLst': {
            'adj1': 16667,
            'adj2': 0
        }
    },
    ROUND_2_SAME_RECTANGLE: {
        'displayName': 'Round Same Side Corner Rectangle',
        'name': 'round2SameRect',
        'avLst': {
            'adj1': 16667,
            'adj2': 0
        }
    },
    SMILEY_FACE: {
        'displayName': 'Smiley Face',
        'name': 'smileyFace',
        'avLst': {
            'adj': 4653
        }
    },
    SNIP_1_RECTANGLE: {
        'displayName': 'Snip Single Corner Rectangle',
        'name': 'snip1Rect',
        'avLst': {
            'adj': 16667
        }
    },
    SNIP_2_DIAG_RECTANGLE: {
        'displayName': 'Snip Diagonal Corner Rectangle',
        'name': 'snip2DiagRect',
        'avLst': {
            'adj1': 0,
            'adj2': 16667
        }
    },
    SNIP_2_SAME_RECTANGLE: {
        'displayName': 'Snip Same Side Corner Rectangle',
        'name': 'snip2SameRect',
        'avLst': {
            'adj1': 16667,
            'adj2': 0
        }
    },
    SNIP_ROUND_RECTANGLE: {
        'displayName': 'Snip and Round Single Corner Rectangle',
        'name': 'snipRoundRect',
        'avLst': {
            'adj1': 16667,
            'adj2': 16667
        }
    },
    SQUARE_TABS: {
        'displayName': 'Square Tabs',
        'name': 'squareTabs',
        'avLst': {}
    },
    STAR_10_POINT: {
        'displayName': '10-Point Star',
        'name': 'star10',
        'avLst': {
            'adj': 42533,
            'hf': 105146
        }
    },
    STAR_12_POINT: {
        'displayName': '12-Point Star',
        'name': 'star12',
        'avLst': {
            'adj': 37500
        }
    },
    STAR_16_POINT: {
        'displayName': '16-Point Star',
        'name': 'star16',
        'avLst': {
            'adj': 37500
        }
    },
    STAR_24_POINT: {
        'displayName': '24-Point Star',
        'name': 'star24',
        'avLst': {
            'adj': 37500
        }
    },
    STAR_32_POINT: {
        'displayName': '32-Point Star',
        'name': 'star32',
        'avLst': {
            'adj': 37500
        }
    },
    STAR_4_POINT: {
        'displayName': '4-Point Star',
        'name': 'star4',
        'avLst': {
            'adj': 12500
        }
    },
    STAR_5_POINT: {
        'displayName': '5-Point Star',
        'name': 'star5',
        'avLst': {
            'adj': 19098,
            'hf': 105146,
            'vf': 110557
        }
    },
    STAR_6_POINT: {
        'displayName': '6-Point Star',
        'name': 'star6',
        'avLst': {
            'adj': 28868,
            'hf': 115470
        }
    },
    STAR_7_POINT: {
        'displayName': '7-Point Star',
        'name': 'star7',
        'avLst': {
            'adj': 34601,
            'hf': 102572,
            'vf': 105210
        }
    },
    STAR_8_POINT: {
        'displayName': '8-Point Star',
        'name': 'star8',
        'avLst': {
            'adj': 37500
        }
    },
    STRIPED_RIGHT_ARROW: {
        'displayName': 'Striped Right Arrow',
        'name': 'stripedRightArrow',
        'avLst': {
            'adj1': 50000,
            'adj2': 50000
        }
    },
    SUN: {
        'displayName': 'Sun',
        'name': 'sun',
        'avLst': {
            'adj': 25000
        }
    },
    SWOOSH_ARROW: {
        'displayName': 'Swoosh Arrow',
        'name': 'swooshArrow',
        'avLst': {
            'adj1': 25000,
            'adj2': 16667
        }
    },
    TEAR: {
        'displayName': 'Teardrop',
        'name': 'teardrop',
        'avLst': {
            'adj': 100000
        }
    },
    TRAPEZOID: {
        'displayName': 'Trapezoid',
        'name': 'trapezoid',
        'avLst': {
            'adj': 25000
        }
    },
    UP_ARROW: {
        'displayName': 'Up Arrow',
        'name': 'upArrow',
        'avLst': {}
    },
    UP_ARROW_CALLOUT: {
        'displayName': 'Up Arrow Callout',
        'name': 'upArrowCallout',
        'avLst': {
            'adj1': 25000,
            'adj2': 25000,
            'adj3': 25000,
            'adj4': 64977
        }
    },
    UP_DOWN_ARROW: {
        'displayName': 'Up-Down Arrow',
        'name': 'upDownArrow',
        'avLst': {
            'adj1': 50000,
            'adj2': 50000
        }
    },
    UP_DOWN_ARROW_CALLOUT: {
        'displayName': 'Up-Down Arrow Callout',
        'name': 'upDownArrowCallout',
        'avLst': {
            'adj1': 25000,
            'adj2': 25000,
            'adj3': 25000,
            'adj4': 48123
        }
    },
    UP_RIBBON: {
        'displayName': 'Up Ribbon',
        'name': 'ribbon2',
        'avLst': {
            'adj1': 16667,
            'adj2': 50000
        }
    },
    U_TURN_ARROW: {
        'displayName': 'U-Turn Arrow',
        'name': 'uturnArrow',
        'avLst': {
            'adj1': 25000,
            'adj2': 25000,
            'adj3': 25000,
            'adj4': 43750,
            'adj5': 75000
        }
    },
    VERTICAL_SCROLL: {
        'displayName': 'Vertical Scroll',
        'name': 'verticalScroll',
        'avLst': {
            'adj': 12500
        }
    },
    WAVE: {
        'displayName': 'Wave',
        'name': 'wave',
        'avLst': {
            'adj1': 12500,
            'adj2': 0
        }
    }
});

/**
 * PptxGenJS: Utility Methods
 */
/**
 * Convert string percentages to number relative to slide size
 * @param {number|string} size - numeric ("5.5") or percentage ("90%")
 * @param {'X' | 'Y'} xyDir - direction
 * @param {ILayout} layout - presentation layout
 * @returns {number} calculated size
 */
function getSmartParseNumber(size, xyDir, layout) {
    // FIRST: Convert string numeric value if reqd
    if (typeof size === 'string' && !isNaN(Number(size)))
        size = Number(size);
    // CASE 1: Number in inches
    // Assume any number less than 100 is inches
    if (typeof size === 'number' && size < 100)
        return inch2Emu(size);
    // CASE 2: Number is already converted to something other than inches
    // Assume any number greater than 100 is not inches! Just return it (its EMU already i guess??)
    if (typeof size === 'number' && size >= 100)
        return size;
    // CASE 3: Percentage (ex: '50%')
    if (typeof size === 'string' && size.indexOf('%') > -1) {
        if (xyDir && xyDir === 'X')
            return Math.round((parseFloat(size) / 100) * layout.width);
        if (xyDir && xyDir === 'Y')
            return Math.round((parseFloat(size) / 100) * layout.height);
        // Default: Assume width (x/cx)
        return Math.round((parseFloat(size) / 100) * layout.width);
    }
    // LAST: Default value
    return 0;
}
/**
 * Basic UUID Generator Adapted
 * @link https://stackoverflow.com/questions/105034/create-guid-uuid-in-javascript#answer-2117523
 * @param {string} uuidFormat - UUID format
 * @returns {string} UUID
 */
function getUuid(uuidFormat) {
    return uuidFormat.replace(/[xy]/g, function (c) {
        var r = (Math.random() * 16) | 0, v = c === 'x' ? r : (r & 0x3) | 0x8;
        return v.toString(16);
    });
}
/**
 * TODO: What does this mehtod do again??
 * shallow mix, returns new object
 */
function getMix(o1, o2, etc) {
    var objMix = {};
    var _loop_1 = function (i) {
        var oN = arguments_1[i];
        if (oN)
            Object.keys(oN).forEach(function (key) {
                objMix[key] = oN[key];
            });
    };
    var arguments_1 = arguments;
    for (var i = 0; i <= arguments.length; i++) {
        _loop_1(i);
    }
    return objMix;
}
/**
 * Replace special XML characters with HTML-encoded strings
 * @param {string} xml - XML string to encode
 * @returns {string} escaped XML
 */
function encodeXmlEntities(xml) {
    // NOTE: Dont use short-circuit eval here as value c/b "0" (zero) etc.!
    if (typeof xml === 'undefined' || xml == null)
        return '';
    return xml
        .toString()
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/\'/g, '&apos;');
}
/**
 * Convert inches into EMU
 * @param {number|string} inches - as string or number
 * @returns {number} EMU value
 */
function inch2Emu(inches) {
    // FIRST: Provide Caller Safety: Numbers may get conv<->conv during flight, so be kind and do some simple checks to ensure inches were passed
    // Any value over 100 damn sure isnt inches, must be EMU already, so just return it
    if (typeof inches === 'number' && inches > 100)
        return inches;
    if (typeof inches === 'string')
        inches = Number(inches.replace(/in*/gi, ''));
    return Math.round(EMU * inches);
}
/**
 * Convert degrees (0..360) to PowerPoint `rot` value
 *
 * @param {number} d - degrees
 * @returns {number} rot - value
 */
function convertRotationDegrees(d) {
    d = d || 0;
    return (d > 360 ? d - 360 : d) * 60000;
}
/**
 * Converts component value to hex value
 * @param {number} c - component color
 * @returns {string} hex string
 */
function componentToHex(c) {
    var hex = c.toString(16);
    return hex.length === 1 ? '0' + hex : hex;
}
/**
 * Converts RGB colors from css selectors to Hex for Presentation colors
 * @param {number} r - red value
 * @param {number} g - green value
 * @param {number} b - blue value
 * @returns {string} XML string
 */
function rgbToHex(r, g, b) {
    if (!Number.isInteger(r)) {
        try {
            console.warn('Integer expected!');
        }
        catch (ex) { }
    }
    return (componentToHex(r) + componentToHex(g) + componentToHex(b)).toUpperCase();
}
/**
 * Create either a `a:schemeClr` - (scheme color) or `a:srgbClr` (hexa representation).
 * @param {string} colorStr - hexa representation (eg. "FFFF00") or a scheme color constant (eg. pptx.colors.ACCENT1)
 * @param {string} innerElements - additional elements that adjust the color and are enclosed by the color element
 * @returns {string} XML string
 */
function createColorElement(colorStr, innerElements) {
    var isHexaRgb = REGEX_HEX_COLOR.test(colorStr);
    if (!isHexaRgb && Object.values(SCHEME_COLOR_NAMES).indexOf(colorStr) === -1) {
        console.warn('"' + colorStr + '" is not a valid scheme color or hexa RGB! "' + DEF_FONT_COLOR + '" is used as a fallback. Pass 6-digit RGB or `pptx.colors` values');
        colorStr = DEF_FONT_COLOR;
    }
    var tagName = isHexaRgb ? 'srgbClr' : 'schemeClr';
    var colorAttr = ' val="' + (isHexaRgb ? (colorStr || '').toUpperCase() : colorStr) + '"';
    return innerElements ? '<a:' + tagName + colorAttr + '>' + innerElements + '</a:' + tagName + '>' : '<a:' + tagName + colorAttr + '/>';
}
/**
 * Create color selection
 * @param {ShapeFill} shapeFill - options
 * @param {string} backColor - color string
 * @returns {string} XML string
 */
function genXmlColorSelection(shapeFill, backColor) {
    var colorVal = '';
    var fillType = 'solid';
    var internalElements = '';
    var outText = '';
    if (backColor && typeof backColor === 'string') {
        outText += "<p:bg><p:bgPr>" + genXmlColorSelection(backColor.replace('#', '')) + "<a:effectLst/></p:bgPr></p:bg>";
    }
    if (shapeFill) {
        if (typeof shapeFill === 'string')
            colorVal = shapeFill;
        else {
            if (shapeFill.type)
                fillType = shapeFill.type;
            if (shapeFill.color)
                colorVal = shapeFill.color;
            if (shapeFill.alpha)
                internalElements += '<a:alpha val="' + (100 - shapeFill.alpha) + '000"/>';
        }
        switch (fillType) {
            case 'solid':
                outText += '<a:solidFill>' + createColorElement(colorVal, internalElements) + '</a:solidFill>';
                break;
            default:
                break;
        }
    }
    return outText;
}

/**
 * PptxGenJS: Table Generation
 */
/**
 * Break text paragraphs into lines based upon table column width (e.g.: Magic Happens Here(tm))
 * @param {ITableCell} cell - table cell
 * @param {number} colWidth - table column width
 * @return {string[]} XML
 */
function parseTextToLines(cell, colWidth) {
    var CHAR = 2.2 + (cell.options && cell.options.autoPageCharWeight ? cell.options.autoPageCharWeight : 0); // Character Constant (An approximation of the Golden Ratio)
    var CPL = (colWidth * EMU) / (((cell.options && cell.options.fontSize) || DEF_FONT_SIZE) / CHAR); // Chars-Per-Line
    var arrLines = [];
    var strCurrLine = '';
    // A: Allow a single space/whitespace as cell text (user-requested feature)
    if (cell.text && cell.text.toString().trim().length === 0)
        return [' '];
    // B: Remove leading/trailing spaces
    var inStr = (cell.text || '').toString().trim();
    // C: Build line array
    inStr.split('\n').forEach(function (line) {
        line.split(' ').forEach(function (word) {
            if (strCurrLine.length + word.length + 1 < CPL) {
                strCurrLine += word + ' ';
            }
            else {
                if (strCurrLine)
                    arrLines.push(strCurrLine);
                strCurrLine = word + ' ';
            }
        });
        // All words for this line have been exhausted, flush buffer to new line, clear line var
        if (strCurrLine)
            arrLines.push(strCurrLine.trim() + CRLF);
        strCurrLine = '';
    });
    // D: Remove trailing linebreak
    arrLines[arrLines.length - 1] = arrLines[arrLines.length - 1].trim();
    return arrLines;
}
/**
 * Takes an array of table rows and breaks into an array of slides, which contain the calculated amount of table rows that fit on that slide
 * @param {[ITableToSlidesCell[]?]} tableRows - HTMLElementID of the table
 * @param {ITableToSlidesOpts} tabOpts - array of options (e.g.: tabsize)
 * @param {ILayout} presLayout - Presentation layout
 * @param {ISlideLayout} masterSlide - master slide (if any)
 * @return {TableRowSlide[]} array of table rows
 */
function getSlidesForTableRows(tableRows, tabOpts, presLayout, masterSlide) {
    if (tableRows === void 0) { tableRows = []; }
    if (tabOpts === void 0) { tabOpts = {}; }
    var arrInchMargins = DEF_SLIDE_MARGIN_IN, emuTabCurrH = 0, emuSlideTabW = EMU * 1, emuSlideTabH = EMU * 1, numCols = 0, tableRowSlides = [
        {
            rows: [],
        },
    ];
    if (tabOpts.verbose) {
        console.log("-- VERBOSE MODE ----------------------------------");
        console.log(".. (PARAMETERS)");
        console.log("presLayout.height ......... = " + presLayout.height / EMU);
        console.log("tabOpts.h ................. = " + tabOpts.h);
        console.log("tabOpts.w ................. = " + tabOpts.w);
        console.log("tabOpts.colW .............. = " + tabOpts.colW);
        console.log("tabOpts.slideMargin ....... = " + (tabOpts.slideMargin || ''));
        console.log(".. (/PARAMETERS)");
    }
    // STEP 1: Calculate margins
    {
        // Important: Use default size as zero cell margin is causing our tables to be too large and touch bottom of slide!
        if (!tabOpts.slideMargin && tabOpts.slideMargin !== 0)
            tabOpts.slideMargin = DEF_SLIDE_MARGIN_IN[0];
        if (masterSlide && typeof masterSlide.margin !== 'undefined') {
            if (Array.isArray(masterSlide.margin))
                arrInchMargins = masterSlide.margin;
            else if (!isNaN(Number(masterSlide.margin)))
                arrInchMargins = [Number(masterSlide.margin), Number(masterSlide.margin), Number(masterSlide.margin), Number(masterSlide.margin)];
        }
        else if (tabOpts.slideMargin || tabOpts.slideMargin === 0) {
            if (Array.isArray(tabOpts.slideMargin))
                arrInchMargins = tabOpts.slideMargin;
            else if (!isNaN(tabOpts.slideMargin))
                arrInchMargins = [tabOpts.slideMargin, tabOpts.slideMargin, tabOpts.slideMargin, tabOpts.slideMargin];
        }
        if (tabOpts.verbose)
            console.log('arrInchMargins ......... = ' + arrInchMargins.toString());
    }
    // STEP 2: Calculate number of columns
    {
        // NOTE: Cells may have a colspan, so merely taking the length of the [0] (or any other) row is not
        // ....: sufficient to determine column count. Therefore, check each cell for a colspan and total cols as reqd
        tableRows[0].forEach(function (cell) {
            if (!cell)
                cell = { type: SLIDE_OBJECT_TYPES.tablecell };
            var cellOpts = cell.options || null;
            numCols += Number(cellOpts && cellOpts.colspan ? cellOpts.colspan : 1);
        });
        if (tabOpts.verbose)
            console.log('numCols ................ = ' + numCols);
    }
    // STEP 3: Calculate tabOpts.w if tabOpts.colW was provided
    if (!tabOpts.w && tabOpts.colW) {
        if (Array.isArray(tabOpts.colW))
            tabOpts.colW.forEach(function (val) {
                typeof tabOpts.w !== 'number' ? (tabOpts.w = 0 + val) : (tabOpts.w += val);
            });
        else {
            tabOpts.w = tabOpts.colW * numCols;
        }
    }
    // STEP 4: Calculate usable space/table size (now that total usable space is known)
    {
        emuSlideTabW =
            typeof tabOpts.w === 'number'
                ? inch2Emu(tabOpts.w)
                : presLayout.width - inch2Emu((typeof tabOpts.x === 'number' ? tabOpts.x : arrInchMargins[1]) + arrInchMargins[3]);
        if (tabOpts.verbose)
            console.log('emuSlideTabW (in) ...... = ' + (emuSlideTabW / EMU).toFixed(1));
    }
    // STEP 5: Calculate column widths if not provided (emuSlideTabW will be used below to determine lines-per-col)
    if (!tabOpts.colW || !Array.isArray(tabOpts.colW)) {
        if (tabOpts.colW && !isNaN(Number(tabOpts.colW))) {
            var arrColW_1 = [];
            tableRows[0].forEach(function () {
                arrColW_1.push(tabOpts.colW);
            });
            tabOpts.colW = [];
            arrColW_1.forEach(function (val) {
                if (Array.isArray(tabOpts.colW))
                    tabOpts.colW.push(val);
            });
        }
        // No column widths provided? Then distribute cols.
        else {
            tabOpts.colW = [];
            for (var iCol = 0; iCol < numCols; iCol++) {
                tabOpts.colW.push(emuSlideTabW / EMU / numCols);
            }
        }
    }
    // STEP 6: **MAIN** Iterate over rows, add table content, create new slides as rows overflow
    var iRow = 0;
    var _loop_1 = function () {
        var row = tableRows.shift();
        iRow++;
        // A: Row variables
        var maxLineHeight = 0;
        var linesRow = [];
        var maxCellMarTopEmu = 0;
        var maxCellMarBtmEmu = 0;
        // B: Create new row in data model
        var currSlide = tableRowSlides[tableRowSlides.length - 1];
        var newRowSlide = [];
        row.forEach(function (cell) {
            newRowSlide.push({
                type: SLIDE_OBJECT_TYPES.tablecell,
                text: '',
                options: cell.options,
            });
            if (cell.options.margin && cell.options.margin[0] && cell.options.margin[0] * ONEPT > maxCellMarTopEmu)
                maxCellMarTopEmu = cell.options.margin[0] * ONEPT;
            else if (tabOpts.margin && tabOpts.margin[0] && tabOpts.margin[0] * ONEPT > maxCellMarTopEmu)
                maxCellMarTopEmu = tabOpts.margin[0] * ONEPT;
            if (cell.options.margin && cell.options.margin[2] && cell.options.margin[2] * ONEPT > maxCellMarBtmEmu)
                maxCellMarBtmEmu = cell.options.margin[2] * ONEPT;
            else if (tabOpts.margin && tabOpts.margin[2] && tabOpts.margin[2] * ONEPT > maxCellMarBtmEmu)
                maxCellMarBtmEmu = tabOpts.margin[2] * ONEPT;
        });
        // C: Calc usable vertical space/table height. Set default value first, adjust below when necessary.
        emuSlideTabH =
            tabOpts.h && typeof tabOpts.h === 'number'
                ? tabOpts.h
                : presLayout.height - inch2Emu(arrInchMargins[0] + arrInchMargins[2]) - (tabOpts.y && typeof tabOpts.y === 'number' ? tabOpts.y : 0);
        if (tabOpts.verbose)
            console.log('emuSlideTabH (in) ...... = ' + (emuSlideTabH / EMU).toFixed(1));
        // D: RULE: Use margins for starting point after the initial Slide, not `opt.y` (ISSUE#43, ISSUE#47, ISSUE#48)
        if (tableRowSlides.length > 1 && typeof tabOpts.newSlideStartY === 'number') {
            emuSlideTabH = tabOpts.h && typeof tabOpts.h === 'number' ? tabOpts.h : presLayout.height - inch2Emu(tabOpts.newSlideStartY + arrInchMargins[2]);
        }
        else if (tableRowSlides.length > 1 && typeof tabOpts.y === 'number') {
            emuSlideTabH = presLayout.height - inch2Emu((tabOpts.y / EMU < arrInchMargins[0] ? tabOpts.y / EMU : arrInchMargins[0]) + arrInchMargins[2]);
            // Use whichever is greater: area between margins or the table H provided (dont shrink usable area - the whole point of over-riding X on paging is to *increarse* usable space)
            if (typeof tabOpts.h === 'number' && emuSlideTabH < tabOpts.h)
                emuSlideTabH = tabOpts.h;
        }
        else if (typeof tabOpts.h === 'number' && typeof tabOpts.y === 'number')
            emuSlideTabH = tabOpts.h ? tabOpts.h : presLayout.height - inch2Emu((tabOpts.y / EMU || arrInchMargins[0]) + arrInchMargins[2]);
        //if (tabOpts.verbose) console.log(`- SLIDE [${tableRowSlides.length}]: emuSlideTabH .. = ${(emuSlideTabH / EMU).toFixed(1)}`)
        // E: **BUILD DATA SET** | Iterate over cells: split text into lines[], set `lineHeight`
        row.forEach(function (cell, iCell) {
            var newCell = {
                type: SLIDE_OBJECT_TYPES.tablecell,
                text: '',
                options: cell.options,
                lines: [],
                lineHeight: inch2Emu(((cell.options && cell.options.fontSize ? cell.options.fontSize : tabOpts.fontSize ? tabOpts.fontSize : DEF_FONT_SIZE) *
                    (LINEH_MODIFIER + (tabOpts.autoPageLineWeight ? tabOpts.autoPageLineWeight : 0))) /
                    100),
            };
            //if (tabOpts.verbose) console.log(`- CELL [${iCell}]: newCell.lineHeight ..... = ${(newCell.lineHeight / EMU).toFixed(2)}`)
            // 1: Exempt cells with `rowspan` from increasing lineHeight (or we could create a new slide when unecessary!)
            if (newCell.options.rowspan)
                newCell.lineHeight = 0;
            // 2: The parseTextToLines method uses `autoPageCharWeight`, so inherit from table options
            newCell.options.autoPageCharWeight = tabOpts.autoPageCharWeight ? tabOpts.autoPageCharWeight : null;
            // 3: **MAIN** Parse cell contents into lines based upon col width, font, etc
            newCell.lines = parseTextToLines(cell, tabOpts.colW[iCell] / ONEPT);
            // 4: Add to array
            linesRow.push(newCell);
        });
        // F: Start row height with margins
        if (tabOpts.verbose)
            console.log("- SLIDE [" + tableRowSlides.length + "]: ROW [" + iRow + "]: maxCellMarTopEmu=" + maxCellMarTopEmu + " / maxCellMarBtmEmu=" + maxCellMarBtmEmu);
        emuTabCurrH += maxCellMarTopEmu + maxCellMarBtmEmu;
        // G: Only create a new row if there is room, otherwise, it'll be an empty row as "A:" below will create a new Slide before loop can populate this row
        if (emuTabCurrH + maxLineHeight <= emuSlideTabH)
            currSlide.rows.push(newRowSlide);
        /* H: **PAGE DATA SET**
         * Add text one-line-a-time to this row's cells until: lines are exhausted OR table height limit is hit
         * Design: Building cells L-to-R/loop style wont work as one could be 100 lines and another 1 line.
         * Therefore, build the whole row, 1-line-at-a-time, spanning all columns.
         * That way, when the vertical size limit is hit, all lines pick up where they need to on the subsequent slide.
         */
        if (tabOpts.verbose)
            console.log("- SLIDE [" + tableRowSlides.length + "]: ROW [" + iRow + "]: START...");
        var _loop_2 = function () {
            // A: Add new Slide if there is no more space to fix 1 new line
            if (emuTabCurrH + maxLineHeight > emuSlideTabH) {
                if (tabOpts.verbose)
                    console.log("** NEW SLIDE CREATED *****************************************" +
                        (" (why?): " + (emuTabCurrH / EMU).toFixed(2) + "+" + (maxLineHeight / EMU).toFixed(2) + " > " + emuSlideTabH / EMU));
                // 1: Add a new slide
                tableRowSlides.push({
                    rows: [],
                });
                // 2: Reset current table height for new Slide
                emuTabCurrH = 0; // This row's emuRowH w/b added below
                // 3: Handle "addHeaderToEach" option /or/ Add new empty row to continue current lines into
                if (tabOpts.addHeaderToEach && tabOpts._arrObjTabHeadRows) {
                    // A: Add remaining cell lines
                    var newRowSlide_1 = [];
                    linesRow.forEach(function (cell) {
                        newRowSlide_1.push({
                            type: SLIDE_OBJECT_TYPES.tablecell,
                            text: cell.lines.join(''),
                            options: cell.options,
                        });
                    });
                    tableRows.unshift(newRowSlide_1);
                    // B: Add header row(s)
                    newRowSlide_1 = [];
                    tabOpts._arrObjTabHeadRows[0].forEach(function (cell) {
                        newRowSlide_1.push(cell);
                    });
                    tableRows.unshift(newRowSlide_1);
                    return "break";
                }
                else {
                    // A: Add new row to new slide
                    var currSlide_1 = tableRowSlides[tableRowSlides.length - 1];
                    var newRowSlide_2 = [];
                    row.forEach(function (cell) {
                        newRowSlide_2.push({
                            type: SLIDE_OBJECT_TYPES.tablecell,
                            text: '',
                            options: cell.options,
                        });
                    });
                    currSlide_1.rows.push(newRowSlide_2);
                }
            }
            // B: Add a line of text to 1-N cells that still have `lines`
            linesRow.forEach(function (cell, idx) {
                if (cell.lines.length > 0) {
                    // 1
                    var currSlide_2 = tableRowSlides[tableRowSlides.length - 1];
                    currSlide_2.rows[currSlide_2.rows.length - 1][idx].text +=
                        (currSlide_2.rows[currSlide_2.rows.length - 1][idx].text.length > 0 && !RegExp(/\n$/g).test(currSlide_2.rows[currSlide_2.rows.length - 1][idx].text)
                            ? CRLF
                            : '').replace(/[\r\n]+$/g, CRLF) + cell.lines.shift();
                    // 2
                    if (cell.lineHeight > maxLineHeight)
                        maxLineHeight = cell.lineHeight;
                }
            });
            // C: Increase table height by one line height as 1-N cells below are
            emuTabCurrH += maxLineHeight;
            if (tabOpts.verbose)
                console.log("- SLIDE [" + tableRowSlides.length + "]: ROW [" + iRow + "]: one line added ... emuTabCurrH = " + (emuTabCurrH / EMU).toFixed(2));
        };
        while (linesRow.filter(function (cell) {
            return cell.lines.length > 0;
        }).length > 0) {
            var state_1 = _loop_2();
            if (state_1 === "break")
                break;
        }
        if (tabOpts.verbose)
            console.log("- SLIDE [" + tableRowSlides.length + "]: ROW [" + iRow + "]: ...COMPLETE ...... emuTabCurrH = " + (emuTabCurrH / EMU).toFixed(2) + " ( emuSlideTabH = " + (emuSlideTabH / EMU).toFixed(2) + " )");
    };
    while (tableRows.length > 0) {
        _loop_1();
    }
    if (tabOpts.verbose) {
        console.log("\n|================================================|\n| FINAL: tableRowSlides.length = " + tableRowSlides.length);
        console.log(tableRowSlides);
        //console.log(JSON.stringify(tableRowSlides,null,2))
        console.log("|================================================|\n\n");
    }
    return tableRowSlides;
}
/**
 * Reproduces an HTML table as a PowerPoint table - including column widths, style, etc. - creates 1 or more slides as needed
 * @param {string} tabEleId - HTMLElementID of the table
 * @param {ITableToSlidesOpts} inOpts - array of options (e.g.: tabsize)
 */
function genTableToSlides(pptx, tabEleId, options, masterSlide) {
    if (options === void 0) { options = {}; }
    var opts = options || {};
    opts.slideMargin = opts.slideMargin || opts.slideMargin === 0 ? opts.slideMargin : 0.5;
    var emuSlideTabW = opts.w || pptx.presLayout.width;
    var arrObjTabHeadRows = [];
    var arrObjTabBodyRows = [];
    var arrObjTabFootRows = [];
    var arrColW = [];
    var arrTabColW = [];
    var arrInchMargins = [0.5, 0.5, 0.5, 0.5]; // TRBL-style
    var intTabW = 0;
    // REALITY-CHECK:
    if (!document.getElementById(tabEleId))
        throw 'tableToSlides: Table ID "' + tabEleId + '" does not exist!';
    // STEP 1: Set margins
    if (masterSlide && masterSlide.margin) {
        if (Array.isArray(masterSlide.margin))
            arrInchMargins = masterSlide.margin;
        else if (!isNaN(masterSlide.margin))
            arrInchMargins = [masterSlide.margin, masterSlide.margin, masterSlide.margin, masterSlide.margin];
        opts.slideMargin = arrInchMargins;
    }
    else if (opts && opts.slideMargin) {
        if (Array.isArray(opts.slideMargin))
            arrInchMargins = opts.slideMargin;
        else if (!isNaN(opts.slideMargin))
            arrInchMargins = [opts.slideMargin, opts.slideMargin, opts.slideMargin, opts.slideMargin];
    }
    emuSlideTabW = (opts.w ? inch2Emu(opts.w) : pptx.presLayout.width) - inch2Emu(arrInchMargins[1] + arrInchMargins[3]);
    if (opts.verbose)
        console.log('-- VERBOSE MODE ----------------------------------');
    if (opts.verbose)
        console.log("opts.h ................. = " + opts.h);
    if (opts.verbose)
        console.log("opts.w ................. = " + opts.w);
    if (opts.verbose)
        console.log("pptx.presLayout.width .. = " + pptx.presLayout.width / EMU);
    if (opts.verbose)
        console.log("emuSlideTabW (in)....... = " + emuSlideTabW / EMU);
    // STEP 2: Grab table col widths - just find the first availble row, either thead/tbody/tfoot, others may have colspsna,s who cares, we only need col widths from 1
    var firstRowCells = document.querySelectorAll("#" + tabEleId + " tr:first-child th");
    if (firstRowCells.length === 0)
        firstRowCells = document.querySelectorAll("#" + tabEleId + " tr:first-child td");
    firstRowCells.forEach(function (cell) {
        if (cell.getAttribute('colspan')) {
            // Guesstimate (divide evenly) col widths
            // NOTE: both j$query and vanilla selectors return {0} when table is not visible)
            for (var idx = 0; idx < Number(cell.getAttribute('colspan')); idx++) {
                arrTabColW.push(Math.round(cell.offsetWidth / Number(cell.getAttribute('colspan'))));
            }
        }
        else {
            arrTabColW.push(cell.offsetWidth);
        }
    });
    arrTabColW.forEach(function (colW) {
        intTabW += colW;
    });
    // STEP 3: Calc/Set column widths by using same column width percent from HTML table
    arrTabColW.forEach(function (colW, idx) {
        var intCalcWidth = Number(((Number(emuSlideTabW) * ((colW / intTabW) * 100)) / 100 / EMU).toFixed(2));
        var intMinWidth = Number(document.querySelector("#" + tabEleId + " thead tr:first-child th:nth-child(" + (idx + 1) + ")").getAttribute('data-pptx-min-width'));
        var intSetWidth = Number(document.querySelector("#" + tabEleId + " thead tr:first-child th:nth-child(" + (idx + 1) + ")").getAttribute('data-pptx-width'));
        arrColW.push(intSetWidth ? intSetWidth : intMinWidth > intCalcWidth ? intMinWidth : intCalcWidth);
    });
    if (opts.verbose) {
        console.log("arrColW ................ = " + arrColW.toString());
    }
    ['thead', 'tbody', 'tfoot'].forEach(function (part) {
        document.querySelectorAll("#" + tabEleId + " " + part + " tr").forEach(function (row) {
            var arrObjTabCells = [];
            Array.from(row.cells).forEach(function (cell) {
                // A: Get RGB text/bkgd colors
                var arrRGB1 = window
                    .getComputedStyle(cell)
                    .getPropertyValue('color')
                    .replace(/\s+/gi, '')
                    .replace('rgba(', '')
                    .replace('rgb(', '')
                    .replace(')', '')
                    .split(',');
                var arrRGB2 = window
                    .getComputedStyle(cell)
                    .getPropertyValue('background-color')
                    .replace(/\s+/gi, '')
                    .replace('rgba(', '')
                    .replace('rgb(', '')
                    .replace(')', '')
                    .split(',');
                if (
                // NOTE: (ISSUE#57): Default for unstyled tables is black bkgd, so use white instead
                window.getComputedStyle(cell).getPropertyValue('background-color') === 'rgba(0, 0, 0, 0)' ||
                    window.getComputedStyle(cell).getPropertyValue('transparent')) {
                    arrRGB2 = ['255', '255', '255'];
                }
                // B: Create option object
                var cellOpts = {
                    align: null,
                    bold: window.getComputedStyle(cell).getPropertyValue('font-weight') === 'bold' ||
                        Number(window.getComputedStyle(cell).getPropertyValue('font-weight')) >= 500
                        ? true
                        : false,
                    border: null,
                    color: rgbToHex(Number(arrRGB1[0]), Number(arrRGB1[1]), Number(arrRGB1[2])),
                    fill: rgbToHex(Number(arrRGB2[0]), Number(arrRGB2[1]), Number(arrRGB2[2])),
                    fontFace: (window.getComputedStyle(cell).getPropertyValue('font-family') || '')
                        .split(',')[0]
                        .replace(/"/g, '')
                        .replace('inherit', '')
                        .replace('initial', '') || null,
                    fontSize: Number(window
                        .getComputedStyle(cell)
                        .getPropertyValue('font-size')
                        .replace(/[a-z]/gi, '')),
                    margin: null,
                    colspan: Number(cell.getAttribute('colspan')) || null,
                    rowspan: Number(cell.getAttribute('rowspan')) || null,
                    valign: null,
                };
                if (['left', 'center', 'right', 'start', 'end'].indexOf(window.getComputedStyle(cell).getPropertyValue('text-align')) > -1) {
                    var align = window
                        .getComputedStyle(cell)
                        .getPropertyValue('text-align')
                        .replace('start', 'left')
                        .replace('end', 'right');
                    cellOpts.align = align === 'center' ? 'center' : align === 'left' ? 'left' : align === 'right' ? 'right' : null;
                }
                if (['top', 'middle', 'bottom'].indexOf(window.getComputedStyle(cell).getPropertyValue('vertical-align')) > -1) {
                    var valign = window.getComputedStyle(cell).getPropertyValue('vertical-align');
                    cellOpts.valign = valign === 'top' ? 'top' : valign === 'middle' ? 'middle' : valign === 'bottom' ? 'bottom' : null;
                }
                // C: Add padding [margin] (if any)
                // NOTE: Margins translate: px->pt 1:1 (e.g.: a 20px padded cell looks the same in PPTX as 20pt Text Inset/Padding)
                if (window.getComputedStyle(cell).getPropertyValue('padding-left')) {
                    cellOpts.margin = [0, 0, 0, 0];
                    new Array('padding-top', 'padding-right', 'padding-bottom', 'padding-left').forEach(function (val, idx) {
                        cellOpts.margin[idx] = Math.round(Number(window
                            .getComputedStyle(cell)
                            .getPropertyValue(val)
                            .replace(/\D/gi, '')));
                    });
                }
                // D: Add border (if any)
                if (window.getComputedStyle(cell).getPropertyValue('border-top-width') ||
                    window.getComputedStyle(cell).getPropertyValue('border-right-width') ||
                    window.getComputedStyle(cell).getPropertyValue('border-bottom-width') ||
                    window.getComputedStyle(cell).getPropertyValue('border-left-width')) {
                    cellOpts.border = [null, null, null, null];
                    new Array('top', 'right', 'bottom', 'left').forEach(function (val, idx) {
                        var intBorderW = Math.round(Number(window
                            .getComputedStyle(cell)
                            .getPropertyValue('border-' + val + '-width')
                            .replace('px', '')));
                        var arrRGB = [];
                        arrRGB = window
                            .getComputedStyle(cell)
                            .getPropertyValue('border-' + val + '-color')
                            .replace(/\s+/gi, '')
                            .replace('rgba(', '')
                            .replace('rgb(', '')
                            .replace(')', '')
                            .split(',');
                        var strBorderC = rgbToHex(Number(arrRGB[0]), Number(arrRGB[1]), Number(arrRGB[2]));
                        cellOpts.border[idx] = { pt: intBorderW, color: strBorderC };
                    });
                }
                // LAST: Add cell
                arrObjTabCells.push({
                    type: SLIDE_OBJECT_TYPES.tablecell,
                    text: cell.innerText,
                    options: cellOpts,
                });
            });
            switch (part) {
                case 'thead':
                    arrObjTabHeadRows.push(arrObjTabCells);
                    break;
                case 'tbody':
                    arrObjTabBodyRows.push(arrObjTabCells);
                    break;
                case 'tfoot':
                    arrObjTabFootRows.push(arrObjTabCells);
                    break;
                default:
            }
        });
    });
    // STEP 5: Break table into Slides as needed
    // Pass head-rows as there is an option to add to each table and the parse func needs this data to fulfill that option
    opts._arrObjTabHeadRows = arrObjTabHeadRows || null;
    opts.colW = arrColW;
    getSlidesForTableRows(arrObjTabHeadRows.concat(arrObjTabBodyRows, arrObjTabFootRows), opts, pptx.presLayout, masterSlide).forEach(function (slide, idx) {
        // A: Create new Slide
        var newSlide = pptx.addSlide(opts.masterSlideName || null);
        // B: DESIGN: Reset `y` to `newSlideStartY` or margin after first Slide (ISSUE#43, ISSUE#47, ISSUE#48)
        if (idx === 0)
            opts.y = opts.y || arrInchMargins[0];
        if (idx > 0)
            opts.y = opts.newSlideStartY || arrInchMargins[0];
        if (opts.verbose)
            console.log('opts.newSlideStartY:' + opts.newSlideStartY + ' / arrInchMargins[0]:' + arrInchMargins[0] + ' => opts.y = ' + opts.y);
        // C: Add table to Slide
        newSlide.addTable(slide.rows, { x: opts.x || arrInchMargins[3], y: opts.y, w: Number(emuSlideTabW) / EMU, colW: arrColW, autoPage: false });
        // D: Add any additional objects
        if (opts.addImage)
            newSlide.addImage({ path: opts.addImage.url, x: opts.addImage.x, y: opts.addImage.y, w: opts.addImage.w, h: opts.addImage.h });
        if (opts.addShape)
            newSlide.addShape(opts.addShape.shape, opts.addShape.opts || {});
        if (opts.addTable)
            newSlide.addTable(opts.addTable.rows, opts.addTable.opts || {});
        if (opts.addText)
            newSlide.addText(opts.addText.text, opts.addText.opts || {});
    });
}

/**
 * PptxGenJS: XML Generation
 */
var imageSizingXml = {
    cover: function (imgSize, boxDim) {
        var imgRatio = imgSize.h / imgSize.w, boxRatio = boxDim.h / boxDim.w, isBoxBased = boxRatio > imgRatio, width = isBoxBased ? boxDim.h / imgRatio : boxDim.w, height = isBoxBased ? boxDim.h : boxDim.w * imgRatio, hzPerc = Math.round(1e5 * 0.5 * (1 - boxDim.w / width)), vzPerc = Math.round(1e5 * 0.5 * (1 - boxDim.h / height));
        return '<a:srcRect l="' + hzPerc + '" r="' + hzPerc + '" t="' + vzPerc + '" b="' + vzPerc + '"/><a:stretch/>';
    },
    contain: function (imgSize, boxDim) {
        var imgRatio = imgSize.h / imgSize.w, boxRatio = boxDim.h / boxDim.w, widthBased = boxRatio > imgRatio, width = widthBased ? boxDim.w : boxDim.h / imgRatio, height = widthBased ? boxDim.w * imgRatio : boxDim.h, hzPerc = Math.round(1e5 * 0.5 * (1 - boxDim.w / width)), vzPerc = Math.round(1e5 * 0.5 * (1 - boxDim.h / height));
        return '<a:srcRect l="' + hzPerc + '" r="' + hzPerc + '" t="' + vzPerc + '" b="' + vzPerc + '"/><a:stretch/>';
    },
    crop: function (imageSize, boxDim) {
        var l = boxDim.x, r = imageSize.w - (boxDim.x + boxDim.w), t = boxDim.y, b = imageSize.h - (boxDim.y + boxDim.h), lPerc = Math.round(1e5 * (l / imageSize.w)), rPerc = Math.round(1e5 * (r / imageSize.w)), tPerc = Math.round(1e5 * (t / imageSize.h)), bPerc = Math.round(1e5 * (b / imageSize.h));
        return '<a:srcRect l="' + lPerc + '" r="' + rPerc + '" t="' + tPerc + '" b="' + bPerc + '"/><a:stretch/>';
    },
};
/**
 * Transforms a slide or slideLayout to resulting XML string - Creates `ppt/slide*.xml`
 * @param {ISlide|ISlideLayout} slideObject - slide object created within createSlideObject
 * @return {string} XML string with <p:cSld> as the root
 */
function slideObjectToXml(slide) {
    var strSlideXml = slide.name ? '<p:cSld name="' + slide.name + '">' : '<p:cSld>';
    var intTableNum = 1;
    // STEP 1: Add background
    if (slide.bkgd) {
        strSlideXml += genXmlColorSelection(null, slide.bkgd);
    }
    else if (!slide.bkgd && slide.name && slide.name === DEF_PRES_LAYOUT_NAME) {
        // NOTE: Default [white] background is needed on slideMaster1.xml to avoid gray background in Keynote (and Finder previews)
        strSlideXml += '<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>';
    }
    // STEP 2: Add background image (using Strech) (if any)
    if (slide.bkgdImgRid) {
        // FIXME: We should be doing this in the slideLayout...
        strSlideXml +=
            '<p:bg>' +
                '<p:bgPr><a:blipFill dpi="0" rotWithShape="1">' +
                '<a:blip r:embed="rId' +
                slide.bkgdImgRid +
                '"><a:lum/></a:blip>' +
                '<a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill>' +
                '<a:effectLst/></p:bgPr>' +
                '</p:bg>';
    }
    // STEP 3: Continue slide by starting spTree node
    strSlideXml += '<p:spTree>';
    strSlideXml += '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>';
    strSlideXml += '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>';
    strSlideXml += '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';
    // STEP 4: Loop over all Slide.data objects and add them to this slide
    slide.data.forEach(function (slideItemObj, idx) {
        var x = 0, y = 0, cx = getSmartParseNumber('75%', 'X', slide.presLayout), cy = 0;
        var placeholderObj;
        var locationAttr = '';
        var shapeType = null;
        if (slide.slideLayout !== undefined && slide.slideLayout.data !== undefined && slideItemObj.options && slideItemObj.options.placeholder) {
            placeholderObj = slide['slideLayout']['data'].filter(function (object) {
                return object.options.placeholder === slideItemObj.options.placeholder;
            })[0];
        }
        // A: Set option vars
        slideItemObj.options = slideItemObj.options || {};
        if (typeof slideItemObj.options.x !== 'undefined')
            x = getSmartParseNumber(slideItemObj.options.x, 'X', slide.presLayout);
        if (typeof slideItemObj.options.y !== 'undefined')
            y = getSmartParseNumber(slideItemObj.options.y, 'Y', slide.presLayout);
        if (typeof slideItemObj.options.w !== 'undefined')
            cx = getSmartParseNumber(slideItemObj.options.w, 'X', slide.presLayout);
        if (typeof slideItemObj.options.h !== 'undefined')
            cy = getSmartParseNumber(slideItemObj.options.h, 'Y', slide.presLayout);
        // If using a placeholder then inherit it's position
        if (placeholderObj) {
            if (placeholderObj.options.x || placeholderObj.options.x === 0)
                x = getSmartParseNumber(placeholderObj.options.x, 'X', slide.presLayout);
            if (placeholderObj.options.y || placeholderObj.options.y === 0)
                y = getSmartParseNumber(placeholderObj.options.y, 'Y', slide.presLayout);
            if (placeholderObj.options.w || placeholderObj.options.w === 0)
                cx = getSmartParseNumber(placeholderObj.options.w, 'X', slide.presLayout);
            if (placeholderObj.options.h || placeholderObj.options.h === 0)
                cy = getSmartParseNumber(placeholderObj.options.h, 'Y', slide.presLayout);
        }
        //
        if (slideItemObj.shape)
            shapeType = getShapeInfo(slideItemObj.shape);
        //
        if (slideItemObj.options.flipH)
            locationAttr += ' flipH="1"';
        if (slideItemObj.options.flipV)
            locationAttr += ' flipV="1"';
        if (slideItemObj.options.rotate)
            locationAttr += ' rot="' + convertRotationDegrees(slideItemObj.options.rotate) + '"';
        // B: Add OBJECT to the current Slide
        switch (slideItemObj.type) {
            case SLIDE_OBJECT_TYPES.table:
                var objTableGrid_1 = {};
                var arrTabRows_1 = slideItemObj.arrTabRows;
                var objTabOpts_1 = slideItemObj.options;
                var intColCnt_1 = 0, intColW = 0;
                var cellOpts_1;
                // Calc number of columns
                // NOTE: Cells may have a colspan, so merely taking the length of the [0] (or any other) row is not
                // ....: sufficient to determine column count. Therefore, check each cell for a colspan and total cols as reqd
                arrTabRows_1[0].forEach(function (cell) {
                    cellOpts_1 = cell.options || null;
                    intColCnt_1 += cellOpts_1 && cellOpts_1.colspan ? Number(cellOpts_1.colspan) : 1;
                });
                // STEP 1: Start Table XML
                // NOTE: Non-numeric cNvPr id values will trigger "presentation needs repair" type warning in MS-PPT-2013
                var strXml_1 = '<p:graphicFrame>' +
                    '  <p:nvGraphicFramePr>' +
                    '    <p:cNvPr id="' +
                    (intTableNum * slide.number + 1) +
                    '" name="Table ' +
                    intTableNum * slide.number +
                    '"/>' +
                    '    <p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>' +
                    '    <p:nvPr><p:extLst><p:ext uri="{D42A27DB-BD31-4B8C-83A1-F6EECF244321}"><p14:modId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="1579011935"/></p:ext></p:extLst></p:nvPr>' +
                    '  </p:nvGraphicFramePr>' +
                    '  <p:xfrm>' +
                    '    <a:off x="' +
                    (x || (x === 0 ? 0 : EMU)) +
                    '" y="' +
                    (y || (y === 0 ? 0 : EMU)) +
                    '"/>' +
                    '    <a:ext cx="' +
                    (cx || (cx === 0 ? 0 : EMU)) +
                    '" cy="' +
                    (cy || EMU) +
                    '"/>' +
                    '  </p:xfrm>' +
                    '  <a:graphic>' +
                    '    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">' +
                    '      <a:tbl>' +
                    '        <a:tblPr/>';
                // + '        <a:tblPr bandRow="1"/>';
                // TODO: Support banded rows, first/last row, etc.
                // NOTE: Banding, etc. only shows when using a table style! (or set alt row color if banding)
                // <a:tblPr firstCol="0" firstRow="0" lastCol="0" lastRow="0" bandCol="0" bandRow="1">
                // STEP 2: Set column widths
                // Evenly distribute cols/rows across size provided when applicable (calc them if only overall dimensions were provided)
                // A: Col widths provided?
                if (Array.isArray(objTabOpts_1.colW)) {
                    strXml_1 += '<a:tblGrid>';
                    for (var col = 0; col < intColCnt_1; col++) {
                        strXml_1 +=
                            '<a:gridCol w="' +
                                Math.round(inch2Emu(objTabOpts_1.colW[col]) || (typeof slideItemObj.options.w === 'number' ? slideItemObj.options.w : 1) / intColCnt_1) +
                                '"/>';
                    }
                    strXml_1 += '</a:tblGrid>';
                }
                // B: Table Width provided without colW? Then distribute cols
                else {
                    intColW = objTabOpts_1.colW ? objTabOpts_1.colW : EMU;
                    if (slideItemObj.options.w && !objTabOpts_1.colW)
                        intColW = Math.round((typeof slideItemObj.options.w === 'number' ? slideItemObj.options.w : 1) / intColCnt_1);
                    strXml_1 += '<a:tblGrid>';
                    for (var col = 0; col < intColCnt_1; col++) {
                        strXml_1 += '<a:gridCol w="' + intColW + '"/>';
                    }
                    strXml_1 += '</a:tblGrid>';
                }
                // STEP 3: Build our row arrays into an actual grid to match the XML we will be building next (ISSUE #36)
                // Note row arrays can arrive "lopsided" as in row1:[1,2,3] row2:[3] when first two cols rowspan!,
                // so a simple loop below in XML building wont suffice to build table correctly.
                // We have to build an actual grid now
                /*
                    EX: (A0:rowspan=3, B1:rowspan=2, C1:colspan=2)

                    /------|------|------|------\
                    |  A0  |  B0  |  C0  |  D0  |
                    |      |  B1  |  C1  |      |
                    |      |      |  C2  |  D2  |
                    \------|------|------|------/
                */
                /*
                    Object ex: key = rowIdx / val = [cells] cellIdx { 0:{type: "tablecell", text: Array(1), options: {…}}, 1:... }
                    {0: {…}, 1: {…}, 2: {…}, 3: {…}}
                */
                arrTabRows_1.forEach(function (row, rIdx) {
                    // A: Create row if needed (recall one may be created in loop below for rowspans, so dont assume we need to create one each iteration)
                    if (!objTableGrid_1[rIdx])
                        objTableGrid_1[rIdx] = {};
                    // B: Loop over all cells
                    row.forEach(function (cell, cIdx) {
                        // DESIGN: NOTE: Row cell arrays can be "uneven" (diff cell count in each) due to rowspan/colspan
                        // Therefore, for each cell we run 0->colCount to determine the correct slot for it to reside
                        // as the uneven/mixed nature of the data means we cannot use the cIdx value alone.
                        // E.g.: the 2nd element in the row array may actually go into the 5th table grid row cell b/c of colspans!
                        for (var idx_1 = 0; cIdx + idx_1 < intColCnt_1; idx_1++) {
                            var currColIdx = cIdx + idx_1;
                            if (!objTableGrid_1[rIdx][currColIdx]) {
                                // A: Set this cell
                                objTableGrid_1[rIdx][currColIdx] = cell;
                                // B: Handle `colspan` or `rowspan` (a {cell} cant have both! TODO: FUTURE: ROWSPAN & COLSPAN in same cell)
                                if (cell && cell.options && cell.options.colspan && !isNaN(Number(cell.options.colspan))) {
                                    for (var idy = 1; idy < Number(cell.options.colspan); idy++) {
                                        objTableGrid_1[rIdx][currColIdx + idy] = { hmerge: true, text: 'hmerge' };
                                    }
                                }
                                else if (cell && cell.options && cell.options.rowspan && !isNaN(Number(cell.options.rowspan))) {
                                    for (var idz = 1; idz < Number(cell.options.rowspan); idz++) {
                                        if (!objTableGrid_1[rIdx + idz])
                                            objTableGrid_1[rIdx + idz] = {};
                                        objTableGrid_1[rIdx + idz][currColIdx] = { vmerge: true, text: 'vmerge' };
                                    }
                                }
                                // C: Break out of colCnt loop now that slot has been filled
                                break;
                            }
                        }
                    });
                });
                /* DEBUG: useful for rowspan/colspan testing
                if ( objTabOpts.verbose ) {
                    console.table(objTableGrid);
                    let arrText = [];
                    objTableGrid.forEach(function(row){ let arrRow = []; row.forEach(row,function(cell){ arrRow.push(cell.text); }); arrText.push(arrRow); });
                    console.table( arrText );
                }
                */
                // STEP 4: Build table rows/cells
                Object.entries(objTableGrid_1).forEach(function (_a) {
                    var rIdx = _a[0], rowObj = _a[1];
                    // A: Table Height provided without rowH? Then distribute rows
                    var intRowH = 0; // IMPORTANT: Default must be zero for auto-sizing to work
                    if (Array.isArray(objTabOpts_1.rowH) && objTabOpts_1.rowH[rIdx])
                        intRowH = inch2Emu(Number(objTabOpts_1.rowH[rIdx]));
                    else if (objTabOpts_1.rowH && !isNaN(Number(objTabOpts_1.rowH)))
                        intRowH = inch2Emu(Number(objTabOpts_1.rowH));
                    else if (slideItemObj.options.cy || slideItemObj.options.h)
                        intRowH =
                            (slideItemObj.options.h ? inch2Emu(slideItemObj.options.h) : typeof slideItemObj.options.cy === 'number' ? slideItemObj.options.cy : 1) /
                                arrTabRows_1.length;
                    // B: Start row
                    strXml_1 += '<a:tr h="' + intRowH + '">';
                    // C: Loop over each CELL
                    Object.entries(rowObj).forEach(function (_a) {
                        var _cIdx = _a[0], cellObj = _a[1];
                        var cell = cellObj;
                        // 1: "hmerge" cells are just place-holders in the table grid - skip those and go to next cell
                        if (cell.hmerge)
                            return;
                        // 2: OPTIONS: Build/set cell options
                        var cellOpts = cell.options || {};
                        cell.options = cellOpts;
                        ['align', 'bold', 'border', 'color', 'fill', 'fontFace', 'fontSize', 'margin', 'underline', 'valign'].forEach(function (name) {
                            if (objTabOpts_1[name] && !cellOpts[name] && cellOpts[name] !== 0)
                                cellOpts[name] = objTabOpts_1[name];
                        });
                        var cellValign = cellOpts.valign
                            ? ' anchor="' +
                                cellOpts.valign
                                    .replace(/^c$/i, 'ctr')
                                    .replace(/^m$/i, 'ctr')
                                    .replace('center', 'ctr')
                                    .replace('middle', 'ctr')
                                    .replace('top', 't')
                                    .replace('btm', 'b')
                                    .replace('bottom', 'b') +
                                '"'
                            : '';
                        var cellColspan = cellOpts.colspan ? ' gridSpan="' + cellOpts.colspan + '"' : '';
                        var cellRowspan = cellOpts.rowspan ? ' rowSpan="' + cellOpts.rowspan + '"' : '';
                        var cellFill = (cell.optImp && cell.optImp.fill) || cellOpts.fill
                            ? ' <a:solidFill><a:srgbClr val="' +
                                ((cell.optImp && cell.optImp.fill) || (typeof cellOpts.fill === 'string' ? cellOpts.fill.replace('#', '') : '')).toUpperCase() +
                                '"/></a:solidFill>'
                            : '';
                        var cellMargin = cellOpts.margin === 0 || cellOpts.margin ? cellOpts.margin : DEF_CELL_MARGIN_PT;
                        if (!Array.isArray(cellMargin) && typeof cellMargin === 'number')
                            cellMargin = [cellMargin, cellMargin, cellMargin, cellMargin];
                        var cellMarginXml = ' marL="' +
                            cellMargin[3] * ONEPT +
                            '" marR="' +
                            cellMargin[1] * ONEPT +
                            '" marT="' +
                            cellMargin[0] * ONEPT +
                            '" marB="' +
                            cellMargin[2] * ONEPT +
                            '"';
                        // TODO: Cell NOWRAP property (text wrap: add to a:tcPr (horzOverflow="overflow" or whatever options exist)
                        // 3: ROWSPAN: Add dummy cells for any active rowspan
                        if (cell.vmerge) {
                            strXml_1 += '<a:tc vMerge="1"><a:tcPr/></a:tc>';
                            return;
                        }
                        // 4: Set CELL content and properties ==================================
                        strXml_1 += '<a:tc' + cellColspan + cellRowspan + '>' + genXmlTextBody(cell) + '<a:tcPr' + cellMarginXml + cellValign + '>';
                        // 5: Borders: Add any borders
                        if (cellOpts.border && !Array.isArray(cellOpts.border) && cellOpts.border.type === 'none') {
                            strXml_1 += '  <a:lnL w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnL>';
                            strXml_1 += '  <a:lnR w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnR>';
                            strXml_1 += '  <a:lnT w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnT>';
                            strXml_1 += '  <a:lnB w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnB>';
                        }
                        else if (cellOpts.border && typeof cellOpts.border === 'string') {
                            strXml_1 +=
                                '  <a:lnL w="' + ONEPT + '" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="' + cellOpts.border + '"/></a:solidFill></a:lnL>';
                            strXml_1 +=
                                '  <a:lnR w="' + ONEPT + '" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="' + cellOpts.border + '"/></a:solidFill></a:lnR>';
                            strXml_1 +=
                                '  <a:lnT w="' + ONEPT + '" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="' + cellOpts.border + '"/></a:solidFill></a:lnT>';
                            strXml_1 +=
                                '  <a:lnB w="' + ONEPT + '" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="' + cellOpts.border + '"/></a:solidFill></a:lnB>';
                        }
                        else if (cellOpts.border && Array.isArray(cellOpts.border)) {
                            [{ idx: 3, name: 'lnL' }, { idx: 1, name: 'lnR' }, { idx: 0, name: 'lnT' }, { idx: 2, name: 'lnB' }].forEach(function (obj) {
                                if (cellOpts.border[obj.idx]) {
                                    var strC = '<a:solidFill><a:srgbClr val="' +
                                        (cellOpts.border[obj.idx].color ? cellOpts.border[obj.idx].color : DEF_CELL_BORDER.color) +
                                        '"/></a:solidFill>';
                                    var intW = cellOpts.border[obj.idx] && (cellOpts.border[obj.idx].pt || cellOpts.border[obj.idx].pt === 0)
                                        ? ONEPT * Number(cellOpts.border[obj.idx].pt)
                                        : ONEPT;
                                    strXml_1 += '<a:' + obj.name + ' w="' + intW + '" cap="flat" cmpd="sng" algn="ctr">' + strC + '</a:' + obj.name + '>';
                                }
                                else
                                    strXml_1 += '<a:' + obj.name + ' w="0"><a:miter lim="400000"/></a:' + obj.name + '>';
                            });
                        }
                        else if (cellOpts.border && !Array.isArray(cellOpts.border)) {
                            var intW = cellOpts.border && (cellOpts.border.pt || cellOpts.border.pt === 0) ? ONEPT * Number(cellOpts.border.pt) : ONEPT;
                            var strClr = '<a:solidFill><a:srgbClr val="' +
                                (cellOpts.border.color ? cellOpts.border.color.replace('#', '') : DEF_CELL_BORDER.color) +
                                '"/></a:solidFill>';
                            var strAttr = '<a:prstDash val="';
                            strAttr += cellOpts.border.type && cellOpts.border.type.toLowerCase().indexOf('dash') > -1 ? 'sysDash' : 'solid';
                            strAttr += '"/><a:round/><a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/>';
                            // *** IMPORTANT! *** LRTB order matters! (Reorder a line below to watch the borders go wonky in MS-PPT-2013!!)
                            strXml_1 += '<a:lnL w="' + intW + '" cap="flat" cmpd="sng" algn="ctr">' + strClr + strAttr + '</a:lnL>';
                            strXml_1 += '<a:lnR w="' + intW + '" cap="flat" cmpd="sng" algn="ctr">' + strClr + strAttr + '</a:lnR>';
                            strXml_1 += '<a:lnT w="' + intW + '" cap="flat" cmpd="sng" algn="ctr">' + strClr + strAttr + '</a:lnT>';
                            strXml_1 += '<a:lnB w="' + intW + '" cap="flat" cmpd="sng" algn="ctr">' + strClr + strAttr + '</a:lnB>';
                            // *** IMPORTANT! *** LRTB order matters!
                        }
                        // 6: Close cell Properties & Cell
                        strXml_1 += cellFill;
                        strXml_1 += '  </a:tcPr>';
                        strXml_1 += ' </a:tc>';
                        // LAST: COLSPAN: Add a 'merged' col for each column being merged (SEE: http://officeopenxml.com/drwTableGrid.php)
                        if (cellOpts.colspan) {
                            for (var tmp = 1; tmp < Number(cellOpts.colspan); tmp++) {
                                strXml_1 += '<a:tc hMerge="1"><a:tcPr/></a:tc>';
                            }
                        }
                    });
                    // D: Complete row
                    strXml_1 += '</a:tr>';
                });
                // STEP 5: Complete table
                strXml_1 += '      </a:tbl>';
                strXml_1 += '    </a:graphicData>';
                strXml_1 += '  </a:graphic>';
                strXml_1 += '</p:graphicFrame>';
                // STEP 6: Set table XML
                strSlideXml += strXml_1;
                // LAST: Increment counter
                intTableNum++;
                break;
            case SLIDE_OBJECT_TYPES.text:
            case SLIDE_OBJECT_TYPES.placeholder:
                // Lines can have zero cy, but text should not
                if (!slideItemObj.options.line && cy === 0)
                    cy = EMU * 0.3;
                // Margin/Padding/Inset for textboxes
                if (slideItemObj.options.margin && Array.isArray(slideItemObj.options.margin)) {
                    slideItemObj.options.bodyProp.lIns = slideItemObj.options.margin[0] * ONEPT || 0;
                    slideItemObj.options.bodyProp.rIns = slideItemObj.options.margin[1] * ONEPT || 0;
                    slideItemObj.options.bodyProp.bIns = slideItemObj.options.margin[2] * ONEPT || 0;
                    slideItemObj.options.bodyProp.tIns = slideItemObj.options.margin[3] * ONEPT || 0;
                }
                else if (typeof slideItemObj.options.margin === 'number') {
                    slideItemObj.options.bodyProp.lIns = slideItemObj.options.margin * ONEPT;
                    slideItemObj.options.bodyProp.rIns = slideItemObj.options.margin * ONEPT;
                    slideItemObj.options.bodyProp.bIns = slideItemObj.options.margin * ONEPT;
                    slideItemObj.options.bodyProp.tIns = slideItemObj.options.margin * ONEPT;
                }
                if (shapeType === null)
                    shapeType = getShapeInfo(null);
                // A: Start SHAPE =======================================================
                strSlideXml += '<p:sp>';
                // B: The addition of the "txBox" attribute is the sole determiner of if an object is a shape or textbox
                strSlideXml += '<p:nvSpPr><p:cNvPr id="' + (idx + 2) + '" name="Object ' + (idx + 1) + '"/>';
                strSlideXml += '<p:cNvSpPr' + (slideItemObj.options && slideItemObj.options.isTextBox ? ' txBox="1"/>' : '/>');
                strSlideXml += '<p:nvPr>';
                strSlideXml += slideItemObj.type === 'placeholder' ? genXmlPlaceholder(slideItemObj) : genXmlPlaceholder(placeholderObj);
                strSlideXml += '</p:nvPr>';
                strSlideXml += '</p:nvSpPr><p:spPr>';
                strSlideXml += '<a:xfrm' + locationAttr + '>';
                strSlideXml += '<a:off x="' + x + '" y="' + y + '"/>';
                strSlideXml += '<a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm>';
                strSlideXml +=
                    '<a:prstGeom prst="' +
                        shapeType.name +
                        '"><a:avLst>' +
                        (slideItemObj.options.rectRadius
                            ? '<a:gd name="adj" fmla="val ' + Math.round((slideItemObj.options.rectRadius * EMU * 100000) / Math.min(cx, cy)) + '"/>'
                            : '') +
                        '</a:avLst></a:prstGeom>';
                // Option: FILL
                strSlideXml += slideItemObj.options.fill ? genXmlColorSelection(slideItemObj.options.fill) : '<a:noFill/>';
                // shape Type: LINE: line color
                if (slideItemObj.options.line) {
                    strSlideXml += '<a:ln' + (slideItemObj.options.lineSize ? ' w="' + slideItemObj.options.lineSize * ONEPT + '"' : '') + '>';
                    strSlideXml += genXmlColorSelection(slideItemObj.options.line);
                    if (slideItemObj.options.lineDash)
                        strSlideXml += '<a:prstDash val="' + slideItemObj.options.lineDash + '"/>';
                    if (slideItemObj.options.lineHead)
                        strSlideXml += '<a:headEnd type="' + slideItemObj.options.lineHead + '"/>';
                    if (slideItemObj.options.lineTail)
                        strSlideXml += '<a:tailEnd type="' + slideItemObj.options.lineTail + '"/>';
                    strSlideXml += '</a:ln>';
                }
                // EFFECTS > SHADOW: REF: @see http://officeopenxml.com/drwSp-effects.php
                if (slideItemObj.options.shadow) {
                    slideItemObj.options.shadow.type = slideItemObj.options.shadow.type || 'outer';
                    slideItemObj.options.shadow.blur = (slideItemObj.options.shadow.blur || 8) * ONEPT;
                    slideItemObj.options.shadow.offset = (slideItemObj.options.shadow.offset || 4) * ONEPT;
                    slideItemObj.options.shadow.angle = (slideItemObj.options.shadow.angle || 270) * 60000;
                    slideItemObj.options.shadow.color = slideItemObj.options.shadow.color || '000000';
                    slideItemObj.options.shadow.opacity = (slideItemObj.options.shadow.opacity || 0.75) * 100000;
                    strSlideXml += '<a:effectLst>';
                    strSlideXml += '<a:' + slideItemObj.options.shadow.type + 'Shdw sx="100000" sy="100000" kx="0" ky="0" ';
                    strSlideXml += ' algn="bl" rotWithShape="0" blurRad="' + slideItemObj.options.shadow.blur + '" ';
                    strSlideXml += ' dist="' + slideItemObj.options.shadow.offset + '" dir="' + slideItemObj.options.shadow.angle + '">';
                    strSlideXml += '<a:srgbClr val="' + slideItemObj.options.shadow.color + '">';
                    strSlideXml += '<a:alpha val="' + slideItemObj.options.shadow.opacity + '"/></a:srgbClr>';
                    strSlideXml += '</a:outerShdw>';
                    strSlideXml += '</a:effectLst>';
                }
                /* TODO: FUTURE: Text wrapping (copied from MS-PPTX export)
                    // Commented out b/c i'm not even sure this works - current code produces text that wraps in shapes and textboxes, so...
                    if ( slideItemObj.options.textWrap ) {
                        strSlideXml += '<a:extLst>'
                                    + '<a:ext uri="{C572A759-6A51-4108-AA02-DFA0A04FC94B}">'
                                    + '<ma14:wrappingTextBoxFlag xmlns:ma14="http://schemas.microsoft.com/office/mac/drawingml/2011/main" val="1"/>'
                                    + '</a:ext>'
                                    + '</a:extLst>';
                    }
                    */
                // B: Close shape Properties
                strSlideXml += '</p:spPr>';
                // C: Add formatted text (text body "bodyPr")
                strSlideXml += genXmlTextBody(slideItemObj);
                // LAST: Close SHAPE =======================================================
                strSlideXml += '</p:sp>';
                break;
            case SLIDE_OBJECT_TYPES.image:
                var sizing = slideItemObj.options.sizing, rounding = slideItemObj.options.rounding, width = cx, height = cy;
                strSlideXml += '<p:pic>';
                strSlideXml += '  <p:nvPicPr>';
                strSlideXml += '    <p:cNvPr id="' + (idx + 2) + '" name="Object ' + (idx + 1) + '" descr="' + encodeXmlEntities(slideItemObj.image) + '">';
                if (slideItemObj.hyperlink && slideItemObj.hyperlink.url)
                    strSlideXml +=
                        '<a:hlinkClick r:id="rId' +
                            slideItemObj.hyperlink.rId +
                            '" tooltip="' +
                            (slideItemObj.hyperlink.tooltip ? encodeXmlEntities(slideItemObj.hyperlink.tooltip) : '') +
                            '"/>';
                if (slideItemObj.hyperlink && slideItemObj.hyperlink.slide)
                    strSlideXml +=
                        '<a:hlinkClick r:id="rId' +
                            slideItemObj.hyperlink.rId +
                            '" tooltip="' +
                            (slideItemObj.hyperlink.tooltip ? encodeXmlEntities(slideItemObj.hyperlink.tooltip) : '') +
                            '" action="ppaction://hlinksldjump"/>';
                strSlideXml += '    </p:cNvPr>';
                strSlideXml += '    <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>';
                strSlideXml += '    <p:nvPr>' + genXmlPlaceholder(placeholderObj) + '</p:nvPr>';
                strSlideXml += '  </p:nvPicPr>';
                strSlideXml += '<p:blipFill>';
                // NOTE: This works for both cases: either `path` or `data` contains the SVG
                if ((slide['relsMedia'] || []).filter(function (rel) {
                    return rel.rId === slideItemObj.imageRid;
                })[0] &&
                    (slide['relsMedia'] || []).filter(function (rel) {
                        return rel.rId === slideItemObj.imageRid;
                    })[0]['extn'] === 'svg') {
                    strSlideXml += '<a:blip r:embed="rId' + (slideItemObj.imageRid - 1) + '">';
                    strSlideXml += ' <a:extLst>';
                    strSlideXml += '  <a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">';
                    strSlideXml += '   <asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rId' + slideItemObj.imageRid + '"/>';
                    strSlideXml += '  </a:ext>';
                    strSlideXml += ' </a:extLst>';
                    strSlideXml += '</a:blip>';
                }
                else {
                    strSlideXml += '<a:blip r:embed="rId' + slideItemObj.imageRid + '"/>';
                }
                if (sizing && sizing.type) {
                    var boxW = sizing.w ? getSmartParseNumber(sizing.w, 'X', slide.presLayout) : cx, boxH = sizing.h ? getSmartParseNumber(sizing.h, 'Y', slide.presLayout) : cy, boxX = getSmartParseNumber(sizing.x || 0, 'X', slide.presLayout), boxY = getSmartParseNumber(sizing.y || 0, 'Y', slide.presLayout);
                    strSlideXml += imageSizingXml[sizing.type]({ w: width, h: height }, { w: boxW, h: boxH, x: boxX, y: boxY });
                    width = boxW;
                    height = boxH;
                }
                else {
                    strSlideXml += '  <a:stretch><a:fillRect/></a:stretch>';
                }
                strSlideXml += '</p:blipFill>';
                strSlideXml += '<p:spPr>';
                strSlideXml += ' <a:xfrm' + locationAttr + '>';
                strSlideXml += '  <a:off x="' + x + '" y="' + y + '"/>';
                strSlideXml += '  <a:ext cx="' + width + '" cy="' + height + '"/>';
                strSlideXml += ' </a:xfrm>';
                strSlideXml += ' <a:prstGeom prst="' + (rounding ? 'ellipse' : 'rect') + '"><a:avLst/></a:prstGeom>';
                strSlideXml += '</p:spPr>';
                strSlideXml += '</p:pic>';
                break;
            case SLIDE_OBJECT_TYPES.media:
                if (slideItemObj.mtype === 'online') {
                    strSlideXml += '<p:pic>';
                    strSlideXml += ' <p:nvPicPr>';
                    // IMPORTANT: <p:cNvPr id="" value is critical - if not the same number as preview image rId, PowerPoint throws error!
                    strSlideXml += ' <p:cNvPr id="' + (slideItemObj.mediaRid + 2) + '" name="Picture' + (idx + 1) + '"/>';
                    strSlideXml += ' <p:cNvPicPr/>';
                    strSlideXml += ' <p:nvPr>';
                    strSlideXml += '  <a:videoFile r:link="rId' + slideItemObj.mediaRid + '"/>';
                    strSlideXml += ' </p:nvPr>';
                    strSlideXml += ' </p:nvPicPr>';
                    // NOTE: `blip` is diferent than videos; also there's no preview "p:extLst" above but exists in videos
                    strSlideXml += ' <p:blipFill><a:blip r:embed="rId' + (slideItemObj.mediaRid + 1) + '"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'; // NOTE: Preview image is required!
                    strSlideXml += ' <p:spPr>';
                    strSlideXml += '  <a:xfrm' + locationAttr + '>';
                    strSlideXml += '   <a:off x="' + x + '" y="' + y + '"/>';
                    strSlideXml += '   <a:ext cx="' + cx + '" cy="' + cy + '"/>';
                    strSlideXml += '  </a:xfrm>';
                    strSlideXml += '  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>';
                    strSlideXml += ' </p:spPr>';
                    strSlideXml += '</p:pic>';
                }
                else {
                    strSlideXml += '<p:pic>';
                    strSlideXml += ' <p:nvPicPr>';
                    // IMPORTANT: <p:cNvPr id="" value is critical - if not the same number as preiew image rId, PowerPoint throws error!
                    strSlideXml +=
                        ' <p:cNvPr id="' +
                            (slideItemObj.mediaRid + 2) +
                            '" name="' +
                            slideItemObj.media
                                .split('/')
                                .pop()
                                .split('.')
                                .shift() +
                            '"><a:hlinkClick r:id="" action="ppaction://media"/></p:cNvPr>';
                    strSlideXml += ' <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>';
                    strSlideXml += ' <p:nvPr>';
                    strSlideXml += '  <a:videoFile r:link="rId' + slideItemObj.mediaRid + '"/>';
                    strSlideXml += '  <p:extLst>';
                    strSlideXml += '   <p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}">';
                    strSlideXml += '    <p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="rId' + (slideItemObj.mediaRid + 1) + '"/>';
                    strSlideXml += '   </p:ext>';
                    strSlideXml += '  </p:extLst>';
                    strSlideXml += ' </p:nvPr>';
                    strSlideXml += ' </p:nvPicPr>';
                    strSlideXml += ' <p:blipFill><a:blip r:embed="rId' + (slideItemObj.mediaRid + 2) + '"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'; // NOTE: Preview image is required!
                    strSlideXml += ' <p:spPr>';
                    strSlideXml += '  <a:xfrm' + locationAttr + '>';
                    strSlideXml += '   <a:off x="' + x + '" y="' + y + '"/>';
                    strSlideXml += '   <a:ext cx="' + cx + '" cy="' + cy + '"/>';
                    strSlideXml += '  </a:xfrm>';
                    strSlideXml += '  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>';
                    strSlideXml += ' </p:spPr>';
                    strSlideXml += '</p:pic>';
                }
                break;
            case SLIDE_OBJECT_TYPES.chart:
                strSlideXml += '<p:graphicFrame>';
                strSlideXml += ' <p:nvGraphicFramePr>';
                strSlideXml += '   <p:cNvPr id="' + (idx + 2) + '" name="Chart ' + (idx + 1) + '"/>';
                strSlideXml += '   <p:cNvGraphicFramePr/>';
                strSlideXml += '   <p:nvPr>' + genXmlPlaceholder(placeholderObj) + '</p:nvPr>';
                strSlideXml += ' </p:nvGraphicFramePr>';
                strSlideXml += ' <p:xfrm>';
                strSlideXml += '  <a:off x="' + x + '" y="' + y + '"/>';
                strSlideXml += '  <a:ext cx="' + cx + '" cy="' + cy + '"/>';
                strSlideXml += ' </p:xfrm>';
                strSlideXml += ' <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">';
                strSlideXml += '  <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">';
                strSlideXml += '   <c:chart r:id="rId' + slideItemObj.chartRid + '" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>';
                strSlideXml += '  </a:graphicData>';
                strSlideXml += ' </a:graphic>';
                strSlideXml += '</p:graphicFrame>';
                break;
            default:
                break;
        }
    });
    // STEP 5: Add slide numbers (if any) last
    if (slide.slideNumberObj) {
        strSlideXml +=
            '<p:sp>' +
                '  <p:nvSpPr>' +
                '    <p:cNvPr id="25" name="Slide Number Placeholder 24"/>' +
                '    <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' +
                '    <p:nvPr><p:ph type="sldNum" sz="quarter" idx="4294967295"/></p:nvPr>' +
                '  </p:nvSpPr>' +
                '  <p:spPr>' +
                '    <a:xfrm>' +
                '      <a:off x="' +
                getSmartParseNumber(slide.slideNumberObj.x, 'X', slide.presLayout) +
                '" y="' +
                getSmartParseNumber(slide.slideNumberObj.y, 'Y', slide.presLayout) +
                '"/>' +
                '      <a:ext cx="' +
                (slide.slideNumberObj.w ? getSmartParseNumber(slide.slideNumberObj.w, 'X', slide.presLayout) : 800000) +
                '" cy="' +
                (slide.slideNumberObj.h ? getSmartParseNumber(slide.slideNumberObj.h, 'Y', slide.presLayout) : 300000) +
                '"/>' +
                '    </a:xfrm>' +
                '    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
                '    <a:extLst><a:ext uri="{C572A759-6A51-4108-AA02-DFA0A04FC94B}"><ma14:wrappingTextBoxFlag val="0" xmlns:ma14="http://schemas.microsoft.com/office/mac/drawingml/2011/main"/></a:ext></a:extLst>' +
                '  </p:spPr>';
        strSlideXml += '<p:txBody>';
        strSlideXml += '  <a:bodyPr/>';
        strSlideXml += '  <a:lstStyle><a:lvl1pPr>';
        if (slide.slideNumberObj.fontFace || slide.slideNumberObj.fontSize || slide.slideNumberObj.color) {
            strSlideXml += '<a:defRPr sz="' + (slide.slideNumberObj.fontSize ? Math.round(slide.slideNumberObj.fontSize) : '12') + '00">';
            if (slide.slideNumberObj.color)
                strSlideXml += genXmlColorSelection(slide.slideNumberObj.color);
            if (slide.slideNumberObj.fontFace)
                strSlideXml +=
                    '<a:latin typeface="' +
                        slide.slideNumberObj.fontFace +
                        '"/><a:ea typeface="' +
                        slide.slideNumberObj.fontFace +
                        '"/><a:cs typeface="' +
                        slide.slideNumberObj.fontFace +
                        '"/>';
            strSlideXml += '</a:defRPr>';
        }
        strSlideXml += '</a:lvl1pPr></a:lstStyle>';
        strSlideXml += '<a:p><a:fld id="' + SLDNUMFLDID + '" type="slidenum"><a:rPr lang="en-US"/><a:t></a:t></a:fld><a:endParaRPr lang="en-US"/></a:p>';
        strSlideXml += '</p:txBody></p:sp>';
    }
    // STEP 6: Close spTree and finalize slide XML
    strSlideXml += '</p:spTree>';
    strSlideXml += '</p:cSld>';
    // LAST: Return
    return strSlideXml;
}
/**
 * Transforms slide relations to XML string.
 * Extra relations that are not dynamic can be passed using the 2nd arg (e.g. theme relation in master file).
 * These relations use rId series that starts with 1-increased maximum of rIds used for dynamic relations.
 * @param {ISlide | ISlideLayout} slide - slide object whose relations are being transformed
 * @param {{ target: string; type: string }[]} defaultRels - array of default relations
 * @return {string} XML
 */
function slideObjectRelationsToXml(slide, defaultRels) {
    var lastRid = 0; // stores maximum rId used for dynamic relations
    var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
    // STEP 1: Add all rels for this Slide
    slide.rels.forEach(function (rel) {
        lastRid = Math.max(lastRid, rel.rId);
        if (rel.type.toLowerCase().indexOf('hyperlink') > -1) {
            if (rel.data === 'slide') {
                strXml +=
                    '<Relationship Id="rId' +
                        rel.rId +
                        '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"' +
                        ' Target="slide' +
                        rel.Target +
                        '.xml"/>';
            }
            else {
                strXml +=
                    '<Relationship Id="rId' +
                        rel.rId +
                        '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"' +
                        ' Target="' +
                        rel.Target +
                        '" TargetMode="External"/>';
            }
        }
        else if (rel.type.toLowerCase().indexOf('notesSlide') > -1) {
            strXml +=
                '<Relationship Id="rId' + rel.rId + '" Target="' + rel.Target + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"/>';
        }
    });
    (slide.relsChart || []).forEach(function (rel) {
        lastRid = Math.max(lastRid, rel.rId);
        strXml += '<Relationship Id="rId' + rel.rId + '" Target="' + rel.Target + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"/>';
    });
    (slide.relsMedia || []).forEach(function (rel) {
        lastRid = Math.max(lastRid, rel.rId);
        if (rel.type.toLowerCase().indexOf('image') > -1) {
            strXml += '<Relationship Id="rId' + rel.rId + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="' + rel.Target + '"/>';
        }
        else if (rel.type.toLowerCase().indexOf('audio') > -1) {
            // As media has *TWO* rel entries per item, check for first one, if found add second rel with alt style
            if (strXml.indexOf(' Target="' + rel.Target + '"') > -1)
                strXml += '<Relationship Id="rId' + rel.rId + '" Type="http://schemas.microsoft.com/office/2007/relationships/media" Target="' + rel.Target + '"/>';
            else
                strXml +=
                    '<Relationship Id="rId' + rel.rId + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio" Target="' + rel.Target + '"/>';
        }
        else if (rel.type.toLowerCase().indexOf('video') > -1) {
            // As media has *TWO* rel entries per item, check for first one, if found add second rel with alt style
            if (strXml.indexOf(' Target="' + rel.Target + '"') > -1)
                strXml += '<Relationship Id="rId' + rel.rId + '" Type="http://schemas.microsoft.com/office/2007/relationships/media" Target="' + rel.Target + '"/>';
            else
                strXml +=
                    '<Relationship Id="rId' + rel.rId + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video" Target="' + rel.Target + '"/>';
        }
        else if (rel.type.toLowerCase().indexOf('online') > -1) {
            // As media has *TWO* rel entries per item, check for first one, if found add second rel with alt style
            if (strXml.indexOf(' Target="' + rel.Target + '"') > -1)
                strXml += '<Relationship Id="rId' + rel.rId + '" Type="http://schemas.microsoft.com/office/2007/relationships/image" Target="' + rel.Target + '"/>';
            else
                strXml +=
                    '<Relationship Id="rId' +
                        rel.rId +
                        '" Target="' +
                        rel.Target +
                        '" TargetMode="External" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"/>';
        }
    });
    // STEP 2: Add default rels
    defaultRels.forEach(function (rel, idx) {
        strXml += '<Relationship Id="rId' + (lastRid + idx + 1) + '" Type="' + rel.type + '" Target="' + rel.target + '"/>';
    });
    strXml += '</Relationships>';
    return strXml;
}
/**
 * Generate XML Paragraph Properties
 * @param {ISlideObject|IText} textObj - text object
 * @param {boolean} isDefault - array of default relations
 * @return {string} XML
 */
function genXmlParagraphProperties(textObj, isDefault) {
    var strXmlBullet = '', strXmlLnSpc = '', strXmlParaSpc = '';
    var bulletLvl0Margin = 342900;
    var tag = isDefault ? 'a:lvl1pPr' : 'a:pPr';
    var paragraphPropXml = '<' + tag + (textObj.options.rtlMode ? ' rtl="1" ' : '');
    // A: Build paragraphProperties
    {
        // OPTION: align
        if (textObj.options.align) {
            switch (textObj.options.align) {
                case 'left':
                    paragraphPropXml += ' algn="l"';
                    break;
                case 'right':
                    paragraphPropXml += ' algn="r"';
                    break;
                case 'center':
                    paragraphPropXml += ' algn="ctr"';
                    break;
                case 'justify':
                    paragraphPropXml += ' algn="just"';
                    break;
                default:
                    break;
            }
        }
        if (textObj.options.lineSpacing) {
            strXmlLnSpc = '<a:lnSpc><a:spcPts val="' + textObj.options.lineSpacing + '00"/></a:lnSpc>';
        }
        // OPTION: indent
        if (textObj.options.indentLevel && !isNaN(Number(textObj.options.indentLevel)) && textObj.options.indentLevel > 0) {
            paragraphPropXml += ' lvl="' + textObj.options.indentLevel + '"';
        }
        // OPTION: Paragraph Spacing: Before/After
        if (textObj.options.paraSpaceBefore && !isNaN(Number(textObj.options.paraSpaceBefore)) && textObj.options.paraSpaceBefore > 0) {
            strXmlParaSpc += '<a:spcBef><a:spcPts val="' + textObj.options.paraSpaceBefore * 100 + '"/></a:spcBef>';
        }
        if (textObj.options.paraSpaceAfter && !isNaN(Number(textObj.options.paraSpaceAfter)) && textObj.options.paraSpaceAfter > 0) {
            strXmlParaSpc += '<a:spcAft><a:spcPts val="' + textObj.options.paraSpaceAfter * 100 + '"/></a:spcAft>';
        }
        // OPTION: bullet
        // NOTE: OOXML uses the unicode character set for Bullets
        // EX: Unicode Character 'BULLET' (U+2022) ==> '<a:buChar char="&#x2022;"/>'
        if (typeof textObj.options.bullet === 'object') {
            if (textObj.options.bullet.type) {
                if (textObj.options.bullet.type.toString().toLowerCase() === 'number') {
                    paragraphPropXml +=
                        ' marL="' +
                            (textObj.options.indentLevel && textObj.options.indentLevel > 0
                                ? bulletLvl0Margin + bulletLvl0Margin * textObj.options.indentLevel
                                : bulletLvl0Margin) +
                            '" indent="-' +
                            bulletLvl0Margin +
                            '"';
                    strXmlBullet = "<a:buSzPct val=\"100000\"/><a:buFont typeface=\"+mj-lt\"/><a:buAutoNum type=\"" + (textObj.options.bullet.style ||
                        'arabicPeriod') + "\" startAt=\"" + (textObj.options.bullet.startAt || '1') + "\"/>";
                }
            }
            else if (textObj.options.bullet.code) {
                var bulletCode = '&#x' + textObj.options.bullet.code + ';';
                // Check value for hex-ness (s/b 4 char hex)
                if (/^[0-9A-Fa-f]{4}$/.test(textObj.options.bullet.code) === false) {
                    console.warn('Warning: `bullet.code should be a 4-digit hex code (ex: 22AB)`!');
                    bulletCode = BULLET_TYPES['DEFAULT'];
                }
                paragraphPropXml +=
                    ' marL="' +
                        (textObj.options.indentLevel && textObj.options.indentLevel > 0 ? bulletLvl0Margin + bulletLvl0Margin * textObj.options.indentLevel : bulletLvl0Margin) +
                        '" indent="-' +
                        bulletLvl0Margin +
                        '"';
                strXmlBullet = '<a:buSzPct val="100000"/><a:buChar char="' + bulletCode + '"/>';
            }
        }
        else if (textObj.options.bullet === true) {
            paragraphPropXml +=
                ' marL="' +
                    (textObj.options.indentLevel && textObj.options.indentLevel > 0 ? bulletLvl0Margin + bulletLvl0Margin * textObj.options.indentLevel : bulletLvl0Margin) +
                    '" indent="-' +
                    bulletLvl0Margin +
                    '"';
            strXmlBullet = '<a:buSzPct val="100000"/><a:buChar char="' + BULLET_TYPES['DEFAULT'] + '"/>';
        }
        else {
            strXmlBullet = '<a:buNone/>';
        }
        // B: Close Paragraph-Properties
        // IMPORTANT: strXmlLnSpc, strXmlParaSpc, and strXmlBullet require strict ordering - anything out of order is ignored. (PPT-Online, PPT for Mac)
        paragraphPropXml += '>' + strXmlLnSpc + strXmlParaSpc + strXmlBullet;
        if (isDefault) {
            paragraphPropXml += genXmlTextRunProperties(textObj.options, true);
        }
        paragraphPropXml += '</' + tag + '>';
    }
    return paragraphPropXml;
}
/**
 * Generate XML Text Run Properties (`a:rPr`)
 * @param {IObjectOptions|ITextOpts} opts - text options
 * @param {boolean} isDefault - whether these are the default text run properties
 * @return {string} XML
 */
function genXmlTextRunProperties(opts, isDefault) {
    var runProps = '';
    var runPropsTag = isDefault ? 'a:defRPr' : 'a:rPr';
    // BEGIN runProperties (ex: `<a:rPr lang="en-US" sz="1600" b="1" dirty="0">`)
    runProps += '<' + runPropsTag + ' lang="' + (opts.lang ? opts.lang : 'en-US') + '"' + (opts.lang ? ' altLang="en-US"' : '');
    runProps += opts.fontSize ? ' sz="' + Math.round(opts.fontSize) + '00"' : ''; // NOTE: Use round so sizes like '7.5' wont cause corrupt pres.
    runProps += opts.bold ? ' b="1"' : '';
    runProps += opts.italic ? ' i="1"' : '';
    runProps += opts.strike ? ' strike="sngStrike"' : '';
    runProps += opts.underline || opts.hyperlink ? ' u="sng"' : '';
    runProps += opts.subscript ? ' baseline="-40000"' : opts.superscript ? ' baseline="30000"' : '';
    runProps += opts.charSpacing ? ' spc="' + opts.charSpacing * 100 + '" kern="0"' : ''; // IMPORTANT: Also disable kerning; otherwise text won't actually expand
    runProps += ' dirty="0">';
    // Color / Font / Outline are children of <a:rPr>, so add them now before closing the runProperties tag
    if (opts.color || opts.fontFace || opts.outline) {
        if (opts.outline && typeof opts.outline === 'object') {
            runProps += '<a:ln w="' + Math.round((opts.outline.size || 0.75) * ONEPT) + '">' + genXmlColorSelection(opts.outline.color || 'FFFFFF') + '</a:ln>';
        }
        if (opts.color)
            runProps += genXmlColorSelection(opts.color);
        if (opts.fontFace) {
            // NOTE: 'cs' = Complex Script, 'ea' = East Asian (use "-120" instead of "0" - per Issue #174); ea must come first (Issue #174)
            runProps +=
                '<a:latin typeface="' +
                    opts.fontFace +
                    '" pitchFamily="34" charset="0"/>' +
                    '<a:ea typeface="' +
                    opts.fontFace +
                    '" pitchFamily="34" charset="-122"/>' +
                    '<a:cs typeface="' +
                    opts.fontFace +
                    '" pitchFamily="34" charset="-120"/>';
        }
    }
    // Hyperlink support
    if (opts.hyperlink) {
        if (typeof opts.hyperlink !== 'object')
            throw "ERROR: text `hyperlink` option should be an object. Ex: `hyperlink:{url:'https://github.com'}` ";
        else if (!opts.hyperlink.url && !opts.hyperlink.slide)
            throw "ERROR: 'hyperlink requires either `url` or `slide`'";
        else if (opts.hyperlink.url) {
            // TODO: (20170410): FUTURE-FEATURE: color (link is always blue in Keynote and PPT online, so usual text run above isnt honored for links..?)
            //runProps += '<a:uFill>'+ genXmlColorSelection('0000FF') +'</a:uFill>'; // Breaks PPT2010! (Issue#74)
            runProps +=
                '<a:hlinkClick r:id="rId' +
                    opts.hyperlink.rId +
                    '" invalidUrl="" action="" tgtFrame="" tooltip="' +
                    (opts.hyperlink.tooltip ? encodeXmlEntities(opts.hyperlink.tooltip) : '') +
                    '" history="1" highlightClick="0" endSnd="0"/>';
        }
        else if (opts.hyperlink.slide) {
            runProps +=
                '<a:hlinkClick r:id="rId' +
                    opts.hyperlink.rId +
                    '" action="ppaction://hlinksldjump" tooltip="' +
                    (opts.hyperlink.tooltip ? encodeXmlEntities(opts.hyperlink.tooltip) : '') +
                    '"/>';
        }
    }
    // END runProperties
    runProps += '</' + runPropsTag + '>';
    return runProps;
}
/**
 * Builds `<a:r></a:r>` text runs for `<a:p>` paragraphs in textBody
 * @param {IText} textObj - Text object
 * @return {string} XML string
 */
function genXmlTextRun(textObj) {
    var arrLines = [];
    var paraProp = '';
    var xmlTextRun = '';
    // 1: ADD runProperties
    var startInfo = genXmlTextRunProperties(textObj.options, false);
    // 2: LINE-BREAKS/MULTI-LINE: Split text into multi-p:
    arrLines = textObj.text.split(CRLF);
    if (arrLines.length > 1) {
        arrLines.forEach(function (line, idx) {
            xmlTextRun += '<a:r>' + startInfo + '<a:t>' + encodeXmlEntities(line);
            // Stop/Start <p>aragraph as long as there is more lines ahead (otherwise its closed at the end of this function)
            if (idx + 1 < arrLines.length)
                xmlTextRun += (textObj.options.breakLine ? CRLF : '') + '</a:t></a:r>';
        });
    }
    else {
        // Handle cases where addText `text` was an array of objects - if a text object doesnt contain a '\n' it still need alignment!
        // The first pPr-align is done in makeXml - use line countr to ensure we only add subsequently as needed
        xmlTextRun = (textObj.options.align && textObj.options.lineIdx > 0 ? paraProp : '') + '<a:r>' + startInfo + '<a:t>' + encodeXmlEntities(textObj.text);
    }
    // Return paragraph with text run
    return xmlTextRun + '</a:t></a:r>';
}
/**
 * Builds `<a:bodyPr></a:bodyPr>` tag for "genXmlTextBody()"
 * @param {ISlideObject | ITableCell} slideObject - various options
 * @return {string} XML string
 */
function genXmlBodyProperties(slideObject) {
    var bodyProperties = '<a:bodyPr';
    if (slideObject && slideObject.type === SLIDE_OBJECT_TYPES.text && slideObject.options.bodyProp) {
        // PPT-2019 EX: <a:bodyPr wrap="square" lIns="1270" tIns="1270" rIns="1270" bIns="1270" rtlCol="0" anchor="ctr"/>
        // A: Enable or disable textwrapping none or square
        bodyProperties += slideObject.options.bodyProp.wrap ? ' wrap="' + slideObject.options.bodyProp.wrap + '"' : ' wrap="square"';
        // B: Textbox margins [padding]
        if (slideObject.options.bodyProp.lIns || slideObject.options.bodyProp.lIns === 0)
            bodyProperties += ' lIns="' + slideObject.options.bodyProp.lIns + '"';
        if (slideObject.options.bodyProp.tIns || slideObject.options.bodyProp.tIns === 0)
            bodyProperties += ' tIns="' + slideObject.options.bodyProp.tIns + '"';
        if (slideObject.options.bodyProp.rIns || slideObject.options.bodyProp.rIns === 0)
            bodyProperties += ' rIns="' + slideObject.options.bodyProp.rIns + '"';
        if (slideObject.options.bodyProp.bIns || slideObject.options.bodyProp.bIns === 0)
            bodyProperties += ' bIns="' + slideObject.options.bodyProp.bIns + '"';
        // C: Add rtl after margins
        bodyProperties += ' rtlCol="0"';
        // D: Add anchorPoints
        if (slideObject.options.bodyProp.anchor)
            bodyProperties += ' anchor="' + slideObject.options.bodyProp.anchor + '"'; // VALS: [t,ctr,b]
        if (slideObject.options.bodyProp.vert)
            bodyProperties += ' vert="' + slideObject.options.bodyProp.vert + '"'; // VALS: [eaVert,horz,mongolianVert,vert,vert270,wordArtVert,wordArtVertRtl]
        // E: Close <a:bodyPr element
        bodyProperties += '>';
        // F: NEW: Add autofit type tags
        if (slideObject.options.shrinkText)
            bodyProperties += '<a:normAutofit fontScale="85000" lnSpcReduction="20000"/>'; // MS-PPT > Format shape > Text Options: "Shrink text on overflow"
        // MS-PPT > Format shape > Text Options: "Resize shape to fit text" [spAutoFit]
        // NOTE: Use of '<a:noAutofit/>' in lieu of '' below causes issues in PPT-2013
        bodyProperties += slideObject.options.bodyProp.autoFit !== false ? '<a:spAutoFit/>' : '';
        // LAST: Close bodyProp
        bodyProperties += '</a:bodyPr>';
    }
    else {
        // DEFAULT:
        bodyProperties += ' wrap="square" rtlCol="0">';
        bodyProperties += '</a:bodyPr>';
    }
    // LAST: Return Close bodyProp
    return slideObject.type === SLIDE_OBJECT_TYPES.tablecell ? '<a:bodyPr/>' : bodyProperties;
}
/**
 * Generate the XML for text and its options (bold, bullet, etc) including text runs (word-level formatting)
 * @note PPT text lines [lines followed by line-breaks] are created using <p>-aragraph's
 * @note Bullets are a paragprah-level formatting device
 * @param {ISlideObject|ITableCell} slideObj - slideObj -OR- table `cell` object
 * @returns XML containing the param object's text and formatting
 */
function genXmlTextBody(slideObj) {
    var opts = slideObj.options || {};
    // FIRST: Shapes without text, etc. may be sent here during build, but have no text to render so return an empty string
    if (opts && slideObj.type !== SLIDE_OBJECT_TYPES.tablecell && (typeof slideObj.text === 'undefined' || slideObj.text === null))
        return '';
    // Vars
    var arrTextObjects = [];
    var tagStart = slideObj.type === SLIDE_OBJECT_TYPES.tablecell ? '<a:txBody>' : '<p:txBody>';
    var tagClose = slideObj.type === SLIDE_OBJECT_TYPES.tablecell ? '</a:txBody>' : '</p:txBody>';
    var strSlideXml = tagStart;
    // STEP 1: Modify slideObj to be consistent array of `{ text:'', options:{} }`
    /* CASES:
        addText( 'string' )
        addText( 'line1\n line2' )
        addText( ['barry','allen'] )
        addText( [{text'word1'}, {text:'word2'}] )
        addText( [{text'line1\n line2'}, {text:'end word'}] )
    */
    // A: Transform string/number into complex object
    if (typeof slideObj.text === 'string' || typeof slideObj.text === 'number') {
        slideObj.text = [{ text: slideObj.text.toString(), options: opts || {} }];
    }
    // STEP 2: Grab options, format line-breaks, etc.
    if (Array.isArray(slideObj.text)) {
        slideObj.text.forEach(function (obj, idx) {
            // A: Set options
            obj.options = obj.options || opts || {};
            if (idx === 0 && obj.options && !obj.options.bullet && opts.bullet)
                obj.options.bullet = opts.bullet;
            // B: Cast to text-object and fix line-breaks (if needed)
            if (typeof obj.text === 'string' || typeof obj.text === 'number') {
                // 1: Convert "\n" or any variation into CRLF
                obj.text = obj.text.toString().replace(/\r*\n/g, CRLF);
                // 2: Handle strings that contain "\n"
                if (obj.text.indexOf(CRLF) > -1) {
                    // Remove trailing linebreak (if any) so the "if" below doesnt create a double CRLF+CRLF line ending!
                    obj.text = obj.text.replace(/\r\n$/g, '');
                    // Plain strings like "hello \n world" or "first line\n" need to have lineBreaks set to become 2 separate lines as intended
                    obj.options.breakLine = true;
                }
                // 3: Add CRLF line ending if `breakLine`
                if (obj.options.breakLine && !obj.options.bullet && !obj.options.align && idx + 1 < slideObj.text.length)
                    obj.text += CRLF;
            }
            // C: If text string has line-breaks, then create a separate text-object for each (much easier than dealing with split inside a loop below)
            if (obj.options.breakLine || obj.text.indexOf(CRLF) > -1) {
                obj.text.split(CRLF).forEach(function (line, lineIdx) {
                    // Add line-breaks if not bullets/aligned (we add CRLF for those below in STEP 3)
                    // NOTE: Use "idx>0" so lines wont start with linebreak (eg:empty first line)
                    arrTextObjects.push({
                        text: (lineIdx > 0 && obj.options.breakLine && !obj.options.bullet && !obj.options.align ? CRLF : '') + line,
                        options: obj.options,
                    });
                });
            }
            else {
                // NOTE: The replace used here is for non-textObjects (plain strings) eg:'hello\nworld'
                arrTextObjects.push(obj);
            }
        });
    }
    // STEP 3: Add bodyProperties
    {
        // A: 'bodyPr'
        strSlideXml += genXmlBodyProperties(slideObj);
        // B: 'lstStyle'
        // NOTE: shape type 'LINE' has different text align needs (a lstStyle.lvl1pPr between bodyPr and p)
        // FIXME: LINE horiz-align doesnt work (text is always to the left inside line) (FYI: the PPT code diff is substantial!)
        if (opts.h === 0 && opts.line && opts.align) {
            strSlideXml += '<a:lstStyle><a:lvl1pPr algn="l"/></a:lstStyle>';
        }
        else if (slideObj.type === 'placeholder') {
            strSlideXml += '<a:lstStyle>';
            strSlideXml += genXmlParagraphProperties(slideObj, true);
            strSlideXml += '</a:lstStyle>';
        }
        else {
            strSlideXml += '<a:lstStyle/>';
        }
    }
    // STEP 4: Loop over each text object and create paragraph props, text run, etc.
    arrTextObjects.forEach(function (textObj, idx) {
        // Clear/Increment loop vars
        var paragraphPropXml = '<a:pPr ' + (textObj.options.rtlMode ? ' rtl="1" ' : '');
        textObj.options.lineIdx = idx;
        // A: Inherit pPr-type options from parent shape's `options`
        textObj.options.align = textObj.options.align || opts.align;
        textObj.options.lineSpacing = textObj.options.lineSpacing || opts.lineSpacing;
        textObj.options.indentLevel = textObj.options.indentLevel || opts.indentLevel;
        textObj.options.paraSpaceBefore = textObj.options.paraSpaceBefore || opts.paraSpaceBefore;
        textObj.options.paraSpaceAfter = textObj.options.paraSpaceAfter || opts.paraSpaceAfter;
        textObj.options.lineIdx = idx;
        paragraphPropXml = genXmlParagraphProperties(textObj, false);
        // B: Start paragraph if this is the first text obj, or if current textObj is about to be bulleted or aligned
        if (idx === 0) {
            // Add paragraphProperties right after <p> before textrun(s) begin
            strSlideXml += '<a:p>' + paragraphPropXml;
        }
        else if (idx > 0 && (typeof textObj.options.bullet !== 'undefined' || typeof textObj.options.align !== 'undefined')) {
            strSlideXml += '</a:p><a:p>' + paragraphPropXml;
        }
        // C: Inherit any main options (color, fontSize, etc.)
        // We only pass the text.options to genXmlTextRun (not the Slide.options),
        // so the run building function cant just fallback to Slide.color, therefore, we need to do that here before passing options below.
        Object.entries(opts).forEach(function (_a) {
            var key = _a[0], val = _a[1];
            // NOTE: This loop will pick up unecessary keys (`x`, etc.), but it doesnt hurt anything
            if (key !== 'bullet' && !textObj.options[key])
                textObj.options[key] = val;
        });
        // D: Add formatted textrun
        strSlideXml += genXmlTextRun(textObj);
    });
    // STEP 5: Append 'endParaRPr' (when needed) and close current open paragraph
    // NOTE: (ISSUE#20, ISSUE#193): Add 'endParaRPr' with font/size props or PPT default (Arial/18pt en-us) is used making row "too tall"/not honoring options
    if (slideObj.type === SLIDE_OBJECT_TYPES.tablecell && (opts.fontSize || opts.fontFace)) {
        if (opts.fontFace) {
            strSlideXml +=
                '<a:endParaRPr lang="' + (opts.lang ? opts.lang : 'en-US') + '"' + (opts.fontSize ? ' sz="' + Math.round(opts.fontSize) + '00"' : '') + ' dirty="0">';
            strSlideXml += '<a:latin typeface="' + opts.fontFace + '" charset="0"/>';
            strSlideXml += '<a:ea typeface="' + opts.fontFace + '" charset="0"/>';
            strSlideXml += '<a:cs typeface="' + opts.fontFace + '" charset="0"/>';
            strSlideXml += '</a:endParaRPr>';
        }
        else {
            strSlideXml +=
                '<a:endParaRPr lang="' + (opts.lang ? opts.lang : 'en-US') + '"' + (opts.fontSize ? ' sz="' + Math.round(opts.fontSize) + '00"' : '') + ' dirty="0"/>';
        }
    }
    else {
        strSlideXml += '<a:endParaRPr lang="' + (opts.lang || 'en-US') + '" dirty="0"/>'; // NOTE: Added 20180101 to address PPT-2007 issues
    }
    strSlideXml += '</a:p>';
    // STEP 6: Close the textBody
    strSlideXml += tagClose;
    // LAST: Return XML
    return strSlideXml;
}
/**
 * Generate an XML Placeholder
 * @param {ISlideObject} placeholderObj
 * @returns XML
 */
function genXmlPlaceholder(placeholderObj) {
    if (!placeholderObj)
        return '';
    var placeholderIdx = placeholderObj.options && placeholderObj.options.placeholderIdx ? placeholderObj.options.placeholderIdx : '';
    var placeholderType = placeholderObj.options && placeholderObj.options.placeholderType ? placeholderObj.options.placeholderType : '';
    return "<p:ph\n\t\t" + (placeholderIdx ? ' idx="' + placeholderIdx + '"' : '') + "\n\t\t" + (placeholderType && PLACEHOLDER_TYPES[placeholderType] ? ' type="' + PLACEHOLDER_TYPES[placeholderType] + '"' : '') + "\n\t\t" + (placeholderObj.text && placeholderObj.text.length > 0 ? ' hasCustomPrompt="1"' : '') + "\n\t\t/>";
}
// XML-GEN: First 6 functions create the base /ppt files
/**
 * Generate XML ContentType
 * @param {ISlide[]} slides - slides
 * @param {ISlideLayout[]} slideLayouts - slide layouts
 * @param {ISlide} masterSlide - master slide
 * @returns XML
 */
function makeXmlContTypes(slides, slideLayouts, masterSlide) {
    var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF;
    strXml += '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">';
    strXml += '<Default Extension="xml" ContentType="application/xml"/>';
    strXml += '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
    strXml += '<Default Extension="jpeg" ContentType="image/jpeg"/>';
    strXml += '<Default Extension="jpg" ContentType="image/jpg"/>';
    // STEP 1: Add standard/any media types used in Presenation
    strXml += '<Default Extension="png" ContentType="image/png"/>';
    strXml += '<Default Extension="gif" ContentType="image/gif"/>';
    strXml += '<Default Extension="m4v" ContentType="video/mp4"/>'; // NOTE: Hard-Code this extension as it wont be created in loop below (as extn !== type)
    strXml += '<Default Extension="mp4" ContentType="video/mp4"/>'; // NOTE: Hard-Code this extension as it wont be created in loop below (as extn !== type)
    slides.forEach(function (slide) {
        (slide.relsMedia || []).forEach(function (rel) {
            if (rel.type !== 'image' && rel.type !== 'online' && rel.type !== 'chart' && rel.extn !== 'm4v' && strXml.indexOf(rel.type) === -1) {
                strXml += '<Default Extension="' + rel.extn + '" ContentType="' + rel.type + '"/>';
            }
        });
    });
    strXml += '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>';
    strXml += '<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>';
    // STEP 2: Add presentation and slide master(s)/slide(s)
    strXml += '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>';
    strXml += '<Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"/>';
    slides.forEach(function (slide, idx) {
        strXml +=
            '<Override PartName="/ppt/slideMasters/slideMaster' +
                (idx + 1) +
                '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>';
        strXml += '<Override PartName="/ppt/slides/slide' + (idx + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>';
        // Add charts if any
        slide.relsChart.forEach(function (rel) {
            strXml += ' <Override PartName="' + rel.Target + '" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>';
        });
    });
    // STEP 3: Core PPT
    strXml += '<Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>';
    strXml += '<Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>';
    strXml += '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>';
    strXml += '<Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>';
    // STEP 4: Add Slide Layouts
    slideLayouts.forEach(function (layout, idx) {
        strXml +=
            '<Override PartName="/ppt/slideLayouts/slideLayout' +
                (idx + 1) +
                '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>';
        (layout.relsChart || []).forEach(function (rel) {
            strXml += ' <Override PartName="' + rel.Target + '" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>';
        });
    });
    // STEP 5: Add notes slide(s)
    slides.forEach(function (_slide, idx) {
        strXml +=
            ' <Override PartName="/ppt/notesSlides/notesSlide' +
                (idx + 1) +
                '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>';
    });
    // STEP 6: Add rels
    masterSlide.relsChart.forEach(function (rel) {
        strXml += ' <Override PartName="' + rel.Target + '" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>';
    });
    masterSlide.relsMedia.forEach(function (rel) {
        if (rel.type !== 'image' && rel.type !== 'online' && rel.type !== 'chart' && rel.extn !== 'm4v' && strXml.indexOf(rel.type) === -1)
            strXml += ' <Default Extension="' + rel.extn + '" ContentType="' + rel.type + '"/>';
    });
    // LAST: Finish XML (Resume core)
    strXml += ' <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>';
    strXml += ' <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>';
    strXml += '</Types>';
    return strXml;
}
/**
 * Creates `_rels/.rels`
 * @returns XML
 */
function makeXmlRootRels() {
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + CRLF + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n\t\t<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>\n\t\t<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>\n\t\t<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"ppt/presentation.xml\"/>\n\t\t</Relationships>";
}
/**
 * Creates `docProps/app.xml`
 * @param {ISlide[]} slides - Presenation Slides
 * @param {string} company - "Company" metadata
 * @returns XML
 */
function makeXmlApp(slides, company) {
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + CRLF + "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">\n\t<TotalTime>0</TotalTime>\n\t<Words>0</Words>\n\t<Application>Microsoft Office PowerPoint</Application>\n\t<PresentationFormat>On-screen Show (16:9)</PresentationFormat>\n\t<Paragraphs>0</Paragraphs>\n\t<Slides>" + slides.length + "</Slides>\n\t<Notes>" + slides.length + "</Notes>\n\t<HiddenSlides>0</HiddenSlides>\n\t<MMClips>0</MMClips>\n\t<ScaleCrop>false</ScaleCrop>\n\t<HeadingPairs>\n\t\t<vt:vector size=\"6\" baseType=\"variant\">\n\t\t\t<vt:variant><vt:lpstr>Fonts Used</vt:lpstr></vt:variant>\n\t\t\t<vt:variant><vt:i4>2</vt:i4></vt:variant>\n\t\t\t<vt:variant><vt:lpstr>Theme</vt:lpstr></vt:variant>\n\t\t\t<vt:variant><vt:i4>1</vt:i4></vt:variant>\n\t\t\t<vt:variant><vt:lpstr>Slide Titles</vt:lpstr></vt:variant>\n\t\t\t<vt:variant><vt:i4>" + slides.length + "</vt:i4></vt:variant>\n\t\t</vt:vector>\n\t</HeadingPairs>\n\t<TitlesOfParts>\n\t\t<vt:vector size=\"" + (slides.length + 1 + 2) + "\" baseType=\"lpstr\">\n\t\t\t<vt:lpstr>Arial</vt:lpstr>\n\t\t\t<vt:lpstr>Calibri</vt:lpstr>\n\t\t\t<vt:lpstr>Office Theme</vt:lpstr>\n\t\t\t" + slides
        .map(function (_slideObj, idx) {
        return '<vt:lpstr>Slide ' + (idx + 1) + '</vt:lpstr>\n';
    })
        .join('') + "\n\t\t</vt:vector>\n\t</TitlesOfParts>\n\t<Company>" + company + "</Company>\n\t<LinksUpToDate>false</LinksUpToDate>\n\t<SharedDoc>false</SharedDoc>\n\t<HyperlinksChanged>false</HyperlinksChanged>\n\t<AppVersion>16.0000</AppVersion>\n\t</Properties>";
}
/**
 * Creates `docProps/core.xml`
 * @param {string} title - metadata data
 * @param {string} company - metadata data
 * @param {string} author - metadata value
 * @param {string} revision - metadata value
 * @returns XML
 */
function makeXmlCore(title, subject, author, revision) {
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n\t<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">\n\t\t<dc:title>" + encodeXmlEntities(title) + "</dc:title>\n\t\t<dc:subject>" + encodeXmlEntities(subject) + "</dc:subject>\n\t\t<dc:creator>" + encodeXmlEntities(author) + "</dc:creator>\n\t\t<cp:lastModifiedBy>" + encodeXmlEntities(author) + "</cp:lastModifiedBy>\n\t\t<cp:revision>" + revision + "</cp:revision>\n\t\t<dcterms:created xsi:type=\"dcterms:W3CDTF\">" + new Date().toISOString().replace(/\.\d\d\dZ/, 'Z') + "</dcterms:created>\n\t\t<dcterms:modified xsi:type=\"dcterms:W3CDTF\">" + new Date().toISOString().replace(/\.\d\d\dZ/, 'Z') + "</dcterms:modified>\n\t</cp:coreProperties>";
}
/**
 * Creates `ppt/_rels/presentation.xml.rels`
 * @param {ISlide[]} slides - Presenation Slides
 * @returns XML
 */
function makeXmlPresentationRels(slides) {
    var intRelNum = 1;
    var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF;
    strXml += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
    strXml += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>';
    for (var idx = 1; idx <= slides.length; idx++) {
        strXml +=
            '<Relationship Id="rId' + ++intRelNum + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide' + idx + '.xml"/>';
    }
    intRelNum++;
    strXml +=
        '<Relationship Id="rId' +
            intRelNum +
            '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="notesMasters/notesMaster1.xml"/>' +
            '<Relationship Id="rId' +
            (intRelNum + 1) +
            '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/>' +
            '<Relationship Id="rId' +
            (intRelNum + 2) +
            '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/>' +
            '<Relationship Id="rId' +
            (intRelNum + 3) +
            '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>' +
            '<Relationship Id="rId' +
            (intRelNum + 4) +
            '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" Target="tableStyles.xml"/>' +
            '</Relationships>';
    return strXml;
}
// XML-GEN: Functions that run 1-N times (once for each Slide)
/**
 * Generates XML for the slide file (`ppt/slides/slide1.xml`)
 * @param {ISlide} slide - the slide object to transform into XML
 * @return {string} XML
 */
function makeXmlSlide(slide) {
    return ("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + CRLF +
        "<p:sld xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
        "xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"" +
        ((slide && slide.hidden ? ' show="0"' : '') + ">") +
        ("" + slideObjectToXml(slide)) +
        "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>");
}
/**
 * Get text content of Notes from Slide
 * @param {ISlide} slide - the slide object to transform into XML
 * @return {string} notes text
 */
function getNotesFromSlide(slide) {
    var notesText = '';
    slide.data.forEach(function (data) {
        if (data.type === 'notes')
            notesText += data.text;
    });
    return notesText.replace(/\r*\n/g, CRLF);
}
/**
 * Generate XML for Notes Master (notesMaster1.xml)
 * @returns {string} XML
 */
function makeXmlNotesMaster() {
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + CRLF + "<p:notesMaster xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:cSld><p:bg><p:bgRef idx=\"1001\"><a:schemeClr val=\"bg1\"/></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id=\"2\" name=\"Header Placeholder 1\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"hdr\" sz=\"quarter\"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"2971800\" cy=\"458788\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert=\"horz\" lIns=\"91440\" tIns=\"45720\" rIns=\"91440\" bIns=\"45720\" rtlCol=\"0\"/><a:lstStyle><a:lvl1pPr algn=\"l\"><a:defRPr sz=\"1200\"/></a:lvl1pPr></a:lstStyle><a:p><a:endParaRPr lang=\"en-US\"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id=\"3\" name=\"Date Placeholder 2\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"dt\" idx=\"1\"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"3884613\" y=\"0\"/><a:ext cx=\"2971800\" cy=\"458788\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert=\"horz\" lIns=\"91440\" tIns=\"45720\" rIns=\"91440\" bIns=\"45720\" rtlCol=\"0\"/><a:lstStyle><a:lvl1pPr algn=\"r\"><a:defRPr sz=\"1200\"/></a:lvl1pPr></a:lstStyle><a:p><a:fld id=\"{5282F153-3F37-0F45-9E97-73ACFA13230C}\" type=\"datetimeFigureOut\"><a:rPr lang=\"en-US\"/><a:t>7/23/19</a:t></a:fld><a:endParaRPr lang=\"en-US\"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id=\"4\" name=\"Slide Image Placeholder 3\"/><p:cNvSpPr><a:spLocks noGrp=\"1\" noRot=\"1\" noChangeAspect=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"sldImg\" idx=\"2\"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"685800\" y=\"1143000\"/><a:ext cx=\"5486400\" cy=\"3086100\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/><a:ln w=\"12700\"><a:solidFill><a:prstClr val=\"black\"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr vert=\"horz\" lIns=\"91440\" tIns=\"45720\" rIns=\"91440\" bIns=\"45720\" rtlCol=\"0\" anchor=\"ctr\"/><a:lstStyle/><a:p><a:endParaRPr lang=\"en-US\"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id=\"5\" name=\"Notes Placeholder 4\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"body\" sz=\"quarter\" idx=\"3\"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"685800\" y=\"4400550\"/><a:ext cx=\"5486400\" cy=\"3600450\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert=\"horz\" lIns=\"91440\" tIns=\"45720\" rIns=\"91440\" bIns=\"45720\" rtlCol=\"0\"/><a:lstStyle/><a:p><a:pPr lvl=\"0\"/><a:r><a:rPr lang=\"en-US\"/><a:t>Click to edit Master text styles</a:t></a:r></a:p><a:p><a:pPr lvl=\"1\"/><a:r><a:rPr lang=\"en-US\"/><a:t>Second level</a:t></a:r></a:p><a:p><a:pPr lvl=\"2\"/><a:r><a:rPr lang=\"en-US\"/><a:t>Third level</a:t></a:r></a:p><a:p><a:pPr lvl=\"3\"/><a:r><a:rPr lang=\"en-US\"/><a:t>Fourth level</a:t></a:r></a:p><a:p><a:pPr lvl=\"4\"/><a:r><a:rPr lang=\"en-US\"/><a:t>Fifth level</a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id=\"6\" name=\"Footer Placeholder 5\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"ftr\" sz=\"quarter\" idx=\"4\"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"0\" y=\"8685213\"/><a:ext cx=\"2971800\" cy=\"458787\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert=\"horz\" lIns=\"91440\" tIns=\"45720\" rIns=\"91440\" bIns=\"45720\" rtlCol=\"0\" anchor=\"b\"/><a:lstStyle><a:lvl1pPr algn=\"l\"><a:defRPr sz=\"1200\"/></a:lvl1pPr></a:lstStyle><a:p><a:endParaRPr lang=\"en-US\"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id=\"7\" name=\"Slide Number Placeholder 6\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"sldNum\" sz=\"quarter\" idx=\"5\"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"3884613\" y=\"8685213\"/><a:ext cx=\"2971800\" cy=\"458787\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert=\"horz\" lIns=\"91440\" tIns=\"45720\" rIns=\"91440\" bIns=\"45720\" rtlCol=\"0\" anchor=\"b\"/><a:lstStyle><a:lvl1pPr algn=\"r\"><a:defRPr sz=\"1200\"/></a:lvl1pPr></a:lstStyle><a:p><a:fld id=\"{CE5E9CC1-C706-0F49-92D6-E571CC5EEA8F}\" type=\"slidenum\"><a:rPr lang=\"en-US\"/><a:t>\u2039#\u203A</a:t></a:fld><a:endParaRPr lang=\"en-US\"/></a:p></p:txBody></p:sp></p:spTree><p:extLst><p:ext uri=\"{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}\"><p14:creationId xmlns:p14=\"http://schemas.microsoft.com/office/powerpoint/2010/main\" val=\"1024086991\"/></p:ext></p:extLst></p:cSld><p:clrMap bg1=\"lt1\" tx1=\"dk1\" bg2=\"lt2\" tx2=\"dk2\" accent1=\"accent1\" accent2=\"accent2\" accent3=\"accent3\" accent4=\"accent4\" accent5=\"accent5\" accent6=\"accent6\" hlink=\"hlink\" folHlink=\"folHlink\"/><p:notesStyle><a:lvl1pPr marL=\"0\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\"><a:defRPr sz=\"1200\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill><a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/></a:defRPr></a:lvl1pPr><a:lvl2pPr marL=\"457200\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\"><a:defRPr sz=\"1200\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill><a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/></a:defRPr></a:lvl2pPr><a:lvl3pPr marL=\"914400\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\"><a:defRPr sz=\"1200\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill><a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/></a:defRPr></a:lvl3pPr><a:lvl4pPr marL=\"1371600\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\"><a:defRPr sz=\"1200\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill><a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/></a:defRPr></a:lvl4pPr><a:lvl5pPr marL=\"1828800\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\"><a:defRPr sz=\"1200\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill><a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/></a:defRPr></a:lvl5pPr><a:lvl6pPr marL=\"2286000\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\"><a:defRPr sz=\"1200\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill><a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/></a:defRPr></a:lvl6pPr><a:lvl7pPr marL=\"2743200\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\"><a:defRPr sz=\"1200\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill><a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/></a:defRPr></a:lvl7pPr><a:lvl8pPr marL=\"3200400\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\"><a:defRPr sz=\"1200\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill><a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/></a:defRPr></a:lvl8pPr><a:lvl9pPr marL=\"3657600\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\"><a:defRPr sz=\"1200\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill><a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/></a:defRPr></a:lvl9pPr></p:notesStyle></p:notesMaster>";
}
/**
 * Creates Notes Slide (`ppt/notesSlides/notesSlide1.xml`)
 * @param {ISlide} slide - the slide object to transform into XML
 * @return {string} XML
 */
function makeXmlNotesSlide(slide) {
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        CRLF +
        '<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
        '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/>' +
        '<p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/>' +
        '<a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/>' +
        '</a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/>' +
        '<p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr>' +
        '<p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/>' +
        '</p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/>' +
        '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>' +
        '<p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>' +
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r>' +
        '<a:rPr lang="en-US" dirty="0"/><a:t>' +
        encodeXmlEntities(getNotesFromSlide(slide)) +
        '</a:t></a:r><a:endParaRPr lang="en-US" dirty="0"/></a:p></p:txBody>' +
        '</p:sp><p:sp><p:nvSpPr><p:cNvPr id="4" name="Slide Number Placeholder 3"/>' +
        '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>' +
        '<p:ph type="sldNum" sz="quarter" idx="10"/></p:nvPr></p:nvSpPr>' +
        '<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p>' +
        '<a:fld id="' +
        SLDNUMFLDID +
        '" type="slidenum">' +
        '<a:rPr lang="en-US"/><a:t>' +
        slide.number +
        '</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>' +
        '</p:spTree><p:extLst><p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}">' +
        '<p14:creationId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="1024086991"/>' +
        '</p:ext></p:extLst></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>');
}
/**
 * Generates the XML layout resource from a layout object
 * @param {ISlideLayout} layout - slide layout (master)
 * @return {string} XML
 */
function makeXmlLayout(layout) {
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n\t\t<p:sldLayout xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" preserve=\"1\">\n\t\t" + slideObjectToXml(layout) + "\n\t\t<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>";
}
/**
 * Creates Slide Master 1 (`ppt/slideMasters/slideMaster1.xml`)
 * @param {ISlide} slide - slide object that represents master slide layout
 * @param {ISlideLayout[]} layouts - slide layouts
 * @return {string} XML
 */
function makeXmlMaster(slide, layouts) {
    // NOTE: Pass layouts as static rels because they are not referenced any time
    var layoutDefs = layouts.map(function (_layoutDef, idx) {
        return '<p:sldLayoutId id="' + (LAYOUT_IDX_SERIES_BASE + idx) + '" r:id="rId' + (slide.rels.length + idx + 1) + '"/>';
    });
    var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF;
    strXml +=
        '<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">';
    strXml += slideObjectToXml(slide);
    strXml +=
        '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>';
    strXml += '<p:sldLayoutIdLst>' + layoutDefs.join('') + '</p:sldLayoutIdLst>';
    strXml += '<p:hf sldNum="0" hdr="0" ftr="0" dt="0"/>';
    strXml +=
        '<p:txStyles>' +
            ' <p:titleStyle>' +
            '  <a:lvl1pPr algn="ctr" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="0"/></a:spcBef><a:buNone/><a:defRPr sz="4400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mj-cs"/></a:defRPr></a:lvl1pPr>' +
            ' </p:titleStyle>' +
            ' <p:bodyStyle>' +
            '  <a:lvl1pPr marL="342900" indent="-342900" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="3200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr>' +
            '  <a:lvl2pPr marL="742950" indent="-285750" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="–"/><a:defRPr sz="2800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl2pPr>' +
            '  <a:lvl3pPr marL="1143000" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl3pPr>' +
            '  <a:lvl4pPr marL="1600200" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="–"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl4pPr>' +
            '  <a:lvl5pPr marL="2057400" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="»"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl5pPr>' +
            '  <a:lvl6pPr marL="2514600" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl6pPr>' +
            '  <a:lvl7pPr marL="2971800" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl7pPr>' +
            '  <a:lvl8pPr marL="3429000" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl8pPr>' +
            '  <a:lvl9pPr marL="3886200" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl9pPr>' +
            ' </p:bodyStyle>' +
            ' <p:otherStyle>' +
            '  <a:defPPr><a:defRPr lang="en-US"/></a:defPPr>' +
            '  <a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr>' +
            '  <a:lvl2pPr marL="457200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl2pPr>' +
            '  <a:lvl3pPr marL="914400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl3pPr>' +
            '  <a:lvl4pPr marL="1371600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl4pPr>' +
            '  <a:lvl5pPr marL="1828800" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl5pPr>' +
            '  <a:lvl6pPr marL="2286000" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl6pPr>' +
            '  <a:lvl7pPr marL="2743200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl7pPr>' +
            '  <a:lvl8pPr marL="3200400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl8pPr>' +
            '  <a:lvl9pPr marL="3657600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl9pPr>' +
            ' </p:otherStyle>' +
            '</p:txStyles>';
    strXml += '</p:sldMaster>';
    return strXml;
}
/**
 * Generates XML string for a slide layout relation file
 * @param {number} layoutNumber - 1-indexed number of a layout that relations are generated for
 * @param {ISlideLayout[]} slideLayouts - Slide Layouts
 * @return {string} XML
 */
function makeXmlSlideLayoutRel(layoutNumber, slideLayouts) {
    return slideObjectRelationsToXml(slideLayouts[layoutNumber - 1], [
        {
            target: '../slideMasters/slideMaster1.xml',
            type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
        },
    ]);
}
/**
 * Creates `ppt/_rels/slide*.xml.rels`
 * @param {ISlide[]} slides
 * @param {ISlideLayout[]} slideLayouts - Slide Layout(s)
 * @param {number} `slideNumber` 1-indexed number of a layout that relations are generated for
 * @return {string} XML
 */
function makeXmlSlideRel(slides, slideLayouts, slideNumber) {
    return slideObjectRelationsToXml(slides[slideNumber - 1], [
        {
            target: '../slideLayouts/slideLayout' + getLayoutIdxForSlide(slides, slideLayouts, slideNumber) + '.xml',
            type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
        },
        {
            target: '../notesSlides/notesSlide' + slideNumber + '.xml',
            type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide',
        },
    ]);
}
/**
 * Generates XML string for a slide relation file.
 * @param {number} slideNumber - 1-indexed number of a layout that relations are generated for
 * @return {string} XML
 */
function makeXmlNotesSlideRel(slideNumber) {
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n\t\t<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n\t\t\t<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster\" Target=\"../notesMasters/notesMaster1.xml\"/>\n\t\t\t<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" Target=\"../slides/slide" + slideNumber + ".xml\"/>\n\t\t</Relationships>";
}
/**
 * Creates `ppt/slideMasters/_rels/slideMaster1.xml.rels`
 * @param {ISlide} masterSlide - Slide object
 * @param {ISlideLayout[]} slideLayouts - Slide Layouts
 * @return {string} XML
 */
function makeXmlMasterRel(masterSlide, slideLayouts) {
    var defaultRels = slideLayouts.map(function (_layoutDef, idx) {
        return { target: "../slideLayouts/slideLayout" + (idx + 1) + ".xml", type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout' };
    });
    defaultRels.push({ target: '../theme/theme1.xml', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme' });
    return slideObjectRelationsToXml(masterSlide, defaultRels);
}
/**
 * Creates `ppt/notesMasters/_rels/notesMaster1.xml.rels`
 * @return {string} XML
 */
function makeXmlNotesMasterRel() {
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + CRLF + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n\t\t<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"../theme/theme1.xml\"/>\n\t\t</Relationships>";
}
/**
 * For the passed slide number, resolves name of a layout that is used for.
 * @param {ISlide[]} slides - srray of slides
 * @param {ISlideLayout[]} slideLayouts - array of slideLayouts
 * @param {number} slideNumber
 * @return {number} slide number
 */
function getLayoutIdxForSlide(slides, slideLayouts, slideNumber) {
    for (var i = 0; i < slideLayouts.length; i++) {
        if (slideLayouts[i].name === slides[slideNumber - 1].slideLayout.name) {
            return i + 1;
        }
    }
    // IMPORTANT: Return 1 (for `slideLayout1.xml`) when no def is found
    // So all objects are in Layout1 and every slide that references it uses this layout.
    return 1;
}
// XML-GEN: Last 5 functions create root /ppt files
/**
 * Creates `ppt/theme/theme1.xml`
 * @return {string} XML
 */
function makeXmlTheme() {
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + CRLF + "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office Theme\"><a:themeElements><a:clrScheme name=\"Office\"><a:dk1><a:sysClr val=\"windowText\" lastClr=\"000000\"/></a:dk1><a:lt1><a:sysClr val=\"window\" lastClr=\"FFFFFF\"/></a:lt1><a:dk2><a:srgbClr val=\"44546A\"/></a:dk2><a:lt2><a:srgbClr val=\"E7E6E6\"/></a:lt2><a:accent1><a:srgbClr val=\"4472C4\"/></a:accent1><a:accent2><a:srgbClr val=\"ED7D31\"/></a:accent2><a:accent3><a:srgbClr val=\"A5A5A5\"/></a:accent3><a:accent4><a:srgbClr val=\"FFC000\"/></a:accent4><a:accent5><a:srgbClr val=\"5B9BD5\"/></a:accent5><a:accent6><a:srgbClr val=\"70AD47\"/></a:accent6><a:hlink><a:srgbClr val=\"0563C1\"/></a:hlink><a:folHlink><a:srgbClr val=\"954F72\"/></a:folHlink></a:clrScheme><a:fontScheme name=\"Office\"><a:majorFont><a:latin typeface=\"Calibri Light\" panose=\"020F0302020204030204\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/><a:font script=\"Jpan\" typeface=\"\u6E38\u30B4\u30B7\u30C3\u30AF Light\"/><a:font script=\"Hang\" typeface=\"\uB9D1\uC740 \uACE0\uB515\"/><a:font script=\"Hans\" typeface=\"\u7B49\u7EBF Light\"/><a:font script=\"Hant\" typeface=\"\u65B0\u7D30\u660E\u9AD4\"/><a:font script=\"Arab\" typeface=\"Times New Roman\"/><a:font script=\"Hebr\" typeface=\"Times New Roman\"/><a:font script=\"Thai\" typeface=\"Angsana New\"/><a:font script=\"Ethi\" typeface=\"Nyala\"/><a:font script=\"Beng\" typeface=\"Vrinda\"/><a:font script=\"Gujr\" typeface=\"Shruti\"/><a:font script=\"Khmr\" typeface=\"MoolBoran\"/><a:font script=\"Knda\" typeface=\"Tunga\"/><a:font script=\"Guru\" typeface=\"Raavi\"/><a:font script=\"Cans\" typeface=\"Euphemia\"/><a:font script=\"Cher\" typeface=\"Plantagenet Cherokee\"/><a:font script=\"Yiii\" typeface=\"Microsoft Yi Baiti\"/><a:font script=\"Tibt\" typeface=\"Microsoft Himalaya\"/><a:font script=\"Thaa\" typeface=\"MV Boli\"/><a:font script=\"Deva\" typeface=\"Mangal\"/><a:font script=\"Telu\" typeface=\"Gautami\"/><a:font script=\"Taml\" typeface=\"Latha\"/><a:font script=\"Syrc\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Orya\" typeface=\"Kalinga\"/><a:font script=\"Mlym\" typeface=\"Kartika\"/><a:font script=\"Laoo\" typeface=\"DokChampa\"/><a:font script=\"Sinh\" typeface=\"Iskoola Pota\"/><a:font script=\"Mong\" typeface=\"Mongolian Baiti\"/><a:font script=\"Viet\" typeface=\"Times New Roman\"/><a:font script=\"Uigh\" typeface=\"Microsoft Uighur\"/><a:font script=\"Geor\" typeface=\"Sylfaen\"/><a:font script=\"Armn\" typeface=\"Arial\"/><a:font script=\"Bugi\" typeface=\"Leelawadee UI\"/><a:font script=\"Bopo\" typeface=\"Microsoft JhengHei\"/><a:font script=\"Java\" typeface=\"Javanese Text\"/><a:font script=\"Lisu\" typeface=\"Segoe UI\"/><a:font script=\"Mymr\" typeface=\"Myanmar Text\"/><a:font script=\"Nkoo\" typeface=\"Ebrima\"/><a:font script=\"Olck\" typeface=\"Nirmala UI\"/><a:font script=\"Osma\" typeface=\"Ebrima\"/><a:font script=\"Phag\" typeface=\"Phagspa\"/><a:font script=\"Syrn\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Syrj\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Syre\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Sora\" typeface=\"Nirmala UI\"/><a:font script=\"Tale\" typeface=\"Microsoft Tai Le\"/><a:font script=\"Talu\" typeface=\"Microsoft New Tai Lue\"/><a:font script=\"Tfng\" typeface=\"Ebrima\"/></a:majorFont><a:minorFont><a:latin typeface=\"Calibri\" panose=\"020F0502020204030204\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/><a:font script=\"Jpan\" typeface=\"\u6E38\u30B4\u30B7\u30C3\u30AF\"/><a:font script=\"Hang\" typeface=\"\uB9D1\uC740 \uACE0\uB515\"/><a:font script=\"Hans\" typeface=\"\u7B49\u7EBF\"/><a:font script=\"Hant\" typeface=\"\u65B0\u7D30\u660E\u9AD4\"/><a:font script=\"Arab\" typeface=\"Arial\"/><a:font script=\"Hebr\" typeface=\"Arial\"/><a:font script=\"Thai\" typeface=\"Cordia New\"/><a:font script=\"Ethi\" typeface=\"Nyala\"/><a:font script=\"Beng\" typeface=\"Vrinda\"/><a:font script=\"Gujr\" typeface=\"Shruti\"/><a:font script=\"Khmr\" typeface=\"DaunPenh\"/><a:font script=\"Knda\" typeface=\"Tunga\"/><a:font script=\"Guru\" typeface=\"Raavi\"/><a:font script=\"Cans\" typeface=\"Euphemia\"/><a:font script=\"Cher\" typeface=\"Plantagenet Cherokee\"/><a:font script=\"Yiii\" typeface=\"Microsoft Yi Baiti\"/><a:font script=\"Tibt\" typeface=\"Microsoft Himalaya\"/><a:font script=\"Thaa\" typeface=\"MV Boli\"/><a:font script=\"Deva\" typeface=\"Mangal\"/><a:font script=\"Telu\" typeface=\"Gautami\"/><a:font script=\"Taml\" typeface=\"Latha\"/><a:font script=\"Syrc\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Orya\" typeface=\"Kalinga\"/><a:font script=\"Mlym\" typeface=\"Kartika\"/><a:font script=\"Laoo\" typeface=\"DokChampa\"/><a:font script=\"Sinh\" typeface=\"Iskoola Pota\"/><a:font script=\"Mong\" typeface=\"Mongolian Baiti\"/><a:font script=\"Viet\" typeface=\"Arial\"/><a:font script=\"Uigh\" typeface=\"Microsoft Uighur\"/><a:font script=\"Geor\" typeface=\"Sylfaen\"/><a:font script=\"Armn\" typeface=\"Arial\"/><a:font script=\"Bugi\" typeface=\"Leelawadee UI\"/><a:font script=\"Bopo\" typeface=\"Microsoft JhengHei\"/><a:font script=\"Java\" typeface=\"Javanese Text\"/><a:font script=\"Lisu\" typeface=\"Segoe UI\"/><a:font script=\"Mymr\" typeface=\"Myanmar Text\"/><a:font script=\"Nkoo\" typeface=\"Ebrima\"/><a:font script=\"Olck\" typeface=\"Nirmala UI\"/><a:font script=\"Osma\" typeface=\"Ebrima\"/><a:font script=\"Phag\" typeface=\"Phagspa\"/><a:font script=\"Syrn\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Syrj\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Syre\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Sora\" typeface=\"Nirmala UI\"/><a:font script=\"Tale\" typeface=\"Microsoft Tai Le\"/><a:font script=\"Talu\" typeface=\"Microsoft New Tai Lue\"/><a:font script=\"Tfng\" typeface=\"Ebrima\"/></a:minorFont></a:fontScheme><a:fmtScheme name=\"Office\"><a:fillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"110000\"/><a:satMod val=\"105000\"/><a:tint val=\"67000\"/></a:schemeClr></a:gs><a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"105000\"/><a:satMod val=\"103000\"/><a:tint val=\"73000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"105000\"/><a:satMod val=\"109000\"/><a:tint val=\"81000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:satMod val=\"103000\"/><a:lumMod val=\"102000\"/><a:tint val=\"94000\"/></a:schemeClr></a:gs><a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:satMod val=\"110000\"/><a:lumMod val=\"100000\"/><a:shade val=\"100000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"99000\"/><a:satMod val=\"120000\"/><a:shade val=\"78000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w=\"6350\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln><a:ln w=\"12700\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln><a:ln w=\"19050\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad=\"57150\" dist=\"19050\" dir=\"5400000\" algn=\"ctr\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"63000\"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:solidFill><a:schemeClr val=\"phClr\"><a:tint val=\"95000\"/><a:satMod val=\"170000\"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"93000\"/><a:satMod val=\"150000\"/><a:shade val=\"98000\"/><a:lumMod val=\"102000\"/></a:schemeClr></a:gs><a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:tint val=\"98000\"/><a:satMod val=\"130000\"/><a:shade val=\"90000\"/><a:lumMod val=\"103000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"63000\"/><a:satMod val=\"120000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/><a:extLst><a:ext uri=\"{05A4C25C-085E-4340-85A3-A5531E510DB2}\"><thm15:themeFamily xmlns:thm15=\"http://schemas.microsoft.com/office/thememl/2012/main\" name=\"Office Theme\" id=\"{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}\" vid=\"{4A3C46E8-61CC-4603-A589-7422A47A8E4A}\"/></a:ext></a:extLst></a:theme>";
}
/**
 * Create presentation file (`ppt/presentation.xml`)
 * @see https://docs.microsoft.com/en-us/office/open-xml/structure-of-a-presentationml-document
 * @see http://www.datypic.com/sc/ooxml/t-p_CT_Presentation.html
 * @param {ISlide[]} slides - array of slides
 * @param {ILayout} pptLayout - presentation layout
 * @param {boolean} rtlMode - RTL mode
 * @return {string} XML
 */
function makeXmlPresentation(slides, pptLayout, rtlMode) {
    var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        CRLF +
        '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
        (rtlMode ? 'rtl="1" ' : '') +
        'saveSubsetFonts="1" autoCompressPictures="0">';
    // IMPORTANT: Steps 1-2-3 must be in this order or PPT will give corruption message on open!
    // STEP 1: Add slide master
    strXml += '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>';
    // STEP 2: Add all Slides
    strXml += '<p:sldIdLst>';
    for (var idx = 0; idx < slides.length; idx++) {
        strXml += '<p:sldId id="' + (idx + 256) + '" r:id="rId' + (idx + 2) + '"/>';
    }
    strXml += '</p:sldIdLst>';
    // STEP 3: Add Notes Master (NOTE: length+2 is from `presentation.xml.rels` func (since we have to match this rId, we just use same logic))
    strXml += '<p:notesMasterIdLst><p:notesMasterId r:id="rId' + (slides.length + 2) + '"/></p:notesMasterIdLst>';
    // STEP 4: Build SLIDE text styles
    strXml +=
        '<p:sldSz cx="' +
            pptLayout.width +
            '" cy="' +
            pptLayout.height +
            '"/>' +
            '<p:notesSz cx="' +
            pptLayout.height +
            '" cy="' +
            pptLayout.width +
            '"/>' +
            '<p:defaultTextStyle>'; //+'<a:defPPr><a:defRPr lang="en-US"/></a:defPPr>'
    for (var idx = 1; idx < 10; idx++) {
        strXml +=
            '<a:lvl' +
                idx +
                'pPr marL="' +
                (idx - 1) * 457200 +
                '" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">' +
                '<a:defRPr sz="1800" kern="1200">' +
                '<a:solidFill><a:schemeClr val="tx1"/></a:solidFill>' +
                '<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/>' +
                '</a:defRPr>' +
                '</a:lvl' +
                idx +
                'pPr>';
    }
    strXml += '</p:defaultTextStyle>';
    strXml += '</p:presentation>';
    return strXml;
}
/**
 * Create `ppt/presProps.xml`
 * @return {string} XML
 */
function makeXmlPresProps() {
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + CRLF + "<p:presentationPr xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"/>";
}
/**
 * Create `ppt/tableStyles.xml`
 * @see: http://openxmldeveloper.org/discussions/formats/f/13/p/2398/8107.aspx
 * @return {string} XML
 */
function makeXmlTableStyles() {
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + CRLF + "<a:tblStyleLst xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" def=\"{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}\"/>";
}
/**
 * Creates `ppt/viewProps.xml`
 * @return {string} XML
 */
function makeXmlViewProps() {
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + CRLF + "<p:viewPr xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:normalViewPr horzBarState=\"maximized\"><p:restoredLeft sz=\"15611\"/><p:restoredTop sz=\"94610\"/></p:normalViewPr><p:slideViewPr><p:cSldViewPr snapToGrid=\"0\" snapToObjects=\"1\"><p:cViewPr varScale=\"1\"><p:scale><a:sx n=\"136\" d=\"100\"/><a:sy n=\"136\" d=\"100\"/></p:scale><p:origin x=\"216\" y=\"312\"/></p:cViewPr><p:guideLst/></p:cSldViewPr></p:slideViewPr><p:notesTextViewPr><p:cViewPr><p:scale><a:sx n=\"1\" d=\"1\"/><a:sy n=\"1\" d=\"1\"/></p:scale><p:origin x=\"0\" y=\"0\"/></p:cViewPr></p:notesTextViewPr><p:gridSpacing cx=\"76200\" cy=\"76200\"/></p:viewPr>";
}
/**
 * Checks shadow options passed by user and performs corrections if needed.
 * @param {IShadowOptions} IShadowOptions - shadow options
 */
function correctShadowOptions(IShadowOptions) {
    if (!IShadowOptions || IShadowOptions === null)
        return;
    // OPT: `type`
    if (IShadowOptions.type !== 'outer' && IShadowOptions.type !== 'inner' && IShadowOptions.type !== 'none') {
        console.warn('Warning: shadow.type options are `outer`, `inner` or `none`.');
        IShadowOptions.type = 'outer';
    }
    // OPT: `angle`
    if (IShadowOptions.angle) {
        // A: REALITY-CHECK
        if (isNaN(Number(IShadowOptions.angle)) || IShadowOptions.angle < 0 || IShadowOptions.angle > 359) {
            console.warn('Warning: shadow.angle can only be 0-359');
            IShadowOptions.angle = 270;
        }
        // B: ROBUST: Cast any type of valid arg to int: '12', 12.3, etc. -> 12
        IShadowOptions.angle = Math.round(Number(IShadowOptions.angle));
    }
    // OPT: `opacity`
    if (IShadowOptions.opacity) {
        // A: REALITY-CHECK
        if (isNaN(Number(IShadowOptions.opacity)) || IShadowOptions.opacity < 0 || IShadowOptions.opacity > 1) {
            console.warn('Warning: shadow.opacity can only be 0-1');
            IShadowOptions.opacity = 0.75;
        }
        // B: ROBUST: Cast any type of valid arg to int: '12', 12.3, etc. -> 12
        IShadowOptions.opacity = Number(IShadowOptions.opacity);
    }
}
function getShapeInfo(shapeName) {
    if (!shapeName)
        return PowerPointShapes.RECTANGLE;
    if (typeof shapeName === 'object' && shapeName.name && shapeName.displayName && shapeName.avLst)
        return shapeName;
    if (PowerPointShapes[shapeName])
        return PowerPointShapes[shapeName];
    var objShape = Object.keys(PowerPointShapes).filter(function (key) {
        return PowerPointShapes[key].name === shapeName || PowerPointShapes[key].displayName;
    })[0];
    if (typeof objShape !== 'undefined' && objShape !== null)
        return objShape;
    return PowerPointShapes.RECTANGLE;
}

/**
 * PptxGenJS: Slide object generators
 */
/** counter for included charts (used for index in their filenames) */
var _chartCounter = 0;
/**
 * Transforms a slide definition to a slide object that is then passed to the XML transformation process.
 * @param {ISlideMasterOptions} slideDef - slide definition
 * @param {ISlide|ISlideLayout} target - empty slide object that should be updated by the passed definition
 */
function createSlideObject(slideDef, target) {
    // STEP 1: Add background
    if (slideDef.bkgd) {
        addBackgroundDefinition(slideDef.bkgd, target);
    }
    // STEP 2: Add all Slide Master objects in the order they were given (Issue#53)
    if (slideDef.objects && Array.isArray(slideDef.objects) && slideDef.objects.length > 0) {
        slideDef.objects.forEach(function (object, idx) {
            var key = Object.keys(object)[0];
            var tgt = target;
            if (MASTER_OBJECTS[key] && key === 'chart')
                addChartDefinition(tgt, object[key].type, object[key].data, object[key].opts);
            else if (MASTER_OBJECTS[key] && key === 'image')
                addImageDefinition(tgt, object[key]);
            else if (MASTER_OBJECTS[key] && key === 'line')
                addShapeDefinition(tgt, BASE_SHAPES.LINE, object[key]);
            else if (MASTER_OBJECTS[key] && key === 'rect')
                addShapeDefinition(tgt, BASE_SHAPES.RECTANGLE, object[key]);
            else if (MASTER_OBJECTS[key] && key === 'text')
                addTextDefinition(tgt, object[key].text, object[key].options, false);
            else if (MASTER_OBJECTS[key] && key === 'placeholder') {
                // TODO: 20180820: Check for existing `name`?
                object[key].options.placeholder = object[key].options.name;
                delete object[key].options.name; // remap name for earier handling internally
                object[key].options.placeholderType = object[key].options.type;
                delete object[key].options.type; // remap name for earier handling internally
                object[key].options.placeholderIdx = 100 + idx;
                addPlaceholderDefinition(tgt, object[key].text, object[key].options);
            }
        });
    }
    // STEP 3: Add Slide Numbers (NOTE: Do this last so numbers are not covered by objects!)
    if (slideDef.slideNumber && typeof slideDef.slideNumber === 'object') {
        target.slideNumberObj = slideDef.slideNumber;
    }
}
/**
 * Generate the chart based on input data.
 * OOXML Chart Spec: ISO/IEC 29500-1:2016(E)
 *
 * @param {CHART_TYPE_NAMES | IChartMulti[]} `type` should belong to: 'column', 'pie'
 * @param {[]} `data` a JSON object with follow the following format
 * @param {IChartOpts} `opt` chart options
 * @param {ISlide} `target` slide object that the chart will be added to
 * @return {object} chart object
 * {
 *   title: 'eSurvey chart',
 *   data: [
 *		{
 *			name: 'Income',
 *			labels: ['2005', '2006', '2007', '2008', '2009'],
 *			values: [23.5, 26.2, 30.1, 29.5, 24.6]
 *		},
 *		{
 *			name: 'Expense',
 *			labels: ['2005', '2006', '2007', '2008', '2009'],
 *			values: [18.1, 22.8, 23.9, 25.1, 25]
 *		}
 *	 ]
 *	}
 */
function addChartDefinition(target, type, data, opt) {
    function correctGridLineOptions(glOpts) {
        if (!glOpts || glOpts.style === 'none')
            return;
        if (glOpts.size !== undefined && (isNaN(Number(glOpts.size)) || glOpts.size <= 0)) {
            console.warn('Warning: chart.gridLine.size must be greater than 0.');
            delete glOpts.size; // delete prop to used defaults
        }
        if (glOpts.style && ['solid', 'dash', 'dot'].indexOf(glOpts.style) < 0) {
            console.warn('Warning: chart.gridLine.style options: `solid`, `dash`, `dot`.');
            delete glOpts.style;
        }
    }
    var chartId = ++_chartCounter;
    var resultObject = {
        type: null,
        text: null,
        options: null,
        chartRid: null,
    };
    // DESIGN: `type` can an object (ex: `pptx.charts.DOUGHNUT`) or an array of chart objects
    // EX: addChartDefinition([ { type:pptx.charts.BAR, data:{name:'', labels:[], values[]} }, {<etc>} ])
    // Multi-Type Charts
    var tmpOpt;
    var tmpData = [], options;
    if (Array.isArray(type)) {
        // For multi-type charts there needs to be data for each type,
        // as well as a single data source for non-series operations.
        // The data is indexed below to keep the data in order when segmented
        // into types.
        type.forEach(function (obj) {
            tmpData = tmpData.concat(obj.data);
        });
        tmpOpt = data || opt;
    }
    else {
        tmpData = data;
        tmpOpt = opt;
    }
    tmpData.forEach(function (item, i) {
        item.index = i;
    });
    options = tmpOpt && typeof tmpOpt === 'object' ? tmpOpt : {};
    // STEP 1: TODO: check for reqd fields, correct type, etc
    // `type` exists in CHART_TYPES
    // Array.isArray(data)
    /*
        if ( Array.isArray(rel.data) && rel.data.length > 0 && typeof rel.data[0] === 'object'
            && rel.data[0].labels && Array.isArray(rel.data[0].labels)
            && rel.data[0].values && Array.isArray(rel.data[0].values) ) {
            obj = rel.data[0];
        }
        else {
            console.warn("USAGE: addChart( 'pie', [ {name:'Sales', labels:['Jan','Feb'], values:[10,20]} ], {x:1, y:1} )");
            return;
        }
        */
    // STEP 2: Set default options/decode user options
    // A: Core
    options.type = type;
    options.x = typeof options.x !== 'undefined' && options.x != null && !isNaN(Number(options.x)) ? options.x : 1;
    options.y = typeof options.y !== 'undefined' && options.y != null && !isNaN(Number(options.y)) ? options.y : 1;
    options.w = options.w || '50%';
    options.h = options.h || '50%';
    // B: Options: misc
    if (['bar', 'col'].indexOf(options.barDir || '') < 0)
        options.barDir = 'col';
    // IMPORTANT: 'bestFit' will cause issues with PPT-Online in some cases, so defualt to 'ctr'!
    if (['bestFit', 'b', 'ctr', 'inBase', 'inEnd', 'l', 'outEnd', 'r', 't'].indexOf(options.dataLabelPosition || '') < 0)
        options.dataLabelPosition = options.type === CHART_TYPES.PIE || options.type === CHART_TYPES.DOUGHNUT ? 'bestFit' : 'ctr';
    options.dataLabelBkgrdColors = options.dataLabelBkgrdColors === true || options.dataLabelBkgrdColors === false ? options.dataLabelBkgrdColors : false;
    if (['b', 'l', 'r', 't', 'tr'].indexOf(options.legendPos || '') < 0)
        options.legendPos = 'r';
    // barGrouping: "21.2.3.17 ST_Grouping (Grouping)"
    if (['clustered', 'standard', 'stacked', 'percentStacked'].indexOf(options.barGrouping || '') < 0)
        options.barGrouping = 'standard';
    if (options.barGrouping.indexOf('tacked') > -1) {
        options.dataLabelPosition = 'ctr'; // IMPORTANT: PPT-Online will not open Presentation when 'outEnd' etc is used on stacked!
        if (!options.barGapWidthPct)
            options.barGapWidthPct = 50;
    }
    // 3D bar: ST_Shape
    if (['cone', 'coneToMax', 'box', 'cylinder', 'pyramid', 'pyramidToMax'].indexOf(options.bar3DShape || '') < 0)
        options.bar3DShape = 'box';
    // lineDataSymbol: http://www.datypic.com/sc/ooxml/a-val-32.html
    // Spec has [plus,star,x] however neither PPT2013 nor PPT-Online support them
    if (['circle', 'dash', 'diamond', 'dot', 'none', 'square', 'triangle'].indexOf(options.lineDataSymbol || '') < 0)
        options.lineDataSymbol = 'circle';
    if (['gap', 'span'].indexOf(options.displayBlanksAs || '') < 0)
        options.displayBlanksAs = 'span';
    if (['standard', 'marker', 'filled'].indexOf(options.radarStyle || '') < 0)
        options.radarStyle = 'standard';
    options.lineDataSymbolSize = options.lineDataSymbolSize && !isNaN(options.lineDataSymbolSize) ? options.lineDataSymbolSize : 6;
    options.lineDataSymbolLineSize = options.lineDataSymbolLineSize && !isNaN(options.lineDataSymbolLineSize) ? options.lineDataSymbolLineSize * ONEPT : 0.75 * ONEPT;
    // `layout` allows the override of PPT defaults to maximize space
    if (options.layout) {
        ['x', 'y', 'w', 'h'].forEach(function (key) {
            var val = options.layout[key];
            if (isNaN(Number(val)) || val < 0 || val > 1) {
                console.warn('Warning: chart.layout.' + key + ' can only be 0-1');
                delete options.layout[key]; // remove invalid value so that default will be used
            }
        });
    }
    // Set gridline defaults
    options.catGridLine = options.catGridLine || (options.type === CHART_TYPES.SCATTER ? { color: 'D9D9D9', size: 1 } : { style: 'none' });
    options.valGridLine = options.valGridLine || (options.type === CHART_TYPES.SCATTER ? { color: 'D9D9D9', size: 1 } : {});
    options.serGridLine = options.serGridLine || (options.type === CHART_TYPES.SCATTER ? { color: 'D9D9D9', size: 1 } : { style: 'none' });
    correctGridLineOptions(options.catGridLine);
    correctGridLineOptions(options.valGridLine);
    correctGridLineOptions(options.serGridLine);
    correctShadowOptions(options.shadow);
    // C: Options: plotArea
    options.showDataTable = options.showDataTable === true || options.showDataTable === false ? options.showDataTable : false;
    options.showDataTableHorzBorder = options.showDataTableHorzBorder === true || options.showDataTableHorzBorder === false ? options.showDataTableHorzBorder : true;
    options.showDataTableVertBorder = options.showDataTableVertBorder === true || options.showDataTableVertBorder === false ? options.showDataTableVertBorder : true;
    options.showDataTableOutline = options.showDataTableOutline === true || options.showDataTableOutline === false ? options.showDataTableOutline : true;
    options.showDataTableKeys = options.showDataTableKeys === true || options.showDataTableKeys === false ? options.showDataTableKeys : true;
    options.showLabel = options.showLabel === true || options.showLabel === false ? options.showLabel : false;
    options.showLegend = options.showLegend === true || options.showLegend === false ? options.showLegend : false;
    options.showPercent = options.showPercent === true || options.showPercent === false ? options.showPercent : true;
    options.showTitle = options.showTitle === true || options.showTitle === false ? options.showTitle : false;
    options.showValue = options.showValue === true || options.showValue === false ? options.showValue : false;
    options.catAxisLineShow = typeof options.catAxisLineShow !== 'undefined' ? options.catAxisLineShow : true;
    options.valAxisLineShow = typeof options.valAxisLineShow !== 'undefined' ? options.valAxisLineShow : true;
    options.serAxisLineShow = typeof options.serAxisLineShow !== 'undefined' ? options.serAxisLineShow : true;
    options.v3DRotX = !isNaN(options.v3DRotX) && options.v3DRotX >= -90 && options.v3DRotX <= 90 ? options.v3DRotX : 30;
    options.v3DRotY = !isNaN(options.v3DRotY) && options.v3DRotY >= 0 && options.v3DRotY <= 360 ? options.v3DRotY : 30;
    options.v3DRAngAx = options.v3DRAngAx === true || options.v3DRAngAx === false ? options.v3DRAngAx : true;
    options.v3DPerspective = !isNaN(options.v3DPerspective) && options.v3DPerspective >= 0 && options.v3DPerspective <= 240 ? options.v3DPerspective : 30;
    // D: Options: chart
    options.barGapWidthPct = !isNaN(options.barGapWidthPct) && options.barGapWidthPct >= 0 && options.barGapWidthPct <= 1000 ? options.barGapWidthPct : 150;
    options.barGapDepthPct = !isNaN(options.barGapDepthPct) && options.barGapDepthPct >= 0 && options.barGapDepthPct <= 1000 ? options.barGapDepthPct : 150;
    options.chartColors = Array.isArray(options.chartColors)
        ? options.chartColors
        : options.type === CHART_TYPES.PIE || options.type === CHART_TYPES.DOUGHNUT
            ? PIECHART_COLORS
            : BARCHART_COLORS;
    options.chartColorsOpacity = options.chartColorsOpacity && !isNaN(options.chartColorsOpacity) ? options.chartColorsOpacity : null;
    //
    options.border = options.border && typeof options.border === 'object' ? options.border : null;
    if (options.border && (!options.border.pt || isNaN(options.border.pt)))
        options.border.pt = 1;
    if (options.border && (!options.border.color || typeof options.border.color !== 'string' || options.border.color.length !== 6))
        options.border.color = '363636';
    //
    options.dataBorder = options.dataBorder && typeof options.dataBorder === 'object' ? options.dataBorder : null;
    if (options.dataBorder && (!options.dataBorder.pt || isNaN(options.dataBorder.pt)))
        options.dataBorder.pt = 0.75;
    if (options.dataBorder && (!options.dataBorder.color || typeof options.dataBorder.color !== 'string' || options.dataBorder.color.length !== 6))
        options.dataBorder.color = 'F9F9F9';
    //
    if (!options.dataLabelFormatCode && options.type === CHART_TYPES.SCATTER)
        options.dataLabelFormatCode = 'General';
    options.dataLabelFormatCode =
        options.dataLabelFormatCode && typeof options.dataLabelFormatCode === 'string'
            ? options.dataLabelFormatCode
            : options.type === CHART_TYPES.PIE || options.type === CHART_TYPES.DOUGHNUT
                ? '0%'
                : '#,##0';
    //
    // Set default format for Scatter chart labels to custom string if not defined
    if (!options.dataLabelFormatScatter && options.type === CHART_TYPES.SCATTER)
        options.dataLabelFormatScatter = 'custom';
    //
    options.lineSize = typeof options.lineSize === 'number' ? options.lineSize : 2;
    options.valAxisMajorUnit = typeof options.valAxisMajorUnit === 'number' ? options.valAxisMajorUnit : null;
    options.valAxisCrossesAt = options.valAxisCrossesAt || 'autoZero';
    // STEP 4: Set props
    resultObject.type = 'chart';
    resultObject.options = options;
    resultObject.chartRid = target.relsChart.length + 1;
    // STEP 5: Add this chart to this Slide Rels (rId/rels count spans all slides! Count all images to get next rId)
    target.relsChart.push({
        rId: target.relsChart.length + 1,
        data: tmpData,
        opts: options,
        type: options.type,
        globalId: chartId,
        fileName: 'chart' + chartId + '.xml',
        Target: '/ppt/charts/chart' + chartId + '.xml',
    });
    target.data.push(resultObject);
    return resultObject;
}
/**
 * Adds an image object to a slide definition.
 * This method can be called with only two args (opt, target) - this is supposed to be the only way in future.
 * @param {IImageOpts} `opt` - object containing `path`/`data`, `x`, `y`, etc.
 * @param {ISlide} `target` - slide that the image should be added to (if not specified as the 2nd arg)
 */
function addImageDefinition(target, opt) {
    var newObject = {
        type: null,
        text: null,
        options: null,
        image: null,
        imageRid: null,
        hyperlink: null,
    };
    // FIRST: Set vars for this image (object param replaces positional args in 1.1.0)
    var intPosX = opt.x || 0;
    var intPosY = opt.y || 0;
    var intWidth = opt.w || 0;
    var intHeight = opt.h || 0;
    var sizing = opt.sizing || null;
    var objHyperlink = opt.hyperlink || '';
    var strImageData = opt.data || '';
    var strImagePath = opt.path || '';
    var imageRelId = target.rels.length + target.relsChart.length + target.relsMedia.length + 1;
    // REALITY-CHECK:
    if (!strImagePath && !strImageData) {
        console.error("ERROR: `addImage()` requires either 'data' or 'path' parameter!");
        return null;
    }
    else if (strImageData && strImageData.toLowerCase().indexOf('base64,') === -1) {
        console.error("ERROR: Image `data` value lacks a base64 header! Ex: 'image/png;base64,NMP[...]')");
        return null;
    }
    // STEP 1: Set extension
    // NOTE: Split to address URLs with params (eg: `path/brent.jpg?someParam=true`)
    var strImgExtn = strImagePath
        .substring(strImagePath.lastIndexOf('/') + 1)
        .split('?')[0]
        .split('.')
        .pop()
        .split('#')[0] || 'png';
    // However, pre-encoded images can be whatever mime-type they want (and good for them!)
    if (strImageData && /image\/(\w+)\;/.exec(strImageData) && /image\/(\w+)\;/.exec(strImageData).length > 0) {
        strImgExtn = /image\/(\w+)\;/.exec(strImageData)[1];
    }
    else if (strImageData && strImageData.toLowerCase().indexOf('image/svg+xml') > -1) {
        strImgExtn = 'svg';
    }
    // STEP 2: Set type/path
    newObject.type = 'image';
    newObject.image = strImagePath || 'preencoded.png';
    // STEP 3: Set image properties & options
    // FIXME: Measure actual image when no intWidth/intHeight params passed
    // ....: This is an async process: we need to make getSizeFromImage use callback, then set H/W...
    // if ( !intWidth || !intHeight ) { var imgObj = getSizeFromImage(strImagePath);
    newObject.options = {
        x: intPosX || 0,
        y: intPosY || 0,
        w: intWidth || 1,
        h: intHeight || 1,
        rounding: typeof opt.rounding === 'boolean' ? opt.rounding : false,
        sizing: sizing,
        placeholder: opt.placeholder,
    };
    // STEP 4: Add this image to this Slide Rels (rId/rels count spans all slides! Count all images to get next rId)
    if (strImgExtn === 'svg') {
        // SVG files consume *TWO* rId's: (a png version and the svg image)
        // <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
        // <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.svg"/>
        target.relsMedia.push({
            path: strImagePath || strImageData + 'png',
            type: 'image/png',
            extn: 'png',
            data: strImageData || '',
            rId: imageRelId,
            Target: '../media/image-' + target.number + '-' + (target.relsMedia.length + 1) + '.png',
            isSvgPng: true,
            svgSize: { w: newObject.options.w, h: newObject.options.h },
        });
        newObject.imageRid = imageRelId;
        target.relsMedia.push({
            path: strImagePath || strImageData,
            type: 'image/svg+xml',
            extn: strImgExtn,
            data: strImageData || '',
            rId: imageRelId + 1,
            Target: '../media/image-' + target.number + '-' + (target.relsMedia.length + 1) + '.' + strImgExtn,
        });
        newObject.imageRid = imageRelId + 1;
    }
    else {
        target.relsMedia.push({
            path: strImagePath || 'preencoded.' + strImgExtn,
            type: 'image/' + strImgExtn,
            extn: strImgExtn,
            data: strImageData || '',
            rId: imageRelId,
            Target: '../media/image-' + target.number + '-' + (target.relsMedia.length + 1) + '.' + strImgExtn,
        });
        newObject.imageRid = imageRelId;
    }
    // STEP 5: Hyperlink support
    if (typeof objHyperlink === 'object') {
        if (!objHyperlink.url && !objHyperlink.slide)
            throw new Error('ERROR: `hyperlink` option requires either: `url` or `slide`');
        else {
            imageRelId++;
            target.rels.push({
                type: SLIDE_OBJECT_TYPES.hyperlink,
                data: objHyperlink.slide ? 'slide' : 'dummy',
                rId: imageRelId,
                Target: objHyperlink.url || objHyperlink.slide.toString(),
            });
            objHyperlink.rId = imageRelId;
            newObject.hyperlink = objHyperlink;
        }
    }
    // STEP 6: Add object to slide
    target.data.push(newObject);
}
/**
 * Adds a media object to a slide definition.
 * @param {ISlide} `target` - slide object that the text will be added to
 * @param {IMediaOpts} `opt` - media options
 */
function addMediaDefinition(target, opt) {
    var intRels = target.relsMedia.length + 1;
    var intPosX = opt.x || 0;
    var intPosY = opt.y || 0;
    var intSizeX = opt.w || 2;
    var intSizeY = opt.h || 2;
    var strData = opt.data || '';
    var strLink = opt.link || '';
    var strPath = opt.path || '';
    var strType = opt.type || 'audio';
    var strExtn = 'mp3';
    var slideData = {
        type: SLIDE_OBJECT_TYPES.media,
    };
    // STEP 1: REALITY-CHECK
    if (!strPath && !strData && strType !== 'online') {
        throw "addMedia() error: either 'data' or 'path' are required!";
    }
    else if (strData && strData.toLowerCase().indexOf('base64,') === -1) {
        throw "addMedia() error: `data` value lacks a base64 header! Ex: 'video/mpeg;base64,NMP[...]')";
    }
    // Online Video: requires `link`
    if (strType === 'online' && !strLink) {
        throw 'addMedia() error: online videos require `link` value';
    }
    // FIXME: 20190707
    //strType = strData ? strData.split(';')[0].split('/')[0] : strType
    strExtn = strData ? strData.split(';')[0].split('/')[1] : strPath.split('.').pop();
    // STEP 2: Set type, media
    slideData.mtype = strType;
    slideData.media = strPath || 'preencoded.mov';
    slideData.options = {};
    // STEP 3: Set media properties & options
    slideData.options.x = intPosX;
    slideData.options.y = intPosY;
    slideData.options.w = intSizeX;
    slideData.options.h = intSizeY;
    // STEP 4: Add this media to this Slide Rels (rId/rels count spans all slides! Count all media to get next rId)
    // NOTE: rId starts at 2 (hence the intRels+1 below) as slideLayout.xml is rId=1!
    if (strType === 'online') {
        // A: Add video
        target.relsMedia.push({
            path: strPath || 'preencoded' + strExtn,
            data: 'dummy',
            type: 'online',
            extn: strExtn,
            rId: intRels + 1,
            Target: strLink,
        });
        slideData.mediaRid = target.relsMedia[target.relsMedia.length - 1].rId;
        // B: Add preview/overlay image
        target.relsMedia.push({
            path: 'preencoded.png',
            data: IMG_PLAYBTN,
            type: 'image/png',
            extn: 'png',
            rId: intRels + 2,
            Target: '../media/image-' + target.number + '-' + (target.relsMedia.length + 1) + '.png',
        });
    }
    else {
        /* NOTE: Audio/Video files consume *TWO* rId's:
         * <Relationship Id="rId2" Target="../media/media1.mov" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"/>
         * <Relationship Id="rId3" Target="../media/media1.mov" Type="http://schemas.microsoft.com/office/2007/relationships/media"/>
         */
        // A: "relationships/video"
        target.relsMedia.push({
            path: strPath || 'preencoded' + strExtn,
            type: strType + '/' + strExtn,
            extn: strExtn,
            data: strData || '',
            rId: intRels + 0,
            Target: '../media/media-' + target.number + '-' + (target.relsMedia.length + 1) + '.' + strExtn,
        });
        slideData.mediaRid = target.relsMedia[target.relsMedia.length - 1].rId;
        // B: "relationships/media"
        target.relsMedia.push({
            path: strPath || 'preencoded' + strExtn,
            type: strType + '/' + strExtn,
            extn: strExtn,
            data: strData || '',
            rId: intRels + 1,
            Target: '../media/media-' + target.number + '-' + (target.relsMedia.length + 0) + '.' + strExtn,
        });
        // C: Add preview/overlay image
        target.relsMedia.push({
            data: IMG_PLAYBTN,
            path: 'preencoded.png',
            type: 'image/png',
            extn: 'png',
            rId: intRels + 2,
            Target: '../media/image-' + target.number + '-' + (target.relsMedia.length + 1) + '.png',
        });
    }
    // LAST
    target.data.push(slideData);
}
/**
 * Adds Notes to a slide.
 * @param {String} `notes`
 * @param {Object} opt (*unused*)
 * @param {ISlide} `target` slide object
 * @since 2.3.0
 */
function addNotesDefinition(target, notes) {
    target.data.push({
        type: SLIDE_OBJECT_TYPES.notes,
        text: notes,
    });
}
/**
 * Adds a placeholder object to a slide definition.
 * @param {String} `text`
 * @param {Object} `opt`
 * @param {ISlide} `target` slide object that the placeholder should be added to
 */
function addPlaceholderDefinition(target, text, opt) {
    // FIXME: there are several tpyes - not all placeholders are text!
    // but it seems to work (see below) - INVESTIGATE: how it s/b written
    return addTextDefinition(target, text, opt, true);
    /*
    this works, albeit not for masters, - it UNDOCUMENTED (oops) and why is type=body (s/b image?), or if we do use body as the locale (like title), than whats 'image' for?
    slide4.addImage({ placeholder:'body', path:(NODEJS ? gPaths.ccLogo.path.replace(/http.+\/examples/, '../common') : gPaths.ccLogo.path) });

    // TODO: TODO-3: this has never worked
    // https://github.com/gitbrent/PptxGenJS/issues/599
    if (opt.type === PLACEHOLDER_TYPES.title || opt.type === PLACEHOLDER_TYPES.body) return addTextDefinition(target, text, opt, true)
    else if (opt.type === PLACEHOLDER_TYPES.image ) return addImageDefinition(target, opt)
    */
}
/**
 * Adds a shape object to a slide definition.
 * @param {IShape} shape shape const object (pptx.shapes)
 * @param {IShapeOptions} opt
 * @param {ISlide} target slide object that the shape should be added to
 */
function addShapeDefinition(target, shape, opt) {
    var options = typeof opt === 'object' ? opt : {};
    var newObject = {
        type: SLIDE_OBJECT_TYPES.text,
        shape: shape,
        options: options,
        text: null,
    };
    // 1: Reality check
    if (!shape || typeof shape !== 'object')
        throw 'Missing/Invalid shape parameter! Example: `addShape(pptx.shapes.LINE, {x:1, y:1, w:1, h:1});`';
    // 2: Set options defaults
    options.x = options.x || (options.x === 0 ? 0 : 1);
    options.y = options.y || (options.y === 0 ? 0 : 1);
    options.w = options.w || (options.w === 0 ? 0 : 1);
    options.h = options.h || (options.h === 0 ? 0 : 1);
    options.line = options.line || (shape.name === 'line' ? '333333' : null);
    options.lineSize = options.lineSize || (shape.name === 'line' ? 1 : null);
    if (['dash', 'dashDot', 'lgDash', 'lgDashDot', 'lgDashDotDot', 'solid', 'sysDash', 'sysDot'].indexOf(options.lineDash || '') < 0)
        options.lineDash = 'solid';
    // 3: Add object to slide
    target.data.push(newObject);
}
/**
 * Adds a table object to a slide definition.
 * @param {ISlide} target - slide object that the table should be added to
 * @param {TableRow[]} arrTabRows - table data
 * @param {ITableOptions} inOpt - table options
 * @param {ISlideLayout} slideLayout - Slide layout
 * @param {ILayout} presLayout - Presenation layout
 * @param {Function} addSlide - method
 * @param {Function} getSlide - method
 */
function addTableDefinition(target, tableRows, options, slideLayout, presLayout, addSlide, getSlide) {
    var opt = options && typeof options === 'object' ? options : {};
    var slides = [target]; // Create array of Slides as more may be added by auto-paging
    // STEP 1: REALITY-CHECK
    {
        // A: check for empty
        if (tableRows === null || tableRows.length === 0 || !Array.isArray(tableRows)) {
            throw "addTable: Array expected! EX: 'slide.addTable( [rows], {options} );' (https://gitbrent.github.io/PptxGenJS/docs/api-tables.html)";
        }
        // B: check for non-well-formatted array (ex: rows=['a','b'] instead of [['a','b']])
        if (!tableRows[0] || !Array.isArray(tableRows[0])) {
            throw "addTable: 'rows' should be an array of cells! EX: 'slide.addTable( [ ['A'], ['B'], {text:'C',options:{align:'center'}} ] );' (https://gitbrent.github.io/PptxGenJS/docs/api-tables.html)";
        }
    }
    // STEP 2: Transform `tableRows` into well-formatted ITableCell's
    // tableRows can be object or plain text array: `[{text:'cell 1'}, {text:'cell 2', options:{color:'ff0000'}}]` | `["cell 1", "cell 2"]`
    var arrRows = [];
    tableRows.forEach(function (row) {
        var newRow = [];
        if (Array.isArray(row)) {
            row.forEach(function (cell) {
                var newCell = {
                    type: SLIDE_OBJECT_TYPES.tablecell,
                    text: '',
                    options: typeof cell === 'object' ? cell.options : null,
                };
                if (typeof cell === 'string' || typeof cell === 'number')
                    newCell.text = cell.toString();
                else if (cell.text) {
                    // Cell can contain complex text type, or string, or number
                    if (typeof cell.text === 'string' || typeof cell.text === 'number')
                        newCell.text = cell.text.toString();
                    else if (cell.text)
                        newCell.text = cell.text;
                    // Capture options
                    if (cell.options)
                        newCell.options = cell.options;
                }
                newRow.push(newCell);
            });
        }
        else {
            console.log('addTable: tableRows has a bad row. A row should be an array of cells. You provided:');
            console.log(row);
        }
        arrRows.push(newRow);
    });
    // STEP 3: Set options
    opt.x = getSmartParseNumber(opt.x || (opt.x === 0 ? 0 : EMU / 2), 'X', presLayout);
    opt.y = getSmartParseNumber(opt.y || (opt.y === 0 ? 0 : EMU / 2), 'Y', presLayout);
    if (opt.h)
        opt.h = getSmartParseNumber(opt.h, 'Y', presLayout); // NOTE: Dont set default `h` - leaving it null triggers auto-rowH in `makeXMLSlide()`
    opt.autoPage = typeof opt.autoPage === 'boolean' ? opt.autoPage : false;
    opt.fontSize = opt.fontSize || DEF_FONT_SIZE;
    opt.autoPageLineWeight = typeof opt.autoPageLineWeight !== 'undefined' && !isNaN(Number(opt.autoPageLineWeight)) ? Number(opt.autoPageLineWeight) : 0;
    opt.margin = opt.margin === 0 || opt.margin ? opt.margin : DEF_CELL_MARGIN_PT;
    if (typeof opt.margin === 'number')
        opt.margin = [Number(opt.margin), Number(opt.margin), Number(opt.margin), Number(opt.margin)];
    if (opt.autoPageLineWeight > 1)
        opt.autoPageLineWeight = 1;
    else if (opt.autoPageLineWeight < -1)
        opt.autoPageLineWeight = -1;
    // Set default color if needed (table option > inherit from Slide > default to black)
    if (!opt.color)
        opt.color = opt.color || DEF_FONT_COLOR;
    // Set/Calc table width
    // Get slide margins - start with default values, then adjust if master or slide margins exist
    var arrTableMargin = DEF_SLIDE_MARGIN_IN;
    // Case 1: Master margins
    if (slideLayout && typeof slideLayout.margin !== 'undefined') {
        if (Array.isArray(slideLayout.margin))
            arrTableMargin = slideLayout.margin;
        else if (!isNaN(Number(slideLayout.margin)))
            arrTableMargin = [Number(slideLayout.margin), Number(slideLayout.margin), Number(slideLayout.margin), Number(slideLayout.margin)];
    }
    // Case 2: Table margins
    /* FIXME: add `margin` option to slide options
        else if ( addNewSlide.margin ) {
            if ( Array.isArray(addNewSlide.margin) ) arrTableMargin = addNewSlide.margin;
            else if ( !isNaN(Number(addNewSlide.margin)) ) arrTableMargin = [Number(addNewSlide.margin), Number(addNewSlide.margin), Number(addNewSlide.margin), Number(addNewSlide.margin)];
        }
    */
    // Calc table width depending upon what data we have - several scenarios exist (including bad data, eg: colW doesnt match col count)
    if (opt.w) {
        opt.w = getSmartParseNumber(opt.w, 'X', presLayout);
    }
    else if (opt.colW) {
        if (typeof opt.colW === 'string' || typeof opt.colW === 'number') {
            opt.w = Math.floor(Number(opt.colW) * arrRows[0].length);
        }
        else if (opt.colW && Array.isArray(opt.colW) && opt.colW.length !== arrRows[0].length) {
            console.warn('addTable: colW.length != data.length! Defaulting to evenly distributed col widths.');
            var numColWidth = Math.floor((presLayout.width / EMU - arrTableMargin[1] - arrTableMargin[3]) / arrRows[0].length);
            opt.colW = [];
            for (var idx = 0; idx < arrRows[0].length; idx++) {
                opt.colW.push(numColWidth);
            }
            opt.w = Math.floor(numColWidth * arrRows[0].length);
        }
    }
    else {
        opt.w = Math.floor(presLayout.width / EMU - arrTableMargin[1] - arrTableMargin[3]);
    }
    // STEP 4: Convert units to EMU now (we use different logic in makeSlide->table - smartCalc is not used)
    if (opt.x && opt.x < 20)
        opt.x = inch2Emu(opt.x);
    if (opt.y && opt.y < 20)
        opt.y = inch2Emu(opt.y);
    if (opt.w && opt.w < 20)
        opt.w = inch2Emu(opt.w);
    if (opt.h && opt.h < 20)
        opt.h = inch2Emu(opt.h);
    // STEP 5: Loop over cells: transform each to ITableCell; check to see whether to skip autopaging while here
    arrRows.forEach(function (row) {
        row.forEach(function (cell, idy) {
            // A: Transform cell data if needed
            /* Table rows can be an object or plain text - transform into object when needed
                // EX:
                var arrTabRows1 = [
                    [ { text:'A1\nA2', options:{rowspan:2, fill:'99FFCC'} } ]
                    ,[ 'B2', 'C2', 'D2', 'E2' ]
                ]
            */
            if (typeof cell === 'number' || typeof cell === 'string') {
                // Grab table formatting `opts` to use here so text style/format inherits as it should
                row[idy] = { type: SLIDE_OBJECT_TYPES.tablecell, text: row[idy].toString(), options: opt };
            }
            else if (typeof cell === 'object') {
                // ARG0: `text`
                if (typeof cell.text === 'number')
                    row[idy].text = row[idy].text.toString();
                else if (typeof cell.text === 'undefined' || cell.text === null)
                    row[idy].text = '';
                // ARG1: `options`: ensure options exists
                row[idy].options = cell.options || {};
                // Set type to tabelcell
                row[idy].type = SLIDE_OBJECT_TYPES.tablecell;
            }
            // B: Check for fine-grained formatting, disable auto-page when found
            // Since genXmlTextBody already checks for text array ( text:[{},..{}] ) we're done!
            // Text in individual cells will be formatted as they are added by calls to genXmlTextBody within table builder
            if (cell.text && Array.isArray(cell.text))
                opt.autoPage = false;
        });
    });
    // STEP 6: Auto-Paging: (via {options} and used internally)
    // (used internally by `tableToSlides()` to not engage recursion - we've already paged the table data, just add this one)
    if (opt && opt.autoPage === false) {
        // Create hyperlink rels (IMPORTANT: Wait until table has been shredded across Slides or all rels will end-up on Slide 1!)
        createHyperlinkRels(target, arrRows);
        // Add data (NOTE: Use `extend` to avoid mutation)
        target.data.push({
            type: SLIDE_OBJECT_TYPES.table,
            arrTabRows: arrRows,
            options: Object.assign({}, opt),
        });
    }
    else {
        // Loop over rows and create 1-N tables as needed (ISSUE#21)
        getSlidesForTableRows(arrRows, opt, presLayout, slideLayout).forEach(function (slide, idx) {
            // A: Create new Slide when needed, otherwise, use existing (NOTE: More than 1 table can be on a Slide, so we will go up AND down the Slide chain)
            if (!getSlide(target.number + idx))
                slides.push(addSlide(slideLayout ? slideLayout.name : null));
            // B: Reset opt.y to `option`/`margin` after first Slide (ISSUE#43, ISSUE#47, ISSUE#48)
            if (idx > 0)
                opt.y = inch2Emu(opt.newSlideStartY || arrTableMargin[0]);
            // C: Add this table to new Slide
            {
                var newSlide = getSlide(target.number + idx);
                opt.autoPage = false;
                // Create hyperlink rels (IMPORTANT: Wait until table has been shredded across Slides or all rels will end-up on Slide 1!)
                createHyperlinkRels(newSlide, slide.rows);
                // Add rows to new slide
                newSlide.addTable(slide.rows, Object.assign({}, opt));
            }
        });
    }
}
/**
 * Adds a text object to a slide definition.
 * @param {string|IText[]} text
 * @param {ITextOpts} opt
 * @param {ISlide} target - slide object that the text should be added to
 * @param {boolean} isPlaceholder` is this a placeholder object
 * @since: 1.0.0
 */
function addTextDefinition(target, text, opts, isPlaceholder) {
    var opt = opts || {};
    if (!opt.bodyProp)
        opt.bodyProp = {};
    var newObject = {
        text: (Array.isArray(text) && text.length === 0 ? '' : text || '') || '',
        type: isPlaceholder ? SLIDE_OBJECT_TYPES.placeholder : SLIDE_OBJECT_TYPES.text,
        options: opt,
        shape: opt.shape,
    };
    // STEP 1: Set some options
    {
        // A: Placeholders should inherit their colors or override them, so don't default them
        if (!opt.placeholder) {
            opt.color = opt.color || target.color || DEF_FONT_COLOR; // Set color (options > inherit from Slide > default to black)
        }
        // B
        if (opt.shape && opt.shape.name === 'line') {
            opt.line = opt.line || '333333';
            opt.lineSize = opt.lineSize || 1;
        }
        // C
        newObject.options.lineSpacing = opt.lineSpacing && !isNaN(opt.lineSpacing) ? opt.lineSpacing : null;
        // D: Transform text options to bodyProperties as thats how we build XML
        newObject.options.bodyProp.autoFit = opt.autoFit || false; // If true, shape will collapse to text size (Fit To shape)
        newObject.options.bodyProp.anchor = !opt.placeholder ? TEXT_VALIGN.ctr : null; // VALS: [t,ctr,b]
        newObject.options.bodyProp.vert = opt.vert || null; // VALS: [eaVert,horz,mongolianVert,vert,vert270,wordArtVert,wordArtVertRtl]
        if ((opt.inset && !isNaN(Number(opt.inset))) || opt.inset === 0) {
            newObject.options.bodyProp.lIns = inch2Emu(opt.inset);
            newObject.options.bodyProp.rIns = inch2Emu(opt.inset);
            newObject.options.bodyProp.tIns = inch2Emu(opt.inset);
            newObject.options.bodyProp.bIns = inch2Emu(opt.inset);
        }
    }
    // STEP 2: Transform `align`/`valign` to XML values, store in bodyProp for XML gen
    {
        if ((newObject.options.align || '').toLowerCase().startsWith('c'))
            newObject.options.bodyProp.align = TEXT_HALIGN.center;
        else if ((newObject.options.align || '').toLowerCase().startsWith('l'))
            newObject.options.bodyProp.align = TEXT_HALIGN.left;
        else if ((newObject.options.align || '').toLowerCase().startsWith('r'))
            newObject.options.bodyProp.align = TEXT_HALIGN.right;
        else if ((newObject.options.align || '').toLowerCase().startsWith('j'))
            newObject.options.bodyProp.align = TEXT_HALIGN.justify;
        if ((newObject.options.valign || '').toLowerCase().startsWith('b'))
            newObject.options.bodyProp.anchor = TEXT_VALIGN.b;
        else if ((newObject.options.valign || '').toLowerCase().startsWith('c'))
            newObject.options.bodyProp.anchor = TEXT_VALIGN.ctr;
        else if ((newObject.options.valign || '').toLowerCase().startsWith('t'))
            newObject.options.bodyProp.anchor = TEXT_VALIGN.t;
    }
    // STEP 3: ROBUST: Set rational values for some shadow props if needed
    correctShadowOptions(opt.shadow);
    // STEP 4: Create hyperlinks
    createHyperlinkRels(target, newObject.text || '');
    // LAST: Add object to Slide
    target.data.push(newObject);
}
/**
 * Adds placeholder objects to slide
 * @param {ISlide} slide - slide object containing layouts
 */
function addPlaceholdersToSlideLayouts(slide) {
    (slide.slideLayout.data || []).forEach(function (slideLayoutObj) {
        if (slideLayoutObj.type === SLIDE_OBJECT_TYPES.placeholder) {
            // A: Search for this placeholder on Slide before we add
            // NOTE: Check to ensure a placeholder does not already exist on the Slide
            // They are created when they have been populated with text (ex: `slide.addText('Hi', { placeholder:'title' });`)
            if (slide.data.filter(function (slideObj) {
                return slideObj.options && slideObj.options.placeholder === slideLayoutObj.options.placeholder;
            }).length === 0) {
                addTextDefinition(slide, '', { placeholder: slideLayoutObj.options.placeholder }, false);
            }
        }
    });
}
/* -------------------------------------------------------------------------------- */
/**
 * Adds a background image or color to a slide definition.
 * @param {String|Object} bkg - color string or an object with image definition
 * @param {ISlide} target - slide object that the background is set to
 */
function addBackgroundDefinition(bkg, target) {
    if (typeof bkg === 'object' && (bkg.src || bkg.path || bkg.data)) {
        // Allow the use of only the data key (`path` isnt reqd)
        bkg.src = bkg.src || bkg.path || null;
        if (!bkg.src)
            bkg.src = 'preencoded.png';
        var strImgExtn = (bkg.src.split('.').pop() || 'png').split('?')[0]; // Handle "blah.jpg?width=540" etc.
        if (strImgExtn === 'jpg')
            strImgExtn = 'jpeg'; // base64-encoded jpg's come out as "data:image/jpeg;base64,/9j/[...]", so correct exttnesion to avoid content warnings at PPT startup
        var intRels = target.relsMedia.length + 1;
        target.relsMedia.push({
            path: bkg.src,
            type: SLIDE_OBJECT_TYPES.image,
            extn: strImgExtn,
            data: bkg.data || null,
            rId: intRels,
            Target: '../media/image' + (target.relsMedia.length + 1) + '.' + strImgExtn,
        });
        target.bkgdImgRid = intRels;
    }
    else if (bkg && typeof bkg === 'string') {
        target.bkgd = bkg;
    }
}
/**
 * Parses text/text-objects from `addText()` and `addTable()` methods; creates 'hyperlink'-type Slide Rels for each hyperlink found
 * @param {ISlide} target - slide object that any hyperlinks will be be added to
 * @param {number | string | IText | IText[] | ITableCell[][]} text - text to parse
 */
function createHyperlinkRels(target, text) {
    var textObjs = [];
    // Only text objects can have hyperlinks, bail when text param is plain text
    if (typeof text === 'string' || typeof text === 'number')
        return;
    // IMPORTANT: "else if" Array.isArray must come before typeof===object! Otherwise, code will exhaust recursion!
    else if (Array.isArray(text))
        textObjs = text;
    else if (typeof text === 'object')
        textObjs = [text];
    textObjs.forEach(function (text) {
        // `text` can be an array of other `text` objects (table cell word-level formatting), continue parsing using recursion
        if (Array.isArray(text))
            createHyperlinkRels(target, text);
        else if (text && typeof text === 'object' && text.options && text.options.hyperlink && !text.options.hyperlink.rId) {
            if (typeof text.options.hyperlink !== 'object')
                console.log("ERROR: text `hyperlink` option should be an object. Ex: `hyperlink: {url:'https://github.com'}` ");
            else if (!text.options.hyperlink.url && !text.options.hyperlink.slide)
                console.log("ERROR: 'hyperlink requires either: `url` or `slide`'");
            else {
                var relId = target.rels.length + target.relsChart.length + target.relsMedia.length + 1;
                target.rels.push({
                    type: SLIDE_OBJECT_TYPES.hyperlink,
                    data: text.options.hyperlink.slide ? 'slide' : 'dummy',
                    rId: relId,
                    Target: encodeXmlEntities(text.options.hyperlink.url) || text.options.hyperlink.slide.toString(),
                });
                text.options.hyperlink.rId = relId;
            }
        }
    });
}

/**
 * PptxGenJS Slide Class
 */
var Slide = /** @class */ (function () {
    function Slide(params) {
        this.addSlide = params.addSlide;
        this.getSlide = params.getSlide;
        this.presLayout = params.presLayout;
        this._setSlideNum = params.setSlideNum;
        this.name = 'Slide ' + params.slideNumber;
        this.number = params.slideNumber;
        this.data = [];
        this.rels = [];
        this.relsChart = [];
        this.relsMedia = [];
        this.slideLayout = params.slideLayout || null;
        // NOTE: Slide Numbers: In order for Slide Numbers to function they need to be in all 3 files: master/layout/slide
        // `defineSlideMaster` and `addNewSlide.slideNumber` will add {slideNumber} to `this.masterSlide` and `this.slideLayouts`
        // so, lastly, add to the Slide now.
        this.slideNumberObj = this.slideLayout && this.slideLayout.slideNumberObj ? this.slideLayout.slideNumberObj : null;
    }
    Object.defineProperty(Slide.prototype, "bkgd", {
        get: function () {
            return this._bkgd;
        },
        // TODO: add comments (also add to index.d.ts)
        set: function (value) {
            this._bkgd = value;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Slide.prototype, "color", {
        get: function () {
            return this._color;
        },
        // TODO: add comments (also add to index.d.ts)
        set: function (value) {
            this._color = value;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Slide.prototype, "hidden", {
        get: function () {
            return this._hidden;
        },
        // TODO: add comments (also add to index.d.ts)
        set: function (value) {
            this._hidden = value;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Slide.prototype, "slideNumber", {
        get: function () {
            return this._slideNumber;
        },
        // TODO: add comments (also add to index.d.ts)
        set: function (value) {
            // NOTE: Slide Numbers: In order for Slide Numbers to function they need to be in all 3 files: master/layout/slide
            this.slideNumberObj = value;
            this._slideNumber = value;
            this._setSlideNum(value);
        },
        enumerable: true,
        configurable: true
    });
    /**
     * Generate the chart based on input data.
     * @see OOXML Chart Spec: ISO/IEC 29500-1:2016(E)
     * @param {CHART_TYPE_NAMES|IChartMulti[]} `type` - chart type
     * @param {object[]} data - a JSON object with follow the following format
     * @param {IChartOpts} options - chart options
     * @example
     * {
     *   title: 'eSurvey chart',
     *   data: [
     *		{
     *			name: 'Income',
     *			labels: ['2005', '2006', '2007', '2008', '2009'],
     *			values: [23.5, 26.2, 30.1, 29.5, 24.6]
     *		},
     *		{
     *			name: 'Expense',
     *			labels: ['2005', '2006', '2007', '2008', '2009'],
     *			values: [18.1, 22.8, 23.9, 25.1, 25]
     *		}
     *	 ]
     * }
     * @return {Slide} this class
     */
    Slide.prototype.addChart = function (type, data, options) {
        addChartDefinition(this, type, data, options);
        return this;
    };
    /**
     * Add Image object
     * @note: Remote images (eg: "http://whatev.com/blah"/from web and/or remote server arent supported yet - we'd need to create an <img>, load it, then send to canvas
     * @see: https://stackoverflow.com/questions/164181/how-to-fetch-a-remote-image-to-display-in-a-canvas)
     * @param {IImageOpts} options - image options
     * @return {Slide} this class
     */
    Slide.prototype.addImage = function (options) {
        addImageDefinition(this, options);
        return this;
    };
    /**
     * Add Media (audio/video) object
     * @param {IMediaOpts} options - media options
     * @return {Slide} this class
     */
    Slide.prototype.addMedia = function (options) {
        addMediaDefinition(this, options);
        return this;
    };
    /**
     * Add Speaker Notes to Slide
     * @docs https://gitbrent.github.io/PptxGenJS/docs/speaker-notes.html
     * @param {string} notes - notes to add to slide
     * @return {Slide} this class
     */
    Slide.prototype.addNotes = function (notes) {
        addNotesDefinition(this, notes);
        return this;
    };
    /**
     * Add shape object to Slide
     * @param {IShape} shape - shape object
     * @param {IShapeOptions} options - shape options
     * @return {Slide} this class
     */
    Slide.prototype.addShape = function (shape, options) {
        addShapeDefinition(this, shape, options);
        return this;
    };
    /**
     * Add shape object to Slide
     * @note can be recursive
     * @param {TableRow[]} arrTabRows - table rows
     * @param {ITableOptions} options - table options
     * @return {Slide} this class
     */
    Slide.prototype.addTable = function (arrTabRows, options) {
        // FIXME: TODO-3: we pass `this` - we dont need to pass layouts - they can be read from this!
        addTableDefinition(this, arrTabRows, options, this.slideLayout, this.presLayout, this.addSlide, this.getSlide);
        return this;
    };
    /**
     * Add text object to Slide
     * @param {string|IText[]} text - text string or complex object
     * @param {ITextOpts} options - text options
     * @return {Slide} this class
     * @since: 1.0.0
     */
    Slide.prototype.addText = function (text, options) {
        addTextDefinition(this, text, options, false);
        return this;
    };
    return Slide;
}());

/**
 * PptxGenJS: Chart Generation
 */
/**
 * Based on passed data, creates Excel Worksheet that is used as a data source for a chart.
 * @param {ISlideRelChart} chartObject - chart object
 * @param {JSZip} zip - file that the resulting XLSX should be added to
 * @return {Promise} promise of generating the XLSX file
 */
function createExcelWorksheet(chartObject, zip) {
    var data = chartObject.data;
    return new Promise(function (resolve, reject) {
        var zipExcel = new JSZip();
        var intBubbleCols = (data.length - 1) * 2 + 1; // 1 for "X-Values", then 2 for every Y-Axis
        // A: Add folders
        zipExcel.folder('_rels');
        zipExcel.folder('docProps');
        zipExcel.folder('xl/_rels');
        zipExcel.folder('xl/tables');
        zipExcel.folder('xl/theme');
        zipExcel.folder('xl/worksheets');
        zipExcel.folder('xl/worksheets/_rels');
        // B: Add core contents
        {
            zipExcel.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
                '  <Default Extension="xml" ContentType="application/xml"/>' +
                '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
                //+ '  <Default Extension="jpeg" ContentType="image/jpg"/><Default Extension="png" ContentType="image/png"/>'
                //+ '  <Default Extension="bmp" ContentType="image/bmp"/><Default Extension="gif" ContentType="image/gif"/><Default Extension="tif" ContentType="image/tif"/><Default Extension="pdf" ContentType="application/pdf"/><Default Extension="mov" ContentType="application/movie"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>'
                //+ '  <Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>'
                '  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
                '  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
                '  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' +
                '  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
                '  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' +
                '  <Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>' +
                '  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
                '  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
                '</Types>\n');
            zipExcel.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
                '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' +
                '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
                '</Relationships>\n');
            zipExcel.file('docProps/app.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
                '<Application>Microsoft Excel</Application>' +
                '<DocSecurity>0</DocSecurity>' +
                '<ScaleCrop>false</ScaleCrop>' +
                '<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>Sheet1</vt:lpstr></vt:vector></TitlesOfParts>' +
                '</Properties>\n');
            zipExcel.file('docProps/core.xml', '<?xml version="1.0" encoding="UTF-8"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
                '<dc:creator>PptxGenJS</dc:creator>' +
                '<cp:lastModifiedBy>Ely, Brent</cp:lastModifiedBy>' +
                '<dcterms:created xsi:type="dcterms:W3CDTF">' +
                new Date().toISOString() +
                '</dcterms:created>' +
                '<dcterms:modified xsi:type="dcterms:W3CDTF">' +
                new Date().toISOString() +
                '</dcterms:modified>' +
                '</cp:coreProperties>\n');
            zipExcel.file('xl/_rels/workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
                '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>' +
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
                '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>' +
                '</Relationships>\n');
            zipExcel.file('xl/styles.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="0" formatCode="General"/></numFmts><fonts count="4"><font><sz val="9"/><color indexed="8"/><name val="Geneva"/></font><font><sz val="9"/><color indexed="8"/><name val="Geneva"/></font><font><sz val="10"/><color indexed="8"/><name val="Geneva"/></font><font><sz val="18"/><color indexed="8"/>' +
                '<name val="Arial"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><dxfs count="0"/><tableStyles count="0"/><colors><indexedColors><rgbColor rgb="ff000000"/><rgbColor rgb="ffffffff"/><rgbColor rgb="ffff0000"/><rgbColor rgb="ff00ff00"/><rgbColor rgb="ff0000ff"/>' +
                '<rgbColor rgb="ffffff00"/><rgbColor rgb="ffff00ff"/><rgbColor rgb="ff00ffff"/><rgbColor rgb="ff000000"/><rgbColor rgb="ffffffff"/><rgbColor rgb="ff878787"/><rgbColor rgb="fff9f9f9"/></indexedColors></colors></styleSheet>\n');
            zipExcel.file('xl/theme/theme1.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light" panose="020F0302020204030204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Yu Gothic Light"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="DengXian Light"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:majorFont><a:minorFont><a:latin typeface="Calibri" panose="020F0502020204030204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Yu Gothic"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="DengXian"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/><a:extLst><a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}"><thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" id="{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}" vid="{4A3C46E8-61CC-4603-A589-7422A47A8E4A}"/></a:ext></a:extLst></a:theme>');
            zipExcel.file('xl/workbook.xml', '<?xml version="1.0" encoding="UTF-8"?>' +
                '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">' +
                '<fileVersion appName="xl" lastEdited="6" lowestEdited="6" rupBuild="14420"/>' +
                '<workbookPr />' +
                '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="15960" windowHeight="18080"/></bookViews>' +
                '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1" /></sheets>' +
                '<calcPr calcId="171026" concurrentCalc="0"/>' +
                '</workbook>\n');
            zipExcel.file('xl/worksheets/_rels/sheet1.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>' +
                '</Relationships>\n');
        }
        // sharedStrings.xml
        {
            // A: Start XML
            var strSharedStrings_1 = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
            if (chartObject.opts.type === CHART_TYPES.BUBBLE) {
                strSharedStrings_1 +=
                    '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' + (intBubbleCols + 1) + '" uniqueCount="' + (intBubbleCols + 1) + '">';
            }
            else if (chartObject.opts.type === CHART_TYPES.SCATTER) {
                strSharedStrings_1 +=
                    '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' + (data.length + 1) + '" uniqueCount="' + (data.length + 1) + '">';
            }
            else {
                strSharedStrings_1 +=
                    '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' +
                        (data[0].labels.length + data.length + 1) +
                        '" uniqueCount="' +
                        (data[0].labels.length + data.length + 1) +
                        '">';
                // B: Add 'blank' for A1
                strSharedStrings_1 += '<si><t xml:space="preserve"></t></si>';
            }
            // C: Add `name`/Series
            if (chartObject.opts.type === CHART_TYPES.BUBBLE) {
                data.forEach(function (objData, idx) {
                    if (idx === 0)
                        strSharedStrings_1 += '<si><t>' + 'X-Axis' + '</t></si>';
                    else {
                        strSharedStrings_1 += '<si><t>' + encodeXmlEntities(objData.name || ' ') + '</t></si>';
                        strSharedStrings_1 += '<si><t>' + encodeXmlEntities('Size ' + idx) + '</t></si>';
                    }
                });
            }
            else {
                data.forEach(function (objData) {
                    strSharedStrings_1 += '<si><t>' + encodeXmlEntities((objData.name || ' ').replace('X-Axis', 'X-Values')) + '</t></si>';
                });
            }
            // D: Add `labels`/Categories
            if (chartObject.opts.type !== CHART_TYPES.BUBBLE && chartObject.opts.type !== CHART_TYPES.SCATTER) {
                data[0].labels.forEach(function (label) {
                    strSharedStrings_1 += '<si><t>' + encodeXmlEntities(label) + '</t></si>';
                });
            }
            strSharedStrings_1 += '</sst>\n';
            zipExcel.file('xl/sharedStrings.xml', strSharedStrings_1);
        }
        // tables/table1.xml
        {
            var strTableXml_1 = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
            if (chartObject.opts.type === CHART_TYPES.BUBBLE) ;
            else if (chartObject.opts.type === CHART_TYPES.SCATTER) {
                strTableXml_1 +=
                    '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:' +
                        LETTERS[data.length - 1] +
                        (data[0].values.length + 1) +
                        '" totalsRowShown="0">';
                strTableXml_1 += '<tableColumns count="' + data.length + '">';
                data.forEach(function (_obj, idx) {
                    strTableXml_1 += '<tableColumn id="' + (idx + 1) + '" name="' + (idx === 0 ? 'X-Values' : 'Y-Value ' + idx) + '" />';
                });
            }
            else {
                strTableXml_1 +=
                    '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:' +
                        LETTERS[data.length] +
                        (data[0].labels.length + 1) +
                        '" totalsRowShown="0">';
                strTableXml_1 += '<tableColumns count="' + (data.length + 1) + '">';
                strTableXml_1 += '<tableColumn id="1" name=" " />';
                data.forEach(function (obj, idx) {
                    strTableXml_1 += '<tableColumn id="' + (idx + 2) + '" name="' + encodeXmlEntities(obj.name) + '" />';
                });
            }
            strTableXml_1 += '</tableColumns>';
            strTableXml_1 += '<tableStyleInfo showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0" />';
            strTableXml_1 += '</table>';
            zipExcel.file('xl/tables/table1.xml', strTableXml_1);
        }
        // worksheets/sheet1.xml
        {
            var strSheetXml_1 = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
            strSheetXml_1 +=
                '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">';
            if (chartObject.opts.type === CHART_TYPES.BUBBLE) {
                strSheetXml_1 += '<dimension ref="A1:' + LETTERS[intBubbleCols - 1] + (data[0].values.length + 1) + '" />';
            }
            else if (chartObject.opts.type === CHART_TYPES.SCATTER) {
                strSheetXml_1 += '<dimension ref="A1:' + LETTERS[data.length - 1] + (data[0].values.length + 1) + '" />';
            }
            else {
                strSheetXml_1 += '<dimension ref="A1:' + LETTERS[data.length] + (data[0].labels.length + 1) + '" />';
            }
            strSheetXml_1 += '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="B1" sqref="B1" /></sheetView></sheetViews>';
            strSheetXml_1 += '<sheetFormatPr baseColWidth="10" defaultColWidth="11.5" defaultRowHeight="12" />';
            if (chartObject.opts.type === CHART_TYPES.BUBBLE) {
                strSheetXml_1 += '<cols>';
                strSheetXml_1 += '<col min="1" max="' + data.length + '" width="11" customWidth="1" />';
                strSheetXml_1 += '</cols>';
                /* EX: INPUT: `data`
                [
                    { name:'X-Axis'  , values:[10,11,12,13,14,15,16,17,18,19,20] },
                    { name:'Y-Axis 1', values:[ 1, 6, 7, 8, 9], sizes:[ 4, 5, 6, 7, 8] },
                    { name:'Y-Axis 2', values:[33,32,42,53,63], sizes:[11,12,13,14,15] }
                ];
                */
                /* EX: OUTPUT: bubbleChart Worksheet:
                    -|----A-----|------B-----|------C-----|------D-----|------E-----|
                    1| X-Values | Y-Values 1 | Y-Sizes 1  | Y-Values 2 | Y-Sizes 2  |
                    2|    11    |     22     |      4     |     33     |      8     |
                    -|----------|------------|------------|------------|------------|
                */
                strSheetXml_1 += '<sheetData>';
                // A: Create header row first (NOTE: Start at index=1 as headers cols start with 'B')
                strSheetXml_1 += '<row r="1" spans="1:' + intBubbleCols + '">';
                strSheetXml_1 += '<c r="A1" t="s"><v>0</v></c>';
                for (var idx = 1; idx < intBubbleCols; idx++) {
                    strSheetXml_1 += '<c r="' + (idx < 26 ? LETTERS[idx] : 'A' + LETTERS[idx % LETTERS.length]) + '1" t="s">'; // NOTE: use `t="s"` for label cols!
                    strSheetXml_1 += '<v>' + idx + '</v>';
                    strSheetXml_1 += '</c>';
                }
                strSheetXml_1 += '</row>';
                // B: Add row for each X-Axis value (Y-Axis* value is optional)
                data[0].values.forEach(function (val, idx) {
                    // Leading col is reserved for the 'X-Axis' value, so hard-code it, then loop over col values
                    strSheetXml_1 += '<row r="' + (idx + 2) + '" spans="1:' + intBubbleCols + '">';
                    strSheetXml_1 += '<c r="A' + (idx + 2) + '"><v>' + val + '</v></c>';
                    // Add Y-Axis 1->N (idy=0 = Xaxis)
                    var idxColLtr = 1;
                    for (var idy = 1; idy < data.length; idy++) {
                        // y-value
                        strSheetXml_1 += '<c r="' + (idxColLtr < 26 ? LETTERS[idxColLtr] : 'A' + LETTERS[idxColLtr % LETTERS.length]) + '' + (idx + 2) + '">';
                        strSheetXml_1 += '<v>' + (data[idy].values[idx] || '') + '</v>';
                        strSheetXml_1 += '</c>';
                        idxColLtr++;
                        // y-size
                        strSheetXml_1 += '<c r="' + (idxColLtr < 26 ? LETTERS[idxColLtr] : 'A' + LETTERS[idxColLtr % LETTERS.length]) + '' + (idx + 2) + '">';
                        strSheetXml_1 += '<v>' + (data[idy].sizes[idx] || '') + '</v>';
                        strSheetXml_1 += '</c>';
                        idxColLtr++;
                    }
                    strSheetXml_1 += '</row>';
                });
            }
            else if (chartObject.opts.type === CHART_TYPES.SCATTER) {
                strSheetXml_1 += '<cols>';
                strSheetXml_1 += '<col min="1" max="' + data.length + '" width="11" customWidth="1" />';
                //data.forEach((obj,idx)=>{ strSheetXml += '<col min="'+(idx+1)+'" max="'+(idx+1)+'" width="11" customWidth="1" />' });
                strSheetXml_1 += '</cols>';
                /* EX: INPUT: `data`
                [
                    { name:'X-Axis'  , values:[10,11,12,13,14,15,16,17,18,19,20] },
                    { name:'Y-Axis 1', values:[ 1, 6, 7, 8, 9] },
                    { name:'Y-Axis 2', values:[33,32,42,53,63] }
                ];
                */
                /* EX: OUTPUT: scatterChart Worksheet:
                    -|----A-----|------B-----|
                    1| X-Values | Y-Values 1 |
                    2|    11    |     22     |
                    -|----------|------------|
                */
                strSheetXml_1 += '<sheetData>';
                // A: Create header row first (NOTE: Start at index=1 as headers cols start with 'B')
                strSheetXml_1 += '<row r="1" spans="1:' + data.length + '">';
                strSheetXml_1 += '<c r="A1" t="s"><v>0</v></c>';
                for (var idx = 1; idx < data.length; idx++) {
                    strSheetXml_1 += '<c r="' + (idx < 26 ? LETTERS[idx] : 'A' + LETTERS[idx % LETTERS.length]) + '1" t="s">'; // NOTE: use `t="s"` for label cols!
                    strSheetXml_1 += '<v>' + idx + '</v>';
                    strSheetXml_1 += '</c>';
                }
                strSheetXml_1 += '</row>';
                // B: Add row for each X-Axis value (Y-Axis* value is optional)
                data[0].values.forEach(function (val, idx) {
                    // Leading col is reserved for the 'X-Axis' value, so hard-code it, then loop over col values
                    strSheetXml_1 += '<row r="' + (idx + 2) + '" spans="1:' + data.length + '">';
                    strSheetXml_1 += '<c r="A' + (idx + 2) + '"><v>' + val + '</v></c>';
                    // Add Y-Axis 1->N
                    for (var idy = 1; idy < data.length; idy++) {
                        strSheetXml_1 += '<c r="' + (idy < 26 ? LETTERS[idy] : 'A' + LETTERS[idy % LETTERS.length]) + '' + (idx + 2) + '">';
                        strSheetXml_1 += '<v>' + (data[idy].values[idx] || data[idy].values[idx] === 0 ? data[idy].values[idx] : '') + '</v>';
                        strSheetXml_1 += '</c>';
                    }
                    strSheetXml_1 += '</row>';
                });
            }
            else {
                strSheetXml_1 += '<cols>';
                strSheetXml_1 += '<col min="1" max="1" width="11" customWidth="1" />';
                //data.forEach(function(){ strSheetXml += '<col min="10" max="100" width="10" customWidth="1" />' });
                strSheetXml_1 += '</cols>';
                strSheetXml_1 += '<sheetData>';
                /* EX: INPUT: `data`
                [
                    { name:'Red', labels:['Jan..May-17'], values:[11,13,14,15,16] },
                    { name:'Amb', labels:['Jan..May-17'], values:[22, 6, 7, 8, 9] },
                    { name:'Grn', labels:['Jan..May-17'], values:[33,32,42,53,63] }
                ];
                */
                /* EX: OUTPUT: lineChart Worksheet:
                    -|---A---|--B--|--C--|--D--|
                    1|       | Red | Amb | Grn |
                    2|Jan-17 |   11|   22|   33|
                    3|Feb-17 |   55|   43|   70|
                    4|Mar-17 |   56|  143|   99|
                    5|Apr-17 |   65|    3|  120|
                    6|May-17 |   75|   93|  170|
                    -|-------|-----|-----|-----|
                */
                // A: Create header row first (NOTE: Start at index=1 as headers cols start with 'B')
                strSheetXml_1 += '<row r="1" spans="1:' + (data.length + 1) + '">';
                strSheetXml_1 += '<c r="A1" t="s"><v>0</v></c>';
                for (var idx = 1; idx <= data.length; idx++) {
                    // FIXME: Max cols is 52
                    strSheetXml_1 += '<c r="' + (idx < 26 ? LETTERS[idx] : 'A' + LETTERS[idx % LETTERS.length]) + '1" t="s">'; // NOTE: use `t="s"` for label cols!
                    strSheetXml_1 += '<v>' + idx + '</v>';
                    strSheetXml_1 += '</c>';
                }
                strSheetXml_1 += '</row>';
                // B: Add data row(s) for each category
                data[0].labels.forEach(function (_cat, idx) {
                    // Leading col is reserved for the label, so hard-code it, then loop over col values
                    strSheetXml_1 += '<row r="' + (idx + 2) + '" spans="1:' + (data.length + 1) + '">';
                    strSheetXml_1 += '<c r="A' + (idx + 2) + '" t="s">';
                    strSheetXml_1 += '<v>' + (data.length + idx + 1) + '</v>';
                    strSheetXml_1 += '</c>';
                    for (var idy = 0; idy < data.length; idy++) {
                        strSheetXml_1 += '<c r="' + (idy + 1 < 26 ? LETTERS[idy + 1] : 'A' + LETTERS[(idy + 1) % LETTERS.length]) + '' + (idx + 2) + '">';
                        strSheetXml_1 += '<v>' + (data[idy].values[idx] || '') + '</v>';
                        strSheetXml_1 += '</c>';
                    }
                    strSheetXml_1 += '</row>';
                });
            }
            strSheetXml_1 += '</sheetData>';
            strSheetXml_1 += '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3" />';
            // Link the `table1.xml` file to define an actual Table in Excel
            // NOTE: This only works with scatter charts - all others give a "cannot find linked file" error
            // ....: Since we dont need the table anyway (chart data can be edited/range selected, etc.), just dont use this
            // ....: Leaving this so nobody foolishly attempts to add this in the future
            // strSheetXml += '<tableParts count="1"><tablePart r:id="rId1" /></tableParts>';
            strSheetXml_1 += '</worksheet>\n';
            zipExcel.file('xl/worksheets/sheet1.xml', strSheetXml_1);
        }
        // C: Add XLSX to PPTX export
        zipExcel
            .generateAsync({ type: 'base64' })
            .then(function (content) {
            // 1: Create the embedded Excel worksheet with labels and data
            zip.file('ppt/embeddings/Microsoft_Excel_Worksheet' + chartObject.globalId + '.xlsx', content, { base64: true });
            // 2: Create the chart.xml and rels files
            zip.file('ppt/charts/_rels/' + chartObject.fileName + '.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/Microsoft_Excel_Worksheet' +
                chartObject.globalId +
                '.xlsx"/>' +
                '</Relationships>');
            zip.file('ppt/charts/' + chartObject.fileName, makeXmlCharts(chartObject));
            // 3: Done
            resolve();
        })
            .catch(function (strErr) {
            reject(strErr);
        });
    });
}
/**
 * Main entry point method for create charts
 * @see: http://www.datypic.com/sc/ooxml/s-dml-chart.xsd.html
 * @param {ISlideRelChart} rel - chart object
 * @return {string} XML
 */
function makeXmlCharts(rel) {
    var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
    var usesSecondaryValAxis = false;
    // STEP 1: Create chart
    {
        // CHARTSPACE: BEGIN vvv
        strXml +=
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
        strXml += '<c:date1904 val="0"/>'; // ppt defaults to 1904 dates, excel to 1900
        strXml += '<c:chart>';
        // OPTION: Title
        if (rel.opts.showTitle) {
            strXml += genXmlTitle({
                title: rel.opts.title || 'Chart Title',
                fontSize: rel.opts.titleFontSize || DEF_FONT_TITLE_SIZE,
                color: rel.opts.titleColor,
                fontFace: rel.opts.titleFontFace,
                rotate: rel.opts.titleRotate,
                titleAlign: rel.opts.titleAlign,
                titlePos: rel.opts.titlePos,
            });
            strXml += '<c:autoTitleDeleted val="0"/>';
        }
        else {
            // NOTE: Add autoTitleDeleted tag in else to prevent default creation of chart title even when showTitle is set to false
            strXml += '<c:autoTitleDeleted val="1"/>';
        }
        // Add 3D view tag
        if (rel.opts.type === CHART_TYPES.BAR3D) {
            strXml += '<c:view3D>';
            strXml += ' <c:rotX val="' + rel.opts.v3DRotX + '"/>';
            strXml += ' <c:rotY val="' + rel.opts.v3DRotY + '"/>';
            strXml += ' <c:rAngAx val="' + (rel.opts.v3DRAngAx === false ? 0 : 1) + '"/>';
            strXml += ' <c:perspective val="' + rel.opts.v3DPerspective + '"/>';
            strXml += '</c:view3D>';
        }
        strXml += '<c:plotArea>';
        // IMPORTANT: Dont specify layout to enable auto-fit: PPT does a great job maximizing space with all 4 TRBL locations
        if (rel.opts.layout) {
            strXml += '<c:layout>';
            strXml += ' <c:manualLayout>';
            strXml += '  <c:layoutTarget val="inner" />';
            strXml += '  <c:xMode val="edge" />';
            strXml += '  <c:yMode val="edge" />';
            strXml += '  <c:x val="' + (rel.opts.layout.x || 0) + '" />';
            strXml += '  <c:y val="' + (rel.opts.layout.y || 0) + '" />';
            strXml += '  <c:w val="' + (rel.opts.layout.w || 1) + '" />';
            strXml += '  <c:h val="' + (rel.opts.layout.h || 1) + '" />';
            strXml += ' </c:manualLayout>';
            strXml += '</c:layout>';
        }
        else {
            strXml += '<c:layout/>';
        }
    }
    // A: Create Chart XML -----------------------------------------------------------
    if (Array.isArray(rel.opts.type)) {
        rel.opts.type.forEach(function (type) {
            // TODO: FIXME: theres `options` on chart rels??
            var options = getMix(rel.opts, type.options);
            //let options: IChartOpts = { type: type.type, }
            var valAxisId = options['secondaryValAxis'] ? AXIS_ID_VALUE_SECONDARY : AXIS_ID_VALUE_PRIMARY;
            var catAxisId = options['secondaryCatAxis'] ? AXIS_ID_CATEGORY_SECONDARY : AXIS_ID_CATEGORY_PRIMARY;
            usesSecondaryValAxis = usesSecondaryValAxis || options['secondaryValAxis'];
            strXml += makeChartType(type.type, type.data, options, valAxisId, catAxisId, true);
        });
    }
    else {
        strXml += makeChartType(rel.opts.type, rel.data, rel.opts, AXIS_ID_VALUE_PRIMARY, AXIS_ID_CATEGORY_PRIMARY, false);
    }
    // B: Axes -----------------------------------------------------------
    if (rel.opts.type !== CHART_TYPES.PIE && rel.opts.type !== CHART_TYPES.DOUGHNUT) {
        // Param check
        if (rel.opts.valAxes && !usesSecondaryValAxis) {
            throw new Error('Secondary axis must be used by one of the multiple charts');
        }
        if (rel.opts.catAxes) {
            if (!rel.opts.valAxes || rel.opts.valAxes.length !== rel.opts.catAxes.length) {
                throw new Error('There must be the same number of value and category axes.');
            }
            strXml += makeCatAxis(getMix(rel.opts, rel.opts.catAxes[0]), AXIS_ID_CATEGORY_PRIMARY, AXIS_ID_VALUE_PRIMARY);
            if (rel.opts.catAxes[1]) {
                strXml += makeCatAxis(getMix(rel.opts, rel.opts.catAxes[1]), AXIS_ID_CATEGORY_SECONDARY, AXIS_ID_VALUE_PRIMARY);
            }
        }
        else {
            strXml += makeCatAxis(rel.opts, AXIS_ID_CATEGORY_PRIMARY, AXIS_ID_VALUE_PRIMARY);
        }
        if (rel.opts.valAxes) {
            strXml += makeValAxis(getMix(rel.opts, rel.opts.valAxes[0]), AXIS_ID_VALUE_PRIMARY);
            if (rel.opts.valAxes[1]) {
                strXml += makeValAxis(getMix(rel.opts, rel.opts.valAxes[1]), AXIS_ID_VALUE_SECONDARY);
            }
        }
        else {
            strXml += makeValAxis(rel.opts, AXIS_ID_VALUE_PRIMARY);
            // Add series axis for 3D bar
            if (rel.opts.type === CHART_TYPES.BAR3D) {
                strXml += makeSerAxis(rel.opts, AXIS_ID_SERIES_PRIMARY, AXIS_ID_VALUE_PRIMARY);
            }
        }
    }
    // C: Chart Properties and plotArea Options: Border, Data Table, Fill, Legend
    {
        // NOTE: DataTable goes between '</c:valAx>' and '<c:spPr>'
        if (rel.opts.showDataTable) {
            strXml += '<c:dTable>';
            strXml += '  <c:showHorzBorder val="' + (rel.opts.showDataTableHorzBorder === false ? 0 : 1) + '"/>';
            strXml += '  <c:showVertBorder val="' + (rel.opts.showDataTableVertBorder === false ? 0 : 1) + '"/>';
            strXml += '  <c:showOutline    val="' + (rel.opts.showDataTableOutline === false ? 0 : 1) + '"/>';
            strXml += '  <c:showKeys       val="' + (rel.opts.showDataTableKeys === false ? 0 : 1) + '"/>';
            strXml += '  <c:spPr>';
            strXml += '    <a:noFill/>';
            strXml +=
                '    <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln>';
            strXml += '    <a:effectLst/>';
            strXml += '  </c:spPr>';
            strXml +=
                '  <c:txPr>\
						  <a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>\
						  <a:lstStyle/>\
						  <a:p>\
							<a:pPr rtl="0">\
							  <a:defRPr sz="1197" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0">\
								<a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill>\
								<a:latin typeface="+mn-lt"/>\
								<a:ea typeface="+mn-ea"/>\
								<a:cs typeface="+mn-cs"/>\
							  </a:defRPr>\
							</a:pPr>\
							<a:endParaRPr lang="en-US"/>\
						  </a:p>\
						</c:txPr>\
					  </c:dTable>';
        }
        strXml += '  <c:spPr>';
        // OPTION: Fill
        strXml += rel.opts.fill ? genXmlColorSelection(rel.opts.fill) : '<a:noFill/>';
        // OPTION: Border
        strXml += rel.opts.border
            ? '<a:ln w="' + rel.opts.border.pt * ONEPT + '"' + ' cap="flat">' + genXmlColorSelection(rel.opts.border.color) + '</a:ln>'
            : '<a:ln><a:noFill/></a:ln>';
        // Close shapeProp/plotArea before Legend
        strXml += '    <a:effectLst/>';
        strXml += '  </c:spPr>';
        strXml += '</c:plotArea>';
        // OPTION: Legend
        // IMPORTANT: Dont specify layout to enable auto-fit: PPT does a great job maximizing space with all 4 TRBL locations
        if (rel.opts.showLegend) {
            strXml += '<c:legend>';
            strXml += '<c:legendPos val="' + rel.opts.legendPos + '"/>';
            strXml += '<c:layout/>';
            strXml += '<c:overlay val="0"/>';
            if (rel.opts.legendFontFace || rel.opts.legendFontSize || rel.opts.legendColor) {
                strXml += '<c:txPr>';
                strXml += '  <a:bodyPr/>';
                strXml += '  <a:lstStyle/>';
                strXml += '  <a:p>';
                strXml += '    <a:pPr>';
                strXml += rel.opts.legendFontSize ? '<a:defRPr sz="' + Number(rel.opts.legendFontSize) * 100 + '">' : '<a:defRPr>';
                if (rel.opts.legendColor)
                    strXml += genXmlColorSelection(rel.opts.legendColor);
                if (rel.opts.legendFontFace)
                    strXml += '<a:latin typeface="' + rel.opts.legendFontFace + '"/>';
                if (rel.opts.legendFontFace)
                    strXml += '<a:cs    typeface="' + rel.opts.legendFontFace + '"/>';
                strXml += '      </a:defRPr>';
                strXml += '    </a:pPr>';
                strXml += '    <a:endParaRPr lang="en-US"/>';
                strXml += '  </a:p>';
                strXml += '</c:txPr>';
            }
            strXml += '</c:legend>';
        }
    }
    strXml += '  <c:plotVisOnly val="1"/>';
    strXml += '  <c:dispBlanksAs val="' + rel.opts.displayBlanksAs + '"/>';
    if (rel.opts.type === CHART_TYPES.SCATTER)
        strXml += '<c:showDLblsOverMax val="1"/>';
    strXml += '</c:chart>';
    // D: CHARTSPACE SHAPE PROPS
    strXml += '<c:spPr>';
    strXml += '  <a:noFill/>';
    strXml += '  <a:ln w="12700" cap="flat"><a:noFill/><a:miter lim="400000"/></a:ln>';
    strXml += '  <a:effectLst/>';
    strXml += '</c:spPr>';
    // E: DATA (Add relID)
    strXml += '<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData>';
    // LAST: chartSpace end
    strXml += '</c:chartSpace>';
    return strXml;
}
/**
 * Create XML string for any given chart type
 * @example: <c:bubbleChart> or <c:lineChart>
 *
 * @param {CHART_TYPE_NAMES} `chartType` chart type name
 * @param {OptsChartData[]} `data` chart data
 * @param {IChartOpts} `opts` chart options
 * @param {string} `valAxisId`
 * @param {string} `catAxisId`
 * @param {boolean} `isMultiTypeChart`
 * @return {string} XML
 */
function makeChartType(chartType, data, opts, valAxisId, catAxisId, isMultiTypeChart) {
    // NOTE: "Chart Range" (as shown in "select Chart Area dialog") is calculated.
    // ....: Ensure each X/Y Axis/Col has same row height (esp. applicable to XY Scatter where X can often be larger than Y's)
    var strXml = '';
    switch (chartType) {
        case CHART_TYPES.AREA:
        case CHART_TYPES.BAR:
        case CHART_TYPES.BAR3D:
        case CHART_TYPES.LINE:
        case CHART_TYPES.RADAR:
            // 1: Start Chart
            strXml += '<c:' + chartType + 'Chart>';
            if (chartType === CHART_TYPES.BAR || chartType === CHART_TYPES.BAR3D) {
                strXml += '<c:barDir val="' + opts.barDir + '"/>';
                strXml += '<c:grouping val="' + opts.barGrouping + '"/>';
            }
            if (chartType === CHART_TYPES.RADAR) {
                strXml += '<c:radarStyle val="' + opts.radarStyle + '"/>';
            }
            strXml += '<c:varyColors val="0"/>';
            // 2: "Series" block for every data row
            /* EX:
                data: [
                 {
                   name: 'Region 1',
                   labels: ['April', 'May', 'June', 'July'],
                   values: [17, 26, 53, 96]
                 },
                 {
                   name: 'Region 2',
                   labels: ['April', 'May', 'June', 'July'],
                   values: [55, 43, 70, 58]
                 }
                ]
            */
            var colorIndex_1 = -1; // Maintain the color index by region
            data.forEach(function (obj) {
                colorIndex_1++;
                var idx = obj.index;
                strXml += '<c:ser>';
                strXml += '  <c:idx val="' + idx + '"/>';
                strXml += '  <c:order val="' + idx + '"/>';
                strXml += '  <c:tx>';
                strXml += '    <c:strRef>';
                strXml += '      <c:f>Sheet1!$' + getExcelColName(idx + 1) + '$1</c:f>';
                strXml += '      <c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>' + encodeXmlEntities(obj.name) + '</c:v></c:pt></c:strCache>';
                strXml += '    </c:strRef>';
                strXml += '  </c:tx>';
                strXml += '  <c:invertIfNegative val="0"/>';
                // Fill and Border
                var strSerColor = opts.chartColors ? opts.chartColors[colorIndex_1 % opts.chartColors.length] : null;
                strXml += '  <c:spPr>';
                if (strSerColor === 'transparent') {
                    strXml += '<a:noFill/>';
                }
                else if (opts.chartColorsOpacity) {
                    strXml += '<a:solidFill>' + createColorElement(strSerColor, '<a:alpha val="' + opts.chartColorsOpacity + '000"/>') + '</a:solidFill>';
                }
                else {
                    strXml += '<a:solidFill>' + createColorElement(strSerColor) + '</a:solidFill>';
                }
                if (chartType === CHART_TYPES.LINE) {
                    if (opts.lineSize === 0) {
                        strXml += '<a:ln><a:noFill/></a:ln>';
                    }
                    else {
                        strXml += '<a:ln w="' + opts.lineSize * ONEPT + '" cap="flat"><a:solidFill>' + createColorElement(strSerColor) + '</a:solidFill>';
                        strXml += '<a:prstDash val="' + (opts.lineDash || 'solid') + '"/><a:round/></a:ln>';
                    }
                }
                else if (opts.dataBorder) {
                    strXml +=
                        '<a:ln w="' +
                            opts.dataBorder.pt * ONEPT +
                            '" cap="flat"><a:solidFill>' +
                            createColorElement(opts.dataBorder.color) +
                            '</a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>';
                }
                strXml += createShadowElement(opts.shadow, DEF_SHAPE_SHADOW);
                strXml += '  </c:spPr>';
                // Data Labels per series
                // [20190117] NOTE: Adding these to RADAR chart causes unrecoverable corruption!
                if (chartType !== CHART_TYPES.RADAR) {
                    strXml += '  <c:dLbls>';
                    strXml += '    <c:numFmt formatCode="' + opts.dataLabelFormatCode + '" sourceLinked="0"/>';
                    if (opts.dataLabelBkgrdColors) {
                        strXml += '    <c:spPr>';
                        strXml += '       <a:solidFill>' + createColorElement(strSerColor) + '</a:solidFill>';
                        strXml += '    </c:spPr>';
                    }
                    strXml += '    <c:txPr>';
                    strXml += '      <a:bodyPr/>';
                    strXml += '      <a:lstStyle/>';
                    strXml += '      <a:p><a:pPr>';
                    strXml += '        <a:defRPr b="0" i="0" strike="noStrike" sz="' + (opts.dataLabelFontSize || DEF_FONT_SIZE) + '00" u="none">';
                    strXml += '          <a:solidFill>' + createColorElement(opts.dataLabelColor || DEF_FONT_COLOR) + '</a:solidFill>';
                    strXml += '          <a:latin typeface="' + (opts.dataLabelFontFace || 'Arial') + '"/>';
                    strXml += '        </a:defRPr>';
                    strXml += '      </a:pPr></a:p>';
                    strXml += '    </c:txPr>';
                    // Setting dLblPos tag for bar3D seems to break the generated chart
                    if (chartType !== CHART_TYPES.AREA && chartType !== CHART_TYPES.BAR3D) {
                        strXml += '<c:dLblPos val="' + (opts.dataLabelPosition || 'outEnd') + '"/>';
                    }
                    strXml += '    <c:showLegendKey val="0"/>';
                    strXml += '    <c:showVal val="' + (opts.showValue ? '1' : '0') + '"/>';
                    strXml += '    <c:showCatName val="0"/>';
                    strXml += '    <c:showSerName val="0"/>';
                    strXml += '    <c:showPercent val="0"/>';
                    strXml += '    <c:showBubbleSize val="0"/>';
                    strXml += '    <c:showLeaderLines val="0"/>';
                    strXml += '  </c:dLbls>';
                }
                // 'c:marker' tag: `lineDataSymbol`
                if (chartType === CHART_TYPES.LINE || chartType === CHART_TYPES.RADAR) {
                    strXml += '<c:marker>';
                    strXml += '  <c:symbol val="' + opts.lineDataSymbol + '"/>';
                    if (opts.lineDataSymbolSize) {
                        // Defaults to "auto" otherwise (but this is usually too small, so there is a default)
                        strXml += '  <c:size val="' + opts.lineDataSymbolSize + '"/>';
                    }
                    strXml += '  <c:spPr>';
                    strXml +=
                        '    <a:solidFill>' +
                            createColorElement(opts.chartColors[idx + 1 > opts.chartColors.length ? Math.floor(Math.random() * opts.chartColors.length) : idx]) +
                            '</a:solidFill>';
                    var symbolLineColor = opts.lineDataSymbolLineColor || strSerColor;
                    strXml +=
                        '    <a:ln w="' +
                            opts.lineDataSymbolLineSize +
                            '" cap="flat"><a:solidFill>' +
                            createColorElement(symbolLineColor) +
                            '</a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>';
                    strXml += '    <a:effectLst/>';
                    strXml += '  </c:spPr>';
                    strXml += '</c:marker>';
                }
                // Color chart bars various colors
                // Allow users with a single data set to pass their own array of colors (check for this using != ours)
                if ((chartType === CHART_TYPES.BAR || chartType === CHART_TYPES.BAR3D) && (data.length === 1 || opts.valueBarColors) && opts.chartColors !== BARCHART_COLORS) {
                    // Series Data Point colors
                    obj.values.forEach(function (value, index) {
                        var arrColors = value < 0 ? opts.invertedColors || BARCHART_COLORS : opts.chartColors || [];
                        strXml += '  <c:dPt>';
                        strXml += '    <c:idx val="' + index + '"/>';
                        strXml += '      <c:invertIfNegative val="' + (opts.invertedColors ? 0 : 1) + '"/>';
                        strXml += '    <c:bubble3D val="0"/>';
                        strXml += '    <c:spPr>';
                        if (opts.lineSize === 0) {
                            strXml += '<a:ln><a:noFill/></a:ln>';
                        }
                        else if (chartType === CHART_TYPES.BAR) {
                            strXml += '<a:solidFill>';
                            strXml += '  <a:srgbClr val="' + arrColors[index % arrColors.length] + '"/>';
                            strXml += '</a:solidFill>';
                        }
                        else {
                            strXml += '<a:ln>';
                            strXml += '  <a:solidFill>';
                            strXml += '   <a:srgbClr val="' + arrColors[index % arrColors.length] + '"/>';
                            strXml += '  </a:solidFill>';
                            strXml += '</a:ln>';
                        }
                        strXml += createShadowElement(opts.shadow, DEF_SHAPE_SHADOW);
                        strXml += '    </c:spPr>';
                        strXml += '  </c:dPt>';
                    });
                }
                // 2: "Categories"
                {
                    strXml += '<c:cat>';
                    if (opts.catLabelFormatCode) {
                        // Use 'numRef' as catLabelFormatCode implies that we are expecting numbers here
                        strXml += '  <c:numRef>';
                        strXml += '    <c:f>Sheet1!' + '$A$2:$A$' + (obj.labels.length + 1) + '</c:f>';
                        strXml += '    <c:numCache>';
                        strXml += '      <c:formatCode>' + opts.catLabelFormatCode + '</c:formatCode>';
                        strXml += '      <c:ptCount val="' + obj.labels.length + '"/>';
                        obj.labels.forEach(function (label, idx) {
                            strXml += '<c:pt idx="' + idx + '"><c:v>' + encodeXmlEntities(label) + '</c:v></c:pt>';
                        });
                        strXml += '    </c:numCache>';
                        strXml += '  </c:numRef>';
                    }
                    else {
                        strXml += '  <c:strRef>';
                        strXml += '    <c:f>Sheet1!' + '$A$2:$A$' + (obj.labels.length + 1) + '</c:f>';
                        strXml += '    <c:strCache>';
                        strXml += '	     <c:ptCount val="' + obj.labels.length + '"/>';
                        obj.labels.forEach(function (label, idx) {
                            strXml += '<c:pt idx="' + idx + '"><c:v>' + encodeXmlEntities(label) + '</c:v></c:pt>';
                        });
                        strXml += '    </c:strCache>';
                        strXml += '  </c:strRef>';
                    }
                    strXml += '</c:cat>';
                }
                // 3: "Values"
                {
                    strXml += '  <c:val>';
                    strXml += '    <c:numRef>';
                    strXml += '      <c:f>Sheet1!' + '$' + getExcelColName(idx + 1) + '$2:$' + getExcelColName(idx + 1) + '$' + (obj.labels.length + 1) + '</c:f>';
                    strXml += '      <c:numCache>';
                    strXml += '        <c:formatCode>General</c:formatCode>';
                    strXml += '	       <c:ptCount val="' + obj.labels.length + '"/>';
                    obj.values.forEach(function (value, idx) {
                        strXml += '<c:pt idx="' + idx + '"><c:v>' + (value || value === 0 ? value : '') + '</c:v></c:pt>';
                    });
                    strXml += '      </c:numCache>';
                    strXml += '    </c:numRef>';
                    strXml += '  </c:val>';
                }
                // Option: `smooth`
                if (chartType === CHART_TYPES.LINE)
                    strXml += '<c:smooth val="' + (opts.lineSmooth ? '1' : '0') + '"/>';
                // 4: Close "SERIES"
                strXml += '</c:ser>';
            });
            // 3: "Data Labels"
            {
                strXml += '  <c:dLbls>';
                strXml += '    <c:numFmt formatCode="' + opts.dataLabelFormatCode + '" sourceLinked="0"/>';
                strXml += '    <c:txPr>';
                strXml += '      <a:bodyPr/>';
                strXml += '      <a:lstStyle/>';
                strXml += '      <a:p><a:pPr>';
                strXml +=
                    '        <a:defRPr b="' + (opts.dataLabelFontBold ? 1 : 0) + '" i="0" strike="noStrike" sz="' + (opts.dataLabelFontSize || DEF_FONT_SIZE) + '00" u="none">';
                strXml += '          <a:solidFill>' + createColorElement(opts.dataLabelColor || DEF_FONT_COLOR) + '</a:solidFill>';
                strXml += '          <a:latin typeface="' + (opts.dataLabelFontFace || 'Arial') + '"/>';
                strXml += '        </a:defRPr>';
                strXml += '      </a:pPr></a:p>';
                strXml += '    </c:txPr>';
                // NOTE: Throwing an error while creating a multi type chart which contains area chart as the below line appears for the other chart type.
                // Either the given change can be made or the below line can be removed to stop the slide containing multi type chart with area to crash.
                if (opts.type !== CHART_TYPES.AREA && opts.type !== CHART_TYPES.RADAR && !isMultiTypeChart)
                    strXml += '<c:dLblPos val="' + (opts.dataLabelPosition || 'outEnd') + '"/>';
                strXml += '    <c:showLegendKey val="0"/>';
                strXml += '    <c:showVal val="' + (opts.showValue ? '1' : '0') + '"/>';
                strXml += '    <c:showCatName val="0"/>';
                strXml += '    <c:showSerName val="0"/>';
                strXml += '    <c:showPercent val="0"/>';
                strXml += '    <c:showBubbleSize val="0"/>';
                strXml += '    <c:showLeaderLines val="0"/>';
                strXml += '  </c:dLbls>';
            }
            // 4: Add more chart options (gapWidth, line Marker, etc.)
            if (chartType === CHART_TYPES.BAR) {
                strXml += '  <c:gapWidth val="' + opts.barGapWidthPct + '"/>';
                strXml += '  <c:overlap val="' + ((opts.barGrouping || '').indexOf('tacked') > -1 ? 100 : 0) + '"/>';
            }
            else if (chartType === CHART_TYPES.BAR3D) {
                strXml += '  <c:gapWidth val="' + opts.barGapWidthPct + '"/>';
                strXml += '  <c:gapDepth val="' + opts.barGapDepthPct + '"/>';
                strXml += '  <c:shape val="' + opts.bar3DShape + '"/>';
            }
            else if (chartType === CHART_TYPES.LINE) {
                strXml += '  <c:marker val="1"/>';
            }
            // 5: Add axisId (NOTE: order matters! (category comes first))
            strXml += '  <c:axId val="' + catAxisId + '"/>';
            strXml += '  <c:axId val="' + valAxisId + '"/>';
            strXml += '  <c:axId val="' + AXIS_ID_SERIES_PRIMARY + '"/>';
            // 6: Close Chart tag
            strXml += '</c:' + chartType + 'Chart>';
            // end switch
            break;
        case CHART_TYPES.SCATTER:
            /*
                `data` = [
                    { name:'X-Axis',    values:[1,2,3,4,5,6,7,8,9,10,11,12] },
                    { name:'Y-Value 1', values:[13, 20, 21, 25] },
                    { name:'Y-Value 2', values:[ 1,  2,  5,  9] }
                ];
            */
            // 1: Start Chart
            strXml += '<c:' + chartType + 'Chart>';
            strXml += '<c:scatterStyle val="lineMarker"/>';
            strXml += '<c:varyColors val="0"/>';
            // 2: Series: (One for each Y-Axis)
            colorIndex_1 = -1;
            data.filter(function (_obj, idx) {
                return idx > 0;
            }).forEach(function (obj, idx) {
                colorIndex_1++;
                strXml += '<c:ser>';
                strXml += '  <c:idx val="' + idx + '"/>';
                strXml += '  <c:order val="' + idx + '"/>';
                strXml += '  <c:tx>';
                strXml += '    <c:strRef>';
                strXml += '      <c:f>Sheet1!$' + LETTERS[idx + 1] + '$1</c:f>';
                strXml += '      <c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>' + obj.name + '</c:v></c:pt></c:strCache>';
                strXml += '    </c:strRef>';
                strXml += '  </c:tx>';
                // 'c:spPr': Fill, Border, Line, LineStyle (dash, etc.), Shadow
                strXml += '  <c:spPr>';
                {
                    var strSerColor = opts.chartColors[colorIndex_1 % opts.chartColors.length];
                    if (strSerColor === 'transparent') {
                        strXml += '<a:noFill/>';
                    }
                    else if (opts.chartColorsOpacity) {
                        strXml += '<a:solidFill>' + createColorElement(strSerColor, '<a:alpha val="' + opts.chartColorsOpacity + '000"/>') + '</a:solidFill>';
                    }
                    else {
                        strXml += '<a:solidFill>' + createColorElement(strSerColor) + '</a:solidFill>';
                    }
                    if (opts.lineSize === 0) {
                        strXml += '<a:ln><a:noFill/></a:ln>';
                    }
                    else {
                        strXml += '<a:ln w="' + opts.lineSize * ONEPT + '" cap="flat"><a:solidFill>' + createColorElement(strSerColor) + '</a:solidFill>';
                        strXml += '<a:prstDash val="' + (opts.lineDash || 'solid') + '"/><a:round/></a:ln>';
                    }
                    // Shadow
                    strXml += createShadowElement(opts.shadow, DEF_SHAPE_SHADOW);
                }
                strXml += '  </c:spPr>';
                // 'c:marker' tag: `lineDataSymbol`
                {
                    var strSerColor = opts.chartColors[colorIndex_1 % opts.chartColors.length];
                    strXml += '<c:marker>';
                    strXml += '  <c:symbol val="' + opts.lineDataSymbol + '"/>';
                    if (opts.lineDataSymbolSize) {
                        // Defaults to "auto" otherwise (but this is usually too small, so there is a default)
                        strXml += '  <c:size val="' + opts.lineDataSymbolSize + '"/>';
                    }
                    strXml += '  <c:spPr>';
                    strXml +=
                        '    <a:solidFill>' +
                            createColorElement(opts.chartColors[idx + 1 > opts.chartColors.length ? Math.floor(Math.random() * opts.chartColors.length) : idx]) +
                            '</a:solidFill>';
                    var symbolLineColor = opts.lineDataSymbolLineColor || strSerColor;
                    strXml +=
                        '    <a:ln w="' +
                            opts.lineDataSymbolLineSize +
                            '" cap="flat"><a:solidFill>' +
                            createColorElement(symbolLineColor) +
                            '</a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>';
                    strXml += '    <a:effectLst/>';
                    strXml += '  </c:spPr>';
                    strXml += '</c:marker>';
                }
                // Option: scatter data point labels
                if (opts.showLabel) {
                    var chartUuid_1 = getUuid('-xxxx-xxxx-xxxx-xxxxxxxxxxxx');
                    if (obj.labels && (opts.dataLabelFormatScatter === 'custom' || opts.dataLabelFormatScatter === 'customXY')) {
                        strXml += '<c:dLbls>';
                        obj.labels.forEach(function (label, idx) {
                            if (opts.dataLabelFormatScatter === 'custom' || opts.dataLabelFormatScatter === 'customXY') {
                                strXml += '  <c:dLbl>';
                                strXml += '    <c:idx val="' + idx + '"/>';
                                strXml += '    <c:tx>';
                                strXml += '      <c:rich>';
                                strXml += '			<a:bodyPr>';
                                strXml += '				<a:spAutoFit/>';
                                strXml += '			</a:bodyPr>';
                                strXml += '        	<a:lstStyle/>';
                                strXml += '        	<a:p>';
                                strXml += '				<a:pPr>';
                                strXml += '					<a:defRPr/>';
                                strXml += '				</a:pPr>';
                                strXml += '          	<a:r>';
                                strXml += '            		<a:rPr lang="' + (opts.lang || 'en-US') + '" dirty="0"/>';
                                strXml += '            		<a:t>' + encodeXmlEntities(label) + '</a:t>';
                                strXml += '          	</a:r>';
                                // Apply XY values at end of custom label
                                // Do not apply the values if the label was empty or just spaces
                                // This allows for selective labelling where required
                                if (opts.dataLabelFormatScatter === 'customXY' && !/^ *$/.test(label)) {
                                    strXml += '          	<a:r>';
                                    strXml += '          		<a:rPr lang="' + (opts.lang || 'en-US') + '" baseline="0" dirty="0"/>';
                                    strXml += '          		<a:t> (</a:t>';
                                    strXml += '          	</a:r>';
                                    strXml += '          	<a:fld id="{' + getUuid('xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx') + '}" type="XVALUE">';
                                    strXml += '          		<a:rPr lang="' + (opts.lang || 'en-US') + '" baseline="0"/>';
                                    strXml += '          		<a:pPr>';
                                    strXml += '          			<a:defRPr/>';
                                    strXml += '          		</a:pPr>';
                                    strXml += '          		<a:t>[' + encodeXmlEntities(obj.name) + '</a:t>';
                                    strXml += '          	</a:fld>';
                                    strXml += '          	<a:r>';
                                    strXml += '          		<a:rPr lang="' + (opts.lang || 'en-US') + '" baseline="0" dirty="0"/>';
                                    strXml += '          		<a:t>, </a:t>';
                                    strXml += '          	</a:r>';
                                    strXml += '          	<a:fld id="{' + getUuid('xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx') + '}" type="YVALUE">';
                                    strXml += '          		<a:rPr lang="' + (opts.lang || 'en-US') + '" baseline="0"/>';
                                    strXml += '          		<a:pPr>';
                                    strXml += '          			<a:defRPr/>';
                                    strXml += '          		</a:pPr>';
                                    strXml += '          		<a:t>[' + encodeXmlEntities(obj.name) + ']</a:t>';
                                    strXml += '          	</a:fld>';
                                    strXml += '          	<a:r>';
                                    strXml += '          		<a:rPr lang="' + (opts.lang || 'en-US') + '" baseline="0" dirty="0"/>';
                                    strXml += '          		<a:t>)</a:t>';
                                    strXml += '          	</a:r>';
                                    strXml += '          	<a:endParaRPr lang="' + (opts.lang || 'en-US') + '" dirty="0"/>';
                                }
                                strXml += '        	</a:p>';
                                strXml += '      </c:rich>';
                                strXml += '    </c:tx>';
                                strXml += '    <c:spPr>';
                                strXml += '    	<a:noFill/>';
                                strXml += '    	<a:ln>';
                                strXml += '    		<a:noFill/>';
                                strXml += '    	</a:ln>';
                                strXml += '    	<a:effectLst/>';
                                strXml += '    </c:spPr>';
                                strXml += '    <c:showLegendKey val="0"/>';
                                strXml += '    <c:showVal val="0"/>';
                                strXml += '    <c:showCatName val="0"/>';
                                strXml += '    <c:showSerName val="0"/>';
                                strXml += '    <c:showPercent val="0"/>';
                                strXml += '    <c:showBubbleSize val="0"/>';
                                strXml += '	  <c:showLeaderLines val="1"/>';
                                strXml += '    <c:extLst>';
                                strXml += '      <c:ext uri="{CE6537A1-D6FC-4f65-9D91-7224C49458BB}" xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart">';
                                strXml += '			<c15:dlblFieldTable/>';
                                strXml += '			<c15:showDataLabelsRange val="0"/>';
                                strXml += '		</c:ext>';
                                strXml += '      <c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}" xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart">';
                                strXml += '			<c16:uniqueId val="{' + '00000000'.substring(0, 8 - (idx + 1).toString().length).toString() + (idx + 1) + chartUuid_1 + '}"/>';
                                strXml += '      </c:ext>';
                                strXml += '		</c:extLst>';
                                strXml += '</c:dLbl>';
                            }
                        });
                        strXml += '</c:dLbls>';
                    }
                    if (opts.dataLabelFormatScatter === 'XY') {
                        strXml += '<c:dLbls>';
                        strXml += '	<c:spPr>';
                        strXml += '		<a:noFill/>';
                        strXml += '		<a:ln>';
                        strXml += '			<a:noFill/>';
                        strXml += '		</a:ln>';
                        strXml += '	  	<a:effectLst/>';
                        strXml += '	</c:spPr>';
                        strXml += '	<c:txPr>';
                        strXml += '		<a:bodyPr>';
                        strXml += '			<a:spAutoFit/>';
                        strXml += '		</a:bodyPr>';
                        strXml += '		<a:lstStyle/>';
                        strXml += '		<a:p>';
                        strXml += '	    	<a:pPr>';
                        strXml += '        		<a:defRPr/>';
                        strXml += '	    	</a:pPr>';
                        strXml += '	    	<a:endParaRPr lang="en-US"/>';
                        strXml += '		</a:p>';
                        strXml += '	</c:txPr>';
                        strXml += '	<c:showLegendKey val="0"/>';
                        strXml += '	<c:showVal val="' + opts.showLabel ? '1' : '0' + '"/>';
                        strXml += '	<c:showCatName val="' + opts.showLabel ? '1' : '0' + '"/>';
                        strXml += '	<c:showSerName val="0"/>';
                        strXml += '	<c:showPercent val="0"/>';
                        strXml += '	<c:showBubbleSize val="0"/>';
                        strXml += '	<c:extLst>';
                        strXml += '		<c:ext uri="{CE6537A1-D6FC-4f65-9D91-7224C49458BB}" xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart">';
                        strXml += '			<c15:showLeaderLines val="1"/>';
                        strXml += '		</c:ext>';
                        strXml += '	</c:extLst>';
                        strXml += '</c:dLbls>';
                    }
                }
                // Color bar chart bars various colors
                // Allow users with a single data set to pass their own array of colors (check for this using != ours)
                if ((data.length === 1 || opts.valueBarColors) && opts.chartColors !== BARCHART_COLORS) {
                    // Series Data Point colors
                    obj.values.forEach(function (value, index) {
                        var arrColors = value < 0 ? opts.invertedColors || BARCHART_COLORS : opts.chartColors || [];
                        strXml += '  <c:dPt>';
                        strXml += '    <c:idx val="' + index + '"/>';
                        strXml += '      <c:invertIfNegative val="' + (opts.invertedColors ? 0 : 1) + '"/>';
                        strXml += '    <c:bubble3D val="0"/>';
                        strXml += '    <c:spPr>';
                        if (opts.lineSize === 0) {
                            strXml += '<a:ln><a:noFill/></a:ln>';
                        }
                        else {
                            strXml += '<a:solidFill>';
                            strXml += ' <a:srgbClr val="' + arrColors[index % arrColors.length] + '"/>';
                            strXml += '</a:solidFill>';
                        }
                        strXml += createShadowElement(opts.shadow, DEF_SHAPE_SHADOW);
                        strXml += '    </c:spPr>';
                        strXml += '  </c:dPt>';
                    });
                }
                // 3: "Values": Scatter Chart has 2: `xVal` and `yVal`
                {
                    // X-Axis is always the same
                    strXml += '<c:xVal>';
                    strXml += '  <c:numRef>';
                    strXml += '    <c:f>Sheet1!$A$2:$A$' + (data[0].values.length + 1) + '</c:f>';
                    strXml += '    <c:numCache>';
                    strXml += '      <c:formatCode>General</c:formatCode>';
                    strXml += '      <c:ptCount val="' + data[0].values.length + '"/>';
                    data[0].values.forEach(function (value, idx) {
                        strXml += '<c:pt idx="' + idx + '"><c:v>' + (value || value === 0 ? value : '') + '</c:v></c:pt>';
                    });
                    strXml += '    </c:numCache>';
                    strXml += '  </c:numRef>';
                    strXml += '</c:xVal>';
                    // Y-Axis vals are this object's `values`
                    strXml += '<c:yVal>';
                    strXml += '  <c:numRef>';
                    strXml += '    <c:f>Sheet1!$' + getExcelColName(idx + 1) + '$2:$' + getExcelColName(idx + 1) + '$' + (data[0].values.length + 1) + '</c:f>';
                    strXml += '    <c:numCache>';
                    strXml += '      <c:formatCode>General</c:formatCode>';
                    // NOTE: Use pt count and iterate over data[0] (X-Axis) as user can have more values than data (eg: timeline where only first few months are populated)
                    strXml += '      <c:ptCount val="' + data[0].values.length + '"/>';
                    data[0].values.forEach(function (_value, idx) {
                        strXml += '<c:pt idx="' + idx + '"><c:v>' + (obj.values[idx] || obj.values[idx] === 0 ? obj.values[idx] : '') + '</c:v></c:pt>';
                    });
                    strXml += '    </c:numCache>';
                    strXml += '  </c:numRef>';
                    strXml += '</c:yVal>';
                }
                // Option: `smooth`
                strXml += '<c:smooth val="' + (opts.lineSmooth ? '1' : '0') + '"/>';
                // 4: Close "SERIES"
                strXml += '</c:ser>';
            });
            // 3: Data Labels
            {
                strXml += '  <c:dLbls>';
                strXml += '    <c:numFmt formatCode="' + opts.dataLabelFormatCode + '" sourceLinked="0"/>';
                strXml += '    <c:txPr>';
                strXml += '      <a:bodyPr/>';
                strXml += '      <a:lstStyle/>';
                strXml += '      <a:p><a:pPr>';
                strXml += '        <a:defRPr b="0" i="0" strike="noStrike" sz="' + (opts.dataLabelFontSize || DEF_FONT_SIZE) + '00" u="none">';
                strXml += '          <a:solidFill>' + createColorElement(opts.dataLabelColor || DEF_FONT_COLOR) + '</a:solidFill>';
                strXml += '          <a:latin typeface="' + (opts.dataLabelFontFace || 'Arial') + '"/>';
                strXml += '        </a:defRPr>';
                strXml += '      </a:pPr></a:p>';
                strXml += '    </c:txPr>';
                strXml += '    <c:dLblPos val="' + (opts.dataLabelPosition || 'outEnd') + '"/>';
                strXml += '    <c:showLegendKey val="0"/>';
                strXml += '    <c:showVal val="' + (opts.showValue ? '1' : '0') + '"/>';
                strXml += '    <c:showCatName val="0"/>';
                strXml += '    <c:showSerName val="0"/>';
                strXml += '    <c:showPercent val="0"/>';
                strXml += '    <c:showBubbleSize val="0"/>';
                strXml += '  </c:dLbls>';
            }
            // 4: Add axisId (NOTE: order matters! (category comes first))
            strXml += '  <c:axId val="' + catAxisId + '"/>';
            strXml += '  <c:axId val="' + valAxisId + '"/>';
            // 5: Close Chart tag
            strXml += '</c:' + chartType + 'Chart>';
            // end switch
            break;
        case CHART_TYPES.BUBBLE:
            /*
                `data` = [
                    { name:'X-Axis',     values:[1,2,3,4,5,6,7,8,9,10,11,12] },
                    { name:'Y-Values 1', values:[13, 20, 21, 25], sizes:[10, 5, 20, 15] },
                    { name:'Y-Values 2', values:[ 1,  2,  5,  9], sizes:[ 5, 3,  9,  3] }
                ];
            */
            // 1: Start Chart
            strXml += '<c:' + chartType + 'Chart>';
            strXml += '<c:varyColors val="0"/>';
            // 2: Series: (One for each Y-Axis)
            colorIndex_1 = -1;
            var idxColLtr_1 = 1;
            data.filter(function (_obj, idx) {
                return idx > 0;
            }).forEach(function (obj, idx) {
                colorIndex_1++;
                strXml += '<c:ser>';
                strXml += '  <c:idx val="' + idx + '"/>';
                strXml += '  <c:order val="' + idx + '"/>';
                // A: `<c:tx>`
                strXml += '  <c:tx>';
                strXml += '    <c:strRef>';
                strXml += '      <c:f>Sheet1!$' + LETTERS[idxColLtr_1] + '$1</c:f>';
                strXml += '      <c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>' + obj.name + '</c:v></c:pt></c:strCache>';
                strXml += '    </c:strRef>';
                strXml += '  </c:tx>';
                // B: '<c:spPr>': Fill, Border, Line, LineStyle (dash, etc.), Shadow
                {
                    strXml += '<c:spPr>';
                    var strSerColor = opts.chartColors[colorIndex_1 % opts.chartColors.length];
                    if (strSerColor === 'transparent') {
                        strXml += '<a:noFill/>';
                    }
                    else if (opts.chartColorsOpacity) {
                        strXml += '<a:solidFill>' + createColorElement(strSerColor, '<a:alpha val="' + opts.chartColorsOpacity + '000"/>') + '</a:solidFill>';
                    }
                    else {
                        strXml += '<a:solidFill>' + createColorElement(strSerColor) + '</a:solidFill>';
                    }
                    if (opts.lineSize === 0) {
                        strXml += '<a:ln><a:noFill/></a:ln>';
                    }
                    else if (opts.dataBorder) {
                        strXml +=
                            '<a:ln w="' +
                                opts.dataBorder.pt * ONEPT +
                                '" cap="flat"><a:solidFill>' +
                                createColorElement(opts.dataBorder.color) +
                                '</a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>';
                    }
                    else {
                        strXml += '<a:ln w="' + opts.lineSize * ONEPT + '" cap="flat"><a:solidFill>' + createColorElement(strSerColor) + '</a:solidFill>';
                        strXml += '<a:prstDash val="' + (opts.lineDash || 'solid') + '"/><a:round/></a:ln>';
                    }
                    // Shadow
                    strXml += createShadowElement(opts.shadow, DEF_SHAPE_SHADOW);
                    strXml += '</c:spPr>';
                }
                // C: '<c:dLbls>' "Data Labels"
                // Let it be defaulted for now
                // D: '<c:xVal>'/'<c:yVal>' "Values": Scatter Chart has 2: `xVal` and `yVal`
                {
                    // X-Axis is always the same
                    strXml += '<c:xVal>';
                    strXml += '  <c:numRef>';
                    strXml += '    <c:f>Sheet1!$A$2:$A$' + (data[0].values.length + 1) + '</c:f>';
                    strXml += '    <c:numCache>';
                    strXml += '      <c:formatCode>General</c:formatCode>';
                    strXml += '      <c:ptCount val="' + data[0].values.length + '"/>';
                    data[0].values.forEach(function (value, idx) {
                        strXml += '<c:pt idx="' + idx + '"><c:v>' + (value || value === 0 ? value : '') + '</c:v></c:pt>';
                    });
                    strXml += '    </c:numCache>';
                    strXml += '  </c:numRef>';
                    strXml += '</c:xVal>';
                    // Y-Axis vals are this object's `values`
                    strXml += '<c:yVal>';
                    strXml += '  <c:numRef>';
                    strXml += '    <c:f>Sheet1!$' + getExcelColName(idxColLtr_1) + '$2:$' + getExcelColName(idxColLtr_1) + '$' + (data[0].values.length + 1) + '</c:f>';
                    idxColLtr_1++;
                    strXml += '    <c:numCache>';
                    strXml += '      <c:formatCode>General</c:formatCode>';
                    // NOTE: Use pt count and iterate over data[0] (X-Axis) as user can have more values than data (eg: timeline where only first few months are populated)
                    strXml += '      <c:ptCount val="' + data[0].values.length + '"/>';
                    data[0].values.forEach(function (_value, idx) {
                        strXml += '<c:pt idx="' + idx + '"><c:v>' + (obj.values[idx] || obj.values[idx] === 0 ? obj.values[idx] : '') + '</c:v></c:pt>';
                    });
                    strXml += '    </c:numCache>';
                    strXml += '  </c:numRef>';
                    strXml += '</c:yVal>';
                }
                // E: '<c:bubbleSize>'
                strXml += '  <c:bubbleSize>';
                strXml += '    <c:numRef>';
                strXml += '      <c:f>Sheet1!' + '$' + getExcelColName(idxColLtr_1) + '$2:$' + getExcelColName(idx + 2) + '$' + (obj.sizes.length + 1) + '</c:f>';
                idxColLtr_1++;
                strXml += '      <c:numCache>';
                strXml += '        <c:formatCode>General</c:formatCode>';
                strXml += '	       <c:ptCount val="' + obj.sizes.length + '"/>';
                obj.sizes.forEach(function (value, idx) {
                    strXml += '<c:pt idx="' + idx + '"><c:v>' + (value || '') + '</c:v></c:pt>';
                });
                strXml += '      </c:numCache>';
                strXml += '    </c:numRef>';
                strXml += '  </c:bubbleSize>';
                strXml += '  <c:bubble3D val="0"/>';
                // F: Close "SERIES"
                strXml += '</c:ser>';
            });
            // 3: Data Labels
            {
                strXml += '  <c:dLbls>';
                strXml += '    <c:numFmt formatCode="' + opts.dataLabelFormatCode + '" sourceLinked="0"/>';
                strXml += '    <c:txPr>';
                strXml += '      <a:bodyPr/>';
                strXml += '      <a:lstStyle/>';
                strXml += '      <a:p><a:pPr>';
                strXml += '        <a:defRPr b="0" i="0" strike="noStrike" sz="' + (opts.dataLabelFontSize || DEF_FONT_SIZE) + '00" u="none">';
                strXml += '          <a:solidFill>' + createColorElement(opts.dataLabelColor || DEF_FONT_COLOR) + '</a:solidFill>';
                strXml += '          <a:latin typeface="' + (opts.dataLabelFontFace || 'Arial') + '"/>';
                strXml += '        </a:defRPr>';
                strXml += '      </a:pPr></a:p>';
                strXml += '    </c:txPr>';
                strXml += '    <c:dLblPos val="ctr"/>';
                strXml += '    <c:showLegendKey val="0"/>';
                strXml += '    <c:showVal val="' + (opts.showValue ? '1' : '0') + '"/>';
                strXml += '    <c:showCatName val="0"/>';
                strXml += '    <c:showSerName val="0"/>';
                strXml += '    <c:showPercent val="0"/>';
                strXml += '    <c:showBubbleSize val="0"/>';
                strXml += '  </c:dLbls>';
            }
            // 4: Add bubble options
            //strXml += '  <c:bubbleScale val="100"/>';
            //strXml += '  <c:showNegBubbles val="0"/>';
            // Commented out to let it default to PPT until we create options
            // 5: Add axisId (NOTE: order matters! (category comes first))
            strXml += '  <c:axId val="' + catAxisId + '"/>';
            strXml += '  <c:axId val="' + valAxisId + '"/>';
            // 6: Close Chart tag
            strXml += '</c:' + chartType + 'Chart>';
            // end switch
            break;
        case CHART_TYPES.DOUGHNUT:
        case CHART_TYPES.PIE:
            // Use the same let name so code blocks from barChart are interchangeable
            var obj = data[0];
            /* EX:
                data: [
                 {
                   name: 'Project Status',
                   labels: ['Red', 'Amber', 'Green', 'Unknown'],
                   values: [10, 20, 38, 2]
                 }
                ]
            */
            // 1: Start Chart
            strXml += '<c:' + chartType + 'Chart>';
            strXml += '  <c:varyColors val="0"/>';
            strXml += '<c:ser>';
            strXml += '  <c:idx val="0"/>';
            strXml += '  <c:order val="0"/>';
            strXml += '  <c:tx>';
            strXml += '    <c:strRef>';
            strXml += '      <c:f>Sheet1!$B$1</c:f>';
            strXml += '      <c:strCache>';
            strXml += '        <c:ptCount val="1"/>';
            strXml += '        <c:pt idx="0"><c:v>' + encodeXmlEntities(obj.name) + '</c:v></c:pt>';
            strXml += '      </c:strCache>';
            strXml += '    </c:strRef>';
            strXml += '  </c:tx>';
            strXml += '  <c:spPr>';
            strXml += '    <a:solidFill><a:schemeClr val="accent1"/></a:solidFill>';
            strXml += '    <a:ln w="9525" cap="flat"><a:solidFill><a:srgbClr val="F9F9F9"/></a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>';
            if (opts.dataNoEffects) {
                strXml += '<a:effectLst/>';
            }
            else {
                strXml += createShadowElement(opts.shadow, DEF_SHAPE_SHADOW);
            }
            strXml += '  </c:spPr>';
            strXml += '<c:explosion val="0"/>';
            // 2: "Data Point" block for every data row
            obj.labels.forEach(function (_label, idx) {
                strXml += '<c:dPt>';
                strXml += '  <c:idx val="' + idx + '"/>';
                strXml += '  <c:explosion val="0"/>';
                strXml += '  <c:spPr>';
                strXml +=
                    '    <a:solidFill>' +
                        createColorElement(opts.chartColors[idx + 1 > opts.chartColors.length ? Math.floor(Math.random() * opts.chartColors.length) : idx]) +
                        '</a:solidFill>';
                if (opts.dataBorder) {
                    strXml +=
                        '<a:ln w="' +
                            opts.dataBorder.pt * ONEPT +
                            '" cap="flat"><a:solidFill>' +
                            createColorElement(opts.dataBorder.color) +
                            '</a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>';
                }
                strXml += createShadowElement(opts.shadow, DEF_SHAPE_SHADOW);
                strXml += '  </c:spPr>';
                strXml += '</c:dPt>';
            });
            // 3: "Data Label" block for every data Label
            strXml += '<c:dLbls>';
            obj.labels.forEach(function (_label, idx) {
                strXml += '<c:dLbl>';
                strXml += '  <c:idx val="' + idx + '"/>';
                strXml += '    <c:numFmt formatCode="' + opts.dataLabelFormatCode + '" sourceLinked="0"/>';
                strXml += '    <c:txPr>';
                strXml += '      <a:bodyPr/><a:lstStyle/>';
                strXml += '      <a:p><a:pPr>';
                strXml +=
                    '        <a:defRPr b="' + (opts.dataLabelFontBold ? 1 : 0) + '" i="0" strike="noStrike" sz="' + (opts.dataLabelFontSize || DEF_FONT_SIZE) + '00" u="none">';
                strXml += '          <a:solidFill>' + createColorElement(opts.dataLabelColor || DEF_FONT_COLOR) + '</a:solidFill>';
                strXml += '          <a:latin typeface="' + (opts.dataLabelFontFace || 'Arial') + '"/>';
                strXml += '        </a:defRPr>';
                strXml += '      </a:pPr></a:p>';
                strXml += '    </c:txPr>';
                if (chartType === CHART_TYPES.PIE) {
                    strXml += '    <c:dLblPos val="' + (opts.dataLabelPosition || 'inEnd') + '"/>';
                }
                strXml += '    <c:showLegendKey val="0"/>';
                strXml += '    <c:showVal val="' + (opts.showValue ? '1' : '0') + '"/>';
                strXml += '    <c:showCatName val="' + (opts.showLabel ? '1' : '0') + '"/>';
                strXml += '    <c:showSerName val="0"/>';
                strXml += '    <c:showPercent val="' + (opts.showPercent ? '1' : '0') + '"/>';
                strXml += '    <c:showBubbleSize val="0"/>';
                strXml += '  </c:dLbl>';
            });
            strXml +=
                '<c:numFmt formatCode="' +
                    opts.dataLabelFormatCode +
                    '" sourceLinked="0"/>\
				<c:txPr>\
				  <a:bodyPr/>\
				  <a:lstStyle/>\
				  <a:p>\
					<a:pPr>\
					  <a:defRPr b="0" i="0" strike="noStrike" sz="1800" u="none">\
						<a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="Arial"/>\
					  </a:defRPr>\
					</a:pPr>\
				  </a:p>\
				</c:txPr>\
				' +
                    (chartType === CHART_TYPES.PIE ? '<c:dLblPos val="ctr"/>' : '') +
                    '\
				<c:showLegendKey val="0"/>\
				<c:showVal val="0"/>\
				<c:showCatName val="1"/>\
				<c:showSerName val="0"/>\
				<c:showPercent val="1"/>\
				<c:showBubbleSize val="0"/>\
				<c:showLeaderLines val="0"/>';
            strXml += '</c:dLbls>';
            // 2: "Categories"
            strXml += '<c:cat>';
            strXml += '  <c:strRef>';
            strXml += '    <c:f>Sheet1!' + '$A$2:$A$' + (obj.labels.length + 1) + '</c:f>';
            strXml += '    <c:strCache>';
            strXml += '	     <c:ptCount val="' + obj.labels.length + '"/>';
            obj.labels.forEach(function (label, idx) {
                strXml += '<c:pt idx="' + idx + '"><c:v>' + encodeXmlEntities(label) + '</c:v></c:pt>';
            });
            strXml += '    </c:strCache>';
            strXml += '  </c:strRef>';
            strXml += '</c:cat>';
            // 3: Create vals
            strXml += '  <c:val>';
            strXml += '    <c:numRef>';
            strXml += '      <c:f>Sheet1!' + '$B$2:$B$' + (obj.labels.length + 1) + '</c:f>';
            strXml += '      <c:numCache>';
            strXml += '	       <c:ptCount val="' + obj.labels.length + '"/>';
            obj.values.forEach(function (value, idx) {
                strXml += '<c:pt idx="' + idx + '"><c:v>' + (value || value === 0 ? value : '') + '</c:v></c:pt>';
            });
            strXml += '      </c:numCache>';
            strXml += '    </c:numRef>';
            strXml += '  </c:val>';
            // 4: Close "SERIES"
            strXml += '  </c:ser>';
            strXml += '  <c:firstSliceAng val="0"/>';
            if (chartType === CHART_TYPES.DOUGHNUT)
                strXml += '  <c:holeSize val="' + (opts.holeSize || 50) + '"/>';
            strXml += '</c:' + chartType + 'Chart>';
            // Done with Doughnut/Pie
            break;
        default:
            break;
    }
    return strXml;
}
/**
 * Create Category axis
 * @param {IChartOpts} opts - chart options
 * @param {string} axisId - value
 * @param {string} valAxisId - value
 * @return {string} XML
 */
function makeCatAxis(opts, axisId, valAxisId) {
    var strXml = '';
    // Build cat axis tag
    // NOTE: Scatter and Bubble chart need two Val axises as they display numbers on x axis
    if (opts.type === CHART_TYPES.SCATTER || opts.type === CHART_TYPES.BUBBLE) {
        strXml += '<c:valAx>';
    }
    else {
        strXml += '<c:' + (opts.catLabelFormatCode ? 'dateAx' : 'catAx') + '>';
    }
    strXml += '  <c:axId val="' + axisId + '"/>';
    strXml += '  <c:scaling>';
    strXml += '<c:orientation val="' + (opts.catAxisOrientation || (opts.barDir === 'col' ? 'minMax' : 'minMax')) + '"/>';
    if (opts.catAxisMaxVal || opts.catAxisMaxVal === 0)
        strXml += '<c:max val="' + opts.catAxisMaxVal + '"/>';
    if (opts.catAxisMinVal || opts.catAxisMinVal === 0)
        strXml += '<c:min val="' + opts.catAxisMinVal + '"/>';
    strXml += '</c:scaling>';
    strXml += '  <c:delete val="' + (opts.catAxisHidden ? 1 : 0) + '"/>';
    strXml += '  <c:axPos val="' + (opts.barDir === 'col' ? 'b' : 'l') + '"/>';
    strXml += opts.catGridLine.style !== 'none' ? createGridLineElement(opts.catGridLine) : '';
    // '<c:title>' comes between '</c:majorGridlines>' and '<c:numFmt>'
    if (opts.showCatAxisTitle) {
        strXml += genXmlTitle({
            color: opts.catAxisTitleColor,
            fontFace: opts.catAxisTitleFontFace,
            fontSize: opts.catAxisTitleFontSize,
            rotate: opts.catAxisTitleRotate,
            title: opts.catAxisTitle || 'Axis Title',
        });
    }
    // NOTE: Adding Val Axis Formatting if scatter or bubble charts
    if (opts.type === CHART_TYPES.SCATTER || opts.type === CHART_TYPES.BUBBLE) {
        strXml += '  <c:numFmt formatCode="' + (opts.valAxisLabelFormatCode ? opts.valAxisLabelFormatCode : 'General') + '" sourceLinked="0"/>';
    }
    else {
        strXml += '  <c:numFmt formatCode="' + (opts.catLabelFormatCode || 'General') + '" sourceLinked="0"/>';
    }
    if (opts.type === CHART_TYPES.SCATTER) {
        strXml += '  <c:majorTickMark val="none"/>';
        strXml += '  <c:minorTickMark val="none"/>';
        strXml += '  <c:tickLblPos val="nextTo"/>';
    }
    else {
        strXml += '  <c:majorTickMark val="' + (opts.catAxisMajorTickMark || 'out') + '"/>';
        strXml += '  <c:minorTickMark val="' + (opts.catAxisMajorTickMark || 'none') + '"/>';
        strXml += '  <c:tickLblPos val="' + (opts.catAxisLabelPos || opts.barDir === 'col' ? 'low' : 'nextTo') + '"/>';
    }
    strXml += '  <c:spPr>';
    strXml += '    <a:ln w="12700" cap="flat">';
    strXml += opts.catAxisLineShow === false ? '<a:noFill/>' : '<a:solidFill><a:srgbClr val="' + DEF_CHART_GRIDLINE.color + '"/></a:solidFill>';
    strXml += '      <a:prstDash val="solid"/>';
    strXml += '      <a:round/>';
    strXml += '    </a:ln>';
    strXml += '  </c:spPr>';
    strXml += '  <c:txPr>';
    strXml += '    <a:bodyPr ' + (opts.catAxisLabelRotate ? 'rot="' + convertRotationDegrees(opts.catAxisLabelRotate) + '"' : '') + '/>'; // don't specify rot 0 so we get the auto behavior
    strXml += '    <a:lstStyle/>';
    strXml += '    <a:p>';
    strXml += '    <a:pPr>';
    strXml += '    <a:defRPr sz="' + (opts.catAxisLabelFontSize || DEF_FONT_SIZE) + '00" b="' + (opts.catAxisLabelFontBold ? 1 : 0) + '" i="0" u="none" strike="noStrike">';
    strXml += '      <a:solidFill><a:srgbClr val="' + (opts.catAxisLabelColor || DEF_FONT_COLOR) + '"/></a:solidFill>';
    strXml += '      <a:latin typeface="' + (opts.catAxisLabelFontFace || 'Arial') + '"/>';
    strXml += '   </a:defRPr>';
    strXml += '  </a:pPr>';
    strXml += '  <a:endParaRPr lang="' + (opts.lang || 'en-US') + '"/>';
    strXml += '  </a:p>';
    strXml += ' </c:txPr>';
    strXml += ' <c:crossAx val="' + valAxisId + '"/>';
    strXml += ' <c:' + (typeof opts.valAxisCrossesAt === 'number' ? 'crossesAt' : 'crosses') + ' val="' + opts.valAxisCrossesAt + '"/>';
    strXml += ' <c:auto val="1"/>';
    strXml += ' <c:lblAlgn val="ctr"/>';
    strXml += ' <c:noMultiLvlLbl val="1"/>';
    if (opts.catAxisLabelFrequency)
        strXml += ' <c:tickLblSkip val="' + opts.catAxisLabelFrequency + '"/>';
    // Issue#149: PPT will auto-adjust these as needed after calcing the date bounds, so we only include them when specified by user
    if (opts.catLabelFormatCode) {
        ['catAxisBaseTimeUnit', 'catAxisMajorTimeUnit', 'catAxisMinorTimeUnit'].forEach(function (opt) {
            // Validate input as poorly chosen/garbage options will cause chart corruption and it wont render at all!
            if (opts[opt] && (typeof opts[opt] !== 'string' || ['days', 'months', 'years'].indexOf(opt.toLowerCase()) === -1)) {
                console.warn('`' + opt + "` must be one of: 'days','months','years' !");
                opts[opt] = null;
            }
        });
        if (opts.catAxisBaseTimeUnit)
            strXml += ' <c:baseTimeUnit  val="' + opts.catAxisBaseTimeUnit.toLowerCase() + '"/>';
        if (opts.catAxisMajorTimeUnit)
            strXml += ' <c:majorTimeUnit val="' + opts.catAxisMajorTimeUnit.toLowerCase() + '"/>';
        if (opts.catAxisMinorTimeUnit)
            strXml += ' <c:minorTimeUnit val="' + opts.catAxisMinorTimeUnit.toLowerCase() + '"/>';
        if (opts.catAxisMajorUnit)
            strXml += ' <c:majorUnit     val="' + opts.catAxisMajorUnit + '"/>';
        if (opts.catAxisMinorUnit)
            strXml += ' <c:minorUnit     val="' + opts.catAxisMinorUnit + '"/>';
    }
    // Close cat axis tag
    // NOTE: Added closing tag of val or cat axis based on chart type
    if (opts.type === CHART_TYPES.SCATTER || opts.type === CHART_TYPES.BUBBLE) {
        strXml += '</c:valAx>';
    }
    else {
        strXml += '</c:' + (opts.catLabelFormatCode ? 'dateAx' : 'catAx') + '>';
    }
    return strXml;
}
/**
 * Create Value Axis (Used by `bar3D`)
 * @param {IChartOpts} opts - chart options
 * @param {string} valAxisId - value
 * @return {string} XML
 */
function makeValAxis(opts, valAxisId) {
    var axisPos = valAxisId === AXIS_ID_VALUE_PRIMARY ? (opts.barDir === 'col' ? 'l' : 'b') : opts.barDir === 'col' ? 'r' : 't';
    var strXml = '';
    var isRight = axisPos === 'r' || axisPos === 't';
    var crosses = isRight ? 'max' : 'autoZero';
    var crossAxId = valAxisId === AXIS_ID_VALUE_PRIMARY ? AXIS_ID_CATEGORY_PRIMARY : AXIS_ID_CATEGORY_SECONDARY;
    strXml += '<c:valAx>';
    strXml += '  <c:axId val="' + valAxisId + '"/>';
    strXml += '  <c:scaling>';
    strXml += '    <c:orientation val="' + (opts.valAxisOrientation || (opts.barDir === 'col' ? 'minMax' : 'minMax')) + '"/>';
    if (opts.valAxisMaxVal || opts.valAxisMaxVal === 0)
        strXml += '<c:max val="' + opts.valAxisMaxVal + '"/>';
    if (opts.valAxisMinVal || opts.valAxisMinVal === 0)
        strXml += '<c:min val="' + opts.valAxisMinVal + '"/>';
    strXml += '  </c:scaling>';
    strXml += '  <c:delete val="' + (opts.valAxisHidden ? 1 : 0) + '"/>';
    strXml += '  <c:axPos val="' + axisPos + '"/>';
    if (opts.valGridLine.style !== 'none')
        strXml += createGridLineElement(opts.valGridLine);
    // '<c:title>' comes between '</c:majorGridlines>' and '<c:numFmt>'
    if (opts.showValAxisTitle) {
        strXml += genXmlTitle({
            color: opts.valAxisTitleColor,
            fontFace: opts.valAxisTitleFontFace,
            fontSize: opts.valAxisTitleFontSize,
            rotate: opts.valAxisTitleRotate,
            title: opts.valAxisTitle || 'Axis Title',
        });
    }
    strXml += ' <c:numFmt formatCode="' + (opts.valAxisLabelFormatCode ? opts.valAxisLabelFormatCode : 'General') + '" sourceLinked="0"/>';
    if (opts.type === CHART_TYPES.SCATTER) {
        strXml += '  <c:majorTickMark val="none"/>';
        strXml += '  <c:minorTickMark val="none"/>';
        strXml += '  <c:tickLblPos val="nextTo"/>';
    }
    else {
        strXml += ' <c:majorTickMark val="' + (opts.valAxisMajorTickMark || 'out') + '"/>';
        strXml += ' <c:minorTickMark val="' + (opts.valAxisMinorTickMark || 'none') + '"/>';
        strXml += ' <c:tickLblPos val="' + (opts.valAxisLabelPos || opts.barDir === 'col' ? 'nextTo' : 'low') + '"/>';
    }
    strXml += ' <c:spPr>';
    strXml += '   <a:ln w="12700" cap="flat">';
    strXml += opts.valAxisLineShow === false ? '<a:noFill/>' : '<a:solidFill><a:srgbClr val="' + DEF_CHART_GRIDLINE.color + '"/></a:solidFill>';
    strXml += '     <a:prstDash val="solid"/>';
    strXml += '     <a:round/>';
    strXml += '   </a:ln>';
    strXml += ' </c:spPr>';
    strXml += ' <c:txPr>';
    strXml += '  <a:bodyPr ' + (opts.valAxisLabelRotate ? 'rot="' + convertRotationDegrees(opts.valAxisLabelRotate) + '"' : '') + '/>'; // don't specify rot 0 so we get the auto behavior
    strXml += '  <a:lstStyle/>';
    strXml += '  <a:p>';
    strXml += '    <a:pPr>';
    strXml += '      <a:defRPr sz="' + (opts.valAxisLabelFontSize || DEF_FONT_SIZE) + '00" b="' + (opts.valAxisLabelFontBold ? 1 : 0) + '" i="0" u="none" strike="noStrike">';
    strXml += '        <a:solidFill><a:srgbClr val="' + (opts.valAxisLabelColor || DEF_FONT_COLOR) + '"/></a:solidFill>';
    strXml += '        <a:latin typeface="' + (opts.valAxisLabelFontFace || 'Arial') + '"/>';
    strXml += '      </a:defRPr>';
    strXml += '    </a:pPr>';
    strXml += '  <a:endParaRPr lang="' + (opts.lang || 'en-US') + '"/>';
    strXml += '  </a:p>';
    strXml += ' </c:txPr>';
    strXml += ' <c:crossAx val="' + crossAxId + '"/>';
    strXml += ' <c:crosses val="' + crosses + '"/>';
    strXml +=
        ' <c:crossBetween val="' +
            (opts.type === CHART_TYPES.SCATTER ||
                (Array.isArray(opts.type) &&
                    opts.type.filter(function (type) {
                        return type.type === CHART_TYPES.AREA;
                    }).length > 0
                    ? true
                    : false)
                ? 'midCat'
                : 'between') +
            '"/>';
    if (opts.valAxisMajorUnit)
        strXml += ' <c:majorUnit val="' + opts.valAxisMajorUnit + '"/>';
    strXml += '</c:valAx>';
    return strXml;
}
/**
 * Create Series Axis (Used by `bar3D`)
 * @param {IChartOpts} opts - chart options
 * @param {string} axisId - axis ID
 * @param {string} valAxisId - value
 * @return {string} XML
 */
function makeSerAxis(opts, axisId, valAxisId) {
    var strXml = '';
    // Build ser axis tag
    strXml += '<c:serAx>';
    strXml += '  <c:axId val="' + axisId + '"/>';
    strXml += '  <c:scaling><c:orientation val="' + (opts.serAxisOrientation || (opts.barDir === 'col' ? 'minMax' : 'minMax')) + '"/></c:scaling>';
    strXml += '  <c:delete val="' + (opts.serAxisHidden ? 1 : 0) + '"/>';
    strXml += '  <c:axPos val="' + (opts.barDir === 'col' ? 'b' : 'l') + '"/>';
    strXml += opts.serGridLine.style !== 'none' ? createGridLineElement(opts.serGridLine) : '';
    // '<c:title>' comes between '</c:majorGridlines>' and '<c:numFmt>'
    if (opts.showSerAxisTitle) {
        strXml += genXmlTitle({
            color: opts.serAxisTitleColor,
            fontFace: opts.serAxisTitleFontFace,
            fontSize: opts.serAxisTitleFontSize,
            rotate: opts.serAxisTitleRotate,
            title: opts.serAxisTitle || 'Axis Title',
        });
    }
    strXml += '  <c:numFmt formatCode="' + (opts.serLabelFormatCode || 'General') + '" sourceLinked="0"/>';
    strXml += '  <c:majorTickMark val="out"/>';
    strXml += '  <c:minorTickMark val="none"/>';
    strXml += '  <c:tickLblPos val="' + (opts.serAxisLabelPos || opts.barDir === 'col' ? 'low' : 'nextTo') + '"/>';
    strXml += '  <c:spPr>';
    strXml += '    <a:ln w="12700" cap="flat">';
    strXml += opts.serAxisLineShow === false ? '<a:noFill/>' : '<a:solidFill><a:srgbClr val="' + DEF_CHART_GRIDLINE.color + '"/></a:solidFill>';
    strXml += '      <a:prstDash val="solid"/>';
    strXml += '      <a:round/>';
    strXml += '    </a:ln>';
    strXml += '  </c:spPr>';
    strXml += '  <c:txPr>';
    strXml += '    <a:bodyPr/>'; // don't specify rot 0 so we get the auto behavior
    strXml += '    <a:lstStyle/>';
    strXml += '    <a:p>';
    strXml += '    <a:pPr>';
    strXml += '    <a:defRPr sz="' + (opts.serAxisLabelFontSize || DEF_FONT_SIZE) + '00" b="0" i="0" u="none" strike="noStrike">';
    strXml += '      <a:solidFill><a:srgbClr val="' + (opts.serAxisLabelColor || DEF_FONT_COLOR) + '"/></a:solidFill>';
    strXml += '      <a:latin typeface="' + (opts.serAxisLabelFontFace || 'Arial') + '"/>';
    strXml += '   </a:defRPr>';
    strXml += '  </a:pPr>';
    strXml += '  <a:endParaRPr lang="' + (opts.lang || 'en-US') + '"/>';
    strXml += '  </a:p>';
    strXml += ' </c:txPr>';
    strXml += ' <c:crossAx val="' + valAxisId + '"/>';
    strXml += ' <c:crosses val="autoZero"/>';
    if (opts.serAxisLabelFrequency)
        strXml += ' <c:tickLblSkip val="' + opts.serAxisLabelFrequency + '"/>';
    // Issue#149: PPT will auto-adjust these as needed after calcing the date bounds, so we only include them when specified by user
    if (opts.serLabelFormatCode) {
        ['serAxisBaseTimeUnit', 'serAxisMajorTimeUnit', 'serAxisMinorTimeUnit'].forEach(function (opt) {
            // Validate input as poorly chosen/garbage options will cause chart corruption and it wont render at all!
            if (opts[opt] && (typeof opts[opt] !== 'string' || ['days', 'months', 'years'].indexOf(opt.toLowerCase()) === -1)) {
                console.warn('`' + opt + "` must be one of: 'days','months','years' !");
                opts[opt] = null;
            }
        });
        if (opts.serAxisBaseTimeUnit)
            strXml += ' <c:baseTimeUnit  val="' + opts.serAxisBaseTimeUnit.toLowerCase() + '"/>';
        if (opts.serAxisMajorTimeUnit)
            strXml += ' <c:majorTimeUnit val="' + opts.serAxisMajorTimeUnit.toLowerCase() + '"/>';
        if (opts.serAxisMinorTimeUnit)
            strXml += ' <c:minorTimeUnit val="' + opts.serAxisMinorTimeUnit.toLowerCase() + '"/>';
        if (opts.serAxisMajorUnit)
            strXml += ' <c:majorUnit     val="' + opts.serAxisMajorUnit + '"/>';
        if (opts.serAxisMinorUnit)
            strXml += ' <c:minorUnit     val="' + opts.serAxisMinorUnit + '"/>';
    }
    // Close ser axis tag
    strXml += '</c:serAx>';
    return strXml;
}
/**
 * Create char title elements
 * @param {IChartTitleOpts} opts - options
 * @return {string} XML `<c:title>`
 */
function genXmlTitle(opts) {
    var align = opts.titleAlign === 'left' || opts.titleAlign === 'right' ? "<a:pPr algn=\"" + opts.titleAlign.substring(0, 1) + "\">" : "<a:pPr>";
    var rotate = opts.rotate ? "<a:bodyPr rot=\"" + convertRotationDegrees(opts.rotate) + "\"/>" : "<a:bodyPr/>"; // don't specify rotation to get default (ex. vertical for cat axis)
    var sizeAttr = opts.fontSize ? 'sz="' + Math.round(opts.fontSize) + '00"' : ''; // only set the font size if specified.  Powerpoint will handle the default size
    var layout = opts.titlePos && opts.titlePos.x && opts.titlePos.y
        ? "<c:layout><c:manualLayout><c:xMode val=\"edge\"/><c:yMode val=\"edge\"/><c:x val=\"" + opts.titlePos.x + "\"/><c:y val=\"" + opts.titlePos.y + "\"/></c:manualLayout></c:layout>"
        : "<c:layout/>";
    return "<c:title>\n\t  <c:tx>\n\t    <c:rich>\n\t      " + rotate + "\n\t      <a:lstStyle/>\n\t      <a:p>\n\t        " + align + "\n\t        <a:defRPr " + sizeAttr + " b=\"0\" i=\"0\" u=\"none\" strike=\"noStrike\">\n\t          <a:solidFill><a:srgbClr val=\"" + (opts.color || DEF_FONT_COLOR) + "\"/></a:solidFill>\n\t          <a:latin typeface=\"" + (opts.fontFace || 'Arial') + "\"/>\n\t        </a:defRPr>\n\t      </a:pPr>\n\t      <a:r>\n\t        <a:rPr " + sizeAttr + " b=\"0\" i=\"0\" u=\"none\" strike=\"noStrike\">\n\t          <a:solidFill><a:srgbClr val=\"" + (opts.color || DEF_FONT_COLOR) + "\"/></a:solidFill>\n\t          <a:latin typeface=\"" + (opts.fontFace || 'Arial') + "\"/>\n\t        </a:rPr>\n\t        <a:t>" + (encodeXmlEntities(opts.title) || '') + "</a:t>\n\t      </a:r>\n\t    </a:p>\n\t    </c:rich>\n\t  </c:tx>\n\t  " + layout + "\n\t  <c:overlay val=\"0\"/>\n\t</c:title>";
}
/**
 * Calc and return excel column name for a given column length
 * @param {number} length - col length
 * @return {string} column name (ex: 'A2')
 */
function getExcelColName(length) {
    var strName = '';
    if (length <= 26) {
        strName = LETTERS[length];
    }
    else {
        strName += LETTERS[Math.floor(length / LETTERS.length) - 1];
        strName += LETTERS[length % LETTERS.length];
    }
    return strName;
}
/**
 * NOTE: Used by both: text and lineChart
 * Creates `a:innerShdw` or `a:outerShdw` depending on pass options `opts`.
 * @param {Object} opts optional shadow properties
 * @param {Object} defaults defaults for unspecified properties in `opts`
 * @see http://officeopenxml.com/drwSp-effects.php
 *	{ type: 'outer', blur: 3, offset: (23000 / 12700), angle: 90, color: '000000', opacity: 0.35, rotateWithShape: true };
 * @return {string} XML
 */
function createShadowElement(options, defaults) {
    if (options === null) {
        return '<a:effectLst/>';
    }
    var strXml = '<a:effectLst>', opts = getMix(defaults, options), type = opts['type'] || 'outer', blur = opts['blur'] * ONEPT, offset = opts['offset'] * ONEPT, angle = opts['angle'] * 60000, color = opts['color'], opacity = opts['opacity'] * 100000, rotateWithShape = opts['rotateWithShape'] ? 1 : 0;
    strXml += '<a:' + type + 'Shdw sx="100000" sy="100000" kx="0" ky="0"  algn="bl" blurRad="' + blur + '" ';
    strXml += 'rotWithShape="' + +rotateWithShape + '"';
    strXml += ' dist="' + offset + '" dir="' + angle + '">';
    strXml += '<a:srgbClr val="' + color + '">'; // TODO: should accept scheme colors implemented in Issue #135
    strXml += '<a:alpha val="' + opacity + '"/></a:srgbClr>';
    strXml += '</a:' + type + 'Shdw>';
    strXml += '</a:effectLst>';
    return strXml;
}
/**
 * Create Grid Line Element
 * @param {OptsChartGridLine} glOpts {size, color, style}
 * @return {string} XML
 */
function createGridLineElement(glOpts) {
    var strXml = '<c:majorGridlines>';
    strXml += ' <c:spPr>';
    strXml += '  <a:ln w="' + Math.round((glOpts.size || DEF_CHART_GRIDLINE.size) * ONEPT) + '" cap="flat">';
    strXml += '  <a:solidFill><a:srgbClr val="' + (glOpts.color || DEF_CHART_GRIDLINE.color) + '"/></a:solidFill>'; // should accept scheme colors as implemented in [Pull #135]
    strXml += '   <a:prstDash val="' + (glOpts.style || DEF_CHART_GRIDLINE.style) + '"/><a:round/>';
    strXml += '  </a:ln>';
    strXml += ' </c:spPr>';
    strXml += '</c:majorGridlines>';
    return strXml;
}

/**
 * PptxGenJS: Media Methods
 */
/**
 * Encode Image/Audio/Video into base64
 * @param {ISlide | ISlideLayout} layout - slide layout
 * @return {Promise} promise of generating the rels
 */
function encodeSlideMediaRels(layout) {
    var fs = typeof require !== 'undefined' && typeof window === 'undefined' ? require('fs') : null; // NodeJS
    var https = typeof require !== 'undefined' && typeof window === 'undefined' ? require('https') : null; // NodeJS
    var imageProms = [];
    // A: Read/Encode each audio/image/video thats not already encoded (eg: base64 provided by user)
    layout.relsMedia
        .filter(function (rel) {
        return rel.type !== 'online' && !rel.data;
    })
        .forEach(function (rel) {
        imageProms.push(new Promise(function (resolve, reject) {
            if (fs && rel.path.indexOf('http') !== 0) {
                // DESIGN: Node local-file encoding is syncronous, so we can load all images here, then call export with a callback (if any)
                try {
                    var bitmap = fs.readFileSync(rel.path);
                    rel.data = Buffer.from(bitmap).toString('base64');
                    resolve('done');
                }
                catch (ex) {
                    rel.data = IMG_BROKEN;
                    reject('ERROR: Unable to read media: "' + rel.path + '"\n' + ex.toString());
                }
            }
            else if (fs && https && rel.path.indexOf('http') === 0) {
                https.get(rel.path, function (res) {
                    var rawData = '';
                    res.setEncoding('binary'); // IMPORTANT: Only binary encoding works
                    res.on('data', function (chunk) { return (rawData += chunk); });
                    res.on('end', function () {
                        rel.data = Buffer.from(rawData, 'binary').toString('base64');
                        resolve('done');
                    });
                    res.on('error', function (ex) {
                        rel.data = IMG_BROKEN;
                        reject('ERROR: Unable to load image: "' + rel.path + '"\n' + ex.toString());
                    });
                });
            }
            else {
                // A: Declare XHR and onload/onerror handlers
                // DESIGN: `XMLHttpRequest()` plus `FileReader()` = Ablity to read any file into base64!
                var xhr_1 = new XMLHttpRequest();
                xhr_1.onload = function () {
                    var reader = new FileReader();
                    reader.onloadend = function () {
                        rel.data = reader.result;
                        if (!rel.isSvgPng) {
                            resolve('done');
                        }
                        else {
                            createSvgPngPreview(rel)
                                .then(function () {
                                resolve('done');
                            })
                                .catch(function (ex) {
                                reject(ex.toString());
                            });
                        }
                    };
                    reader.readAsDataURL(xhr_1.response);
                };
                xhr_1.onerror = function (ex) {
                    rel.data = IMG_BROKEN;
                    reject('ERROR: Unable to load image: "' + rel.path + '"\n' + ex.toString());
                };
                // B: Execute request
                xhr_1.open('GET', rel.path);
                xhr_1.responseType = 'blob';
                xhr_1.send();
            }
        }));
    });
    // B: SVG: base64 data still requires a png to be generated (`isSvgPng` flag this as the preview image, not the SVG itself)
    layout.relsMedia
        .filter(function (rel) {
        return rel.isSvgPng && rel.data;
    })
        .forEach(function (rel) {
        if (fs) {
            //console.log('Sorry, SVG is not supported in Node (more info: https://github.com/gitbrent/PptxGenJS/issues/401)')
            rel.data = IMG_BROKEN;
            imageProms.push(Promise.resolve().then(function () {
                return 'done';
            }));
        }
        else {
            imageProms.push(createSvgPngPreview(rel));
        }
    });
    return imageProms;
}
/**
 * Create SVG preview image
 * @param {ISlideRelMedia} rel - slide rel
 * @return {Promise} promise
 */
function createSvgPngPreview(rel) {
    return new Promise(function (resolve, reject) {
        // A: Create
        var image = new Image();
        // B: Set onload event
        image.onload = function () {
            // First: Check for any errors: This is the best method (try/catch wont work, etc.)
            if (image.width + image.height === 0) {
                image.onerror('h/w=0');
            }
            var canvas = document.createElement('CANVAS');
            var ctx = canvas.getContext('2d');
            canvas.width = image.width;
            canvas.height = image.height;
            ctx.drawImage(image, 0, 0);
            // Users running on local machine will get the following error:
            // "SecurityError: Failed to execute 'toDataURL' on 'HTMLCanvasElement': Tainted canvases may not be exported."
            // when the canvas.toDataURL call executes below.
            try {
                rel.data = canvas.toDataURL(rel.type);
                resolve('done');
            }
            catch (ex) {
                image.onerror(ex);
            }
            canvas = null;
        };
        image.onerror = function (ex) {
            rel.data = IMG_BROKEN;
            reject(ex.toString());
        };
        // C: Load image
        image.src = typeof rel.data === 'string' ? rel.data : IMG_BROKEN;
    });
}

/*\
|*|  :: pptxgen.ts ::
|*|
|*|  JavaScript framework that creates PowerPoint (pptx) presentations
|*|  https://github.com/gitbrent/PptxGenJS
|*|
|*|  This framework is released under the MIT Public License (MIT)
|*|
|*|  PptxGenJS (C) 2015-2020 Brent Ely -- https://github.com/gitbrent
|*|
|*|  Some code derived from the OfficeGen project:
|*|  github.com/Ziv-Barber/officegen/ (Copyright 2013 Ziv Barber)
|*|
|*|  Permission is hereby granted, free of charge, to any person obtaining a copy
|*|  of this software and associated documentation files (the "Software"), to deal
|*|  in the Software without restriction, including without limitation the rights
|*|  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
|*|  copies of the Software, and to permit persons to whom the Software is
|*|  furnished to do so, subject to the following conditions:
|*|
|*|  The above copyright notice and this permission notice shall be included in all
|*|  copies or substantial portions of the Software.
|*|
|*|  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
|*|  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
|*|  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
|*|  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
|*|  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
|*|  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
|*|  SOFTWARE.
\*/
var PptxGenJS = /** @class */ (function () {
    function PptxGenJS() {
        var _this = this;
        /**
         * Library Version
         */
        this._version = '3.0.0-beta.7';
        // Global props
        this._charts = CHART_TYPES;
        this._colors = SCHEME_COLOR_NAMES;
        this._shapes = PowerPointShapes;
        /**
         * Provides an API for `addTableDefinition` to create slides as needed for auto-paging
         * @param {string} masterName - slide master name
         * @return {ISlide} new Slide
         */
        this.addNewSlide = function (masterName) {
            return _this.addSlide(masterName);
        };
        /**
         * Provides an API for `addTableDefinition` to create slides as needed for auto-paging
         * @since 3.0.0
         * @param {number} slideNum - slide number
         * @return {ISlide} Slide
         */
        this.getSlide = function (slideNum) {
            return _this.slides.filter(function (slide) {
                return slide.number === slideNum;
            })[0];
        };
        /**
         * Enables the `Slide` class to set PptxGenJS [Presentation] master/layout slidenumbers
         * @param {ISlideNumber} slideNum - slide number config
         */
        this.setSlideNumber = function (slideNum) {
            // 1: Add slideNumber to slideMaster1.xml
            _this.masterSlide.slideNumberObj = slideNum;
            // 2: Add slideNumber to DEF_PRES_LAYOUT_NAME layout
            _this.slideLayouts.filter(function (layout) {
                return layout.name === DEF_PRES_LAYOUT_NAME;
            })[0].slideNumberObj = slideNum;
        };
        /**
         * Create all chart and media rels for this Presenation
         * @param {ISlide | ISlideLayout} slide - slide with rels
         * @param {JSZIP} zip - JSZip instance
         * @param {Promise<any>[]} chartPromises - promise array
         */
        this.createChartMediaRels = function (slide, zip, chartPromises) {
            slide.relsChart.forEach(function (rel) { return chartPromises.push(createExcelWorksheet(rel, zip)); });
            slide.relsMedia.forEach(function (rel) {
                if (rel.type !== 'online' && rel.type !== 'hyperlink') {
                    // A: Loop vars
                    var data = rel.data && typeof rel.data === 'string' ? rel.data : '';
                    // B: Users will undoubtedly pass various string formats, so correct prefixes as needed
                    if (data.indexOf(',') === -1 && data.indexOf(';') === -1)
                        data = 'image/png;base64,' + data;
                    else if (data.indexOf(',') === -1)
                        data = 'image/png;base64,' + data;
                    else if (data.indexOf(';') === -1)
                        data = 'image/png;' + data;
                    // C: Add media
                    zip.file(rel.Target.replace('..', 'ppt'), data.split(',').pop(), { base64: true });
                }
            });
        };
        /**
         * Create and export the .pptx file
         * @param {string} exportName - output file type
         * @param {Blob} blobContent - Blob content
         * @return {Promise<string>} Promise with file name
         */
        this.writeFileToBrowser = function (exportName, blobContent) {
            return new Promise(function (resolve, _reject) {
                // STEP 1: Create element
                var eleLink = document.createElement('a');
                eleLink.setAttribute('style', 'display:none;');
                document.body.appendChild(eleLink);
                // STEP 2: Download file to browser
                // DESIGN: Use `createObjectURL()` (or MS-specific func for IE11) to D/L files in client browsers (FYI: synchronously executed)
                if (window.navigator.msSaveOrOpenBlob) {
                    // @see https://docs.microsoft.com/en-us/microsoft-edge/dev-guide/html5/file-api/blob
                    var blob_1 = new Blob([blobContent], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
                    eleLink.onclick = function () {
                        window.navigator.msSaveOrOpenBlob(blob_1, exportName);
                    };
                    eleLink.click();
                    // Clean-up
                    document.body.removeChild(eleLink);
                    // Done
                    resolve(exportName);
                }
                else if (window.URL.createObjectURL) {
                    var url_1 = window.URL.createObjectURL(new Blob([blobContent], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' }));
                    eleLink.href = url_1;
                    eleLink.download = exportName;
                    eleLink.click();
                    // Clean-up (NOTE: Add a slight delay before removing to avoid 'blob:null' error in Firefox Issue#81)
                    setTimeout(function () {
                        window.URL.revokeObjectURL(url_1);
                        document.body.removeChild(eleLink);
                    }, 100);
                    // Done
                    resolve(exportName);
                }
            });
        };
        /**
         * Create and export the .pptx file
         * @param {WRITE_OUTPUT_TYPE} outputType - output file type
         * @return {Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array>} Promise with data or stream (node) or filename (browser)
         */
        this.exportPresentation = function (outputType) {
            return new Promise(function (resolve, reject) {
                var arrChartPromises = [];
                var arrMediaPromises = [];
                var zip = new JSZip();
                // STEP 1: Read/Encode all Media before zip as base64 content, etc. is required
                _this.slides.forEach(function (slide) {
                    arrMediaPromises = arrMediaPromises.concat(encodeSlideMediaRels(slide));
                });
                _this.slideLayouts.forEach(function (layout) {
                    arrMediaPromises = arrMediaPromises.concat(encodeSlideMediaRels(layout));
                });
                arrMediaPromises = arrMediaPromises.concat(encodeSlideMediaRels(_this.masterSlide));
                // STEP 2: Wait for Promises (if any) then generate the PPTX file
                Promise.all(arrMediaPromises).then(function () {
                    // A: Add empty placeholder objects to slides that don't already have them
                    _this.slides.forEach(function (slide) {
                        if (slide.slideLayout)
                            addPlaceholdersToSlideLayouts(slide);
                    });
                    // B: Add all required folders and files
                    zip.folder('_rels');
                    zip.folder('docProps');
                    zip.folder('ppt').folder('_rels');
                    zip.folder('ppt/charts').folder('_rels');
                    zip.folder('ppt/embeddings');
                    zip.folder('ppt/media');
                    zip.folder('ppt/slideLayouts').folder('_rels');
                    zip.folder('ppt/slideMasters').folder('_rels');
                    zip.folder('ppt/slides').folder('_rels');
                    zip.folder('ppt/theme');
                    zip.folder('ppt/notesMasters').folder('_rels');
                    zip.folder('ppt/notesSlides').folder('_rels');
                    zip.file('[Content_Types].xml', makeXmlContTypes(_this.slides, _this.slideLayouts, _this.masterSlide));
                    zip.file('_rels/.rels', makeXmlRootRels());
                    zip.file('docProps/app.xml', makeXmlApp(_this.slides, _this.company));
                    zip.file('docProps/core.xml', makeXmlCore(_this.title, _this.subject, _this.author, _this.revision));
                    zip.file('ppt/_rels/presentation.xml.rels', makeXmlPresentationRels(_this.slides));
                    zip.file('ppt/theme/theme1.xml', makeXmlTheme());
                    zip.file('ppt/presentation.xml', makeXmlPresentation(_this.slides, _this.presLayout, _this.rtlMode));
                    zip.file('ppt/presProps.xml', makeXmlPresProps());
                    zip.file('ppt/tableStyles.xml', makeXmlTableStyles());
                    zip.file('ppt/viewProps.xml', makeXmlViewProps());
                    // C: Create a Layout/Master/Rel/Slide file for each SlideLayout and Slide
                    _this.slideLayouts.forEach(function (layout, idx) {
                        zip.file('ppt/slideLayouts/slideLayout' + (idx + 1) + '.xml', makeXmlLayout(layout));
                        zip.file('ppt/slideLayouts/_rels/slideLayout' + (idx + 1) + '.xml.rels', makeXmlSlideLayoutRel(idx + 1, _this.slideLayouts));
                    });
                    _this.slides.forEach(function (slide, idx) {
                        zip.file('ppt/slides/slide' + (idx + 1) + '.xml', makeXmlSlide(slide));
                        zip.file('ppt/slides/_rels/slide' + (idx + 1) + '.xml.rels', makeXmlSlideRel(_this.slides, _this.slideLayouts, idx + 1));
                        // Create all slide notes related items. Notes of empty strings are created for slides which do not have notes specified, to keep track of _rels.
                        zip.file('ppt/notesSlides/notesSlide' + (idx + 1) + '.xml', makeXmlNotesSlide(slide));
                        zip.file('ppt/notesSlides/_rels/notesSlide' + (idx + 1) + '.xml.rels', makeXmlNotesSlideRel(idx + 1));
                    });
                    zip.file('ppt/slideMasters/slideMaster1.xml', makeXmlMaster(_this.masterSlide, _this.slideLayouts));
                    zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', makeXmlMasterRel(_this.masterSlide, _this.slideLayouts));
                    zip.file('ppt/notesMasters/notesMaster1.xml', makeXmlNotesMaster());
                    zip.file('ppt/notesMasters/_rels/notesMaster1.xml.rels', makeXmlNotesMasterRel());
                    // D: Create all Rels (images, media, chart data)
                    _this.slideLayouts.forEach(function (layout) {
                        _this.createChartMediaRels(layout, zip, arrChartPromises);
                    });
                    _this.slides.forEach(function (slide) {
                        _this.createChartMediaRels(slide, zip, arrChartPromises);
                    });
                    _this.createChartMediaRels(_this.masterSlide, zip, arrChartPromises);
                    // E: Wait for Promises (if any) then generate the PPTX file
                    Promise.all(arrChartPromises)
                        .then(function () {
                        if (outputType === 'STREAM') {
                            // A: stream file
                            zip.generateAsync({ type: 'nodebuffer' }).then(function (content) {
                                resolve(content);
                            });
                        }
                        else if (outputType) {
                            // B: Node [fs]: Output type user option or default
                            resolve(zip.generateAsync({ type: outputType }));
                        }
                        else {
                            // C: Browser: Output blob as app/ms-pptx
                            resolve(zip.generateAsync({ type: 'blob' }));
                        }
                    })
                        .catch(function (err) {
                        reject(err);
                    });
                });
            });
        };
        // Set available layouts
        this.LAYOUTS = {
            LAYOUT_4x3: { name: 'screen4x3', width: 9144000, height: 6858000 },
            LAYOUT_16x9: { name: 'screen16x9', width: 9144000, height: 5143500 },
            LAYOUT_16x10: { name: 'screen16x10', width: 9144000, height: 5715000 },
            LAYOUT_WIDE: { name: 'custom', width: 12192000, height: 6858000 },
        };
        // Core
        this._author = 'PptxGenJS';
        this._company = 'PptxGenJS';
        this._revision = '1'; // Note: Must be a whole number
        this._subject = 'PptxGenJS Presentation';
        this._title = 'PptxGenJS Presentation';
        // PptxGenJS props
        this._presLayout = {
            name: this.LAYOUTS[DEF_PRES_LAYOUT].name,
            width: this.LAYOUTS[DEF_PRES_LAYOUT].width,
            height: this.LAYOUTS[DEF_PRES_LAYOUT].height,
        };
        this._rtlMode = false;
        this._isBrowser = false;
        //
        this.slideLayouts = [
            {
                presLayout: this._presLayout,
                name: DEF_PRES_LAYOUT_NAME,
                number: 1000,
                slide: null,
                data: [],
                rels: [],
                relsChart: [],
                relsMedia: [],
                margin: DEF_SLIDE_MARGIN_IN,
                slideNumberObj: null,
            },
        ];
        this.slides = [];
        this.masterSlide = {
            addChart: null,
            addImage: null,
            addMedia: null,
            addNotes: null,
            addShape: null,
            addTable: null,
            addText: null,
            //
            presLayout: this._presLayout,
            name: null,
            number: null,
            data: [],
            rels: [],
            relsChart: [],
            relsMedia: [],
            slideLayout: null,
            slideNumberObj: null,
        };
    }
    Object.defineProperty(PptxGenJS.prototype, "layout", {
        get: function () {
            return this._layout;
        },
        set: function (value) {
            var newLayout = this.LAYOUTS[value];
            if (newLayout) {
                this._layout = value;
                this._presLayout = newLayout;
            }
            else {
                throw 'UNKNOWN-LAYOUT';
            }
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(PptxGenJS.prototype, "version", {
        get: function () {
            return this._version;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(PptxGenJS.prototype, "author", {
        get: function () {
            return this._author;
        },
        set: function (value) {
            this._author = value;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(PptxGenJS.prototype, "company", {
        get: function () {
            return this._company;
        },
        set: function (value) {
            this._company = value;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(PptxGenJS.prototype, "revision", {
        get: function () {
            return this._revision;
        },
        set: function (value) {
            this._revision = value;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(PptxGenJS.prototype, "subject", {
        get: function () {
            return this._subject;
        },
        set: function (value) {
            this._subject = value;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(PptxGenJS.prototype, "title", {
        get: function () {
            return this._title;
        },
        set: function (value) {
            this._title = value;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(PptxGenJS.prototype, "rtlMode", {
        get: function () {
            return this._rtlMode;
        },
        set: function (value) {
            this._rtlMode = value;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(PptxGenJS.prototype, "isBrowser", {
        get: function () {
            return this._isBrowser;
        },
        set: function (value) {
            this._isBrowser = value;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(PptxGenJS.prototype, "charts", {
        get: function () {
            return this._charts;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(PptxGenJS.prototype, "colors", {
        get: function () {
            return this._colors;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(PptxGenJS.prototype, "shapes", {
        get: function () {
            return this._shapes;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(PptxGenJS.prototype, "presLayout", {
        get: function () {
            return this._presLayout;
        },
        enumerable: true,
        configurable: true
    });
    // EXPORT METHODS
    /**
     * Export the current Presenation to stream
     * @since 3.0.0
     * @returns {Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array>} file stream
     */
    PptxGenJS.prototype.stream = function () {
        var _this = this;
        return new Promise(function (resolve, reject) {
            _this.exportPresentation('STREAM')
                .then(function (content) {
                resolve(content);
            })
                .catch(function (ex) {
                reject(ex);
            });
        });
    };
    /**
     * Export the current Presenation as JSZip content with the selected type
     * @since 3.0.0
     * @param {JSZIP_OUTPUT_TYPE} outputType - 'arraybuffer' | 'base64' | 'binarystring' | 'blob' | 'nodebuffer' | 'uint8array'
     * @returns {Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array>} file content in selected type
     */
    PptxGenJS.prototype.write = function (outputType) {
        var _this = this;
        return new Promise(function (resolve, reject) {
            _this.exportPresentation(outputType)
                .then(function (content) {
                resolve(content);
            })
                .catch(function (ex) {
                reject(ex + '\nDid you mean to use writeFile() instead?');
            });
        });
    };
    /**
     * Export the current Presenation. Writes file to local file system if `fs` exists, otherwise, initiates download in browsers
     * @since 3.0.0
     * @param {string} exportName - file name
     * @returns {Promise<string>} the presentation name
     */
    PptxGenJS.prototype.writeFile = function (exportName) {
        var _this = this;
        return new Promise(function (resolve, reject) {
            var fs = typeof require !== 'undefined' && typeof window === 'undefined' ? require('fs') : null; // NodeJS
            var fileName = exportName
                ? exportName
                    .toString()
                    .toLowerCase()
                    .endsWith('.pptx')
                    ? exportName
                    : exportName + '.pptx'
                : 'Presenation.pptx';
            _this.exportPresentation(fs ? 'nodebuffer' : null)
                .then(function (content) {
                if (fs) {
                    // Node: Output
                    fs.writeFile(fileName, content, function () {
                        resolve(fileName);
                    });
                }
                else {
                    // Browser: Output blob as app/ms-pptx
                    resolve(_this.writeFileToBrowser(fileName, content));
                }
            })
                .catch(function (ex) {
                reject(ex);
            });
        });
    };
    // PRESENTATION METHODS
    /**
     * Add a Slide to Presenation
     * @param {string} masterSlideName - Master Slide name
     * @returns {ISlide} the new Slide
     */
    PptxGenJS.prototype.addSlide = function (masterSlideName) {
        var newSlide = new Slide({
            addSlide: this.addNewSlide,
            getSlide: this.getSlide,
            presLayout: this.presLayout,
            setSlideNum: this.setSlideNumber,
            slideNumber: this.slides.length + 1,
            slideLayout: masterSlideName
                ? this.slideLayouts.filter(function (layout) {
                    return layout.name === masterSlideName;
                })[0] || this.LAYOUTS[DEF_PRES_LAYOUT]
                : this.LAYOUTS[DEF_PRES_LAYOUT],
        });
        this.slides.push(newSlide);
        return newSlide;
    };
    /**
     * Define a custom Slide Layout
     * @example pptx.defineLayout({ name:'A3', width:16.5, height:11.7 });
     * @see https://support.office.com/en-us/article/Change-the-size-of-your-slides-040a811c-be43-40b9-8d04-0de5ed79987e
     * @param {IUserLayout} layout - an object with user-defined w/h
     */
    PptxGenJS.prototype.defineLayout = function (layout) {
        if (!layout)
            console.warn('defineLayout requires `{name, width, height}`');
        else if (!layout.name)
            console.warn('defineLayout requires `name`');
        else if (!layout.width)
            console.warn('defineLayout requires `width`');
        else if (!layout.height)
            console.warn('defineLayout requires `height`');
        else if (typeof layout.height !== 'number')
            console.warn('defineLayout `height` should be a number (inches)');
        else if (typeof layout.width !== 'number')
            console.warn('defineLayout `width` should be a number (inches)');
        this.LAYOUTS[layout.name] = { name: layout.name, width: Math.round(Number(layout.width) * EMU), height: Math.round(Number(layout.height) * EMU) };
    };
    /**
     * Adds a new slide master [layout] to the Presentation
     * @param {ISlideMasterOptions} slideMasterOpts - layout definition
     */
    PptxGenJS.prototype.defineSlideMaster = function (slideMasterOpts) {
        if (!slideMasterOpts.title)
            throw Error('defineSlideMaster() object argument requires a `title` value. (https://gitbrent.github.io/PptxGenJS/docs/masters.html)');
        var newLayout = {
            presLayout: this.presLayout,
            name: slideMasterOpts.title,
            number: 1000 + this.slideLayouts.length + 1,
            slide: null,
            data: [],
            rels: [],
            relsChart: [],
            relsMedia: [],
            margin: slideMasterOpts.margin || DEF_SLIDE_MARGIN_IN,
            slideNumberObj: slideMasterOpts.slideNumber || null,
        };
        // STEP 1: Create the Slide Master/Layout
        createSlideObject(slideMasterOpts, newLayout);
        // STEP 2: Add it to layout defs
        this.slideLayouts.push(newLayout);
        // STEP 3: Add slideNumber to master slide (if any)
        if (newLayout.slideNumberObj && !this.masterSlide.slideNumberObj)
            this.masterSlide.slideNumberObj = newLayout.slideNumberObj;
    };
    // HTML-TO-SLIDES METHODS
    /**
     * Reproduces an HTML table as a PowerPoint table - including column widths, style, etc. - creates 1 or more slides as needed
     * @note `verbose` option is undocumented; used for verbose output of layout process
     * @param {string} tabEleId - HTMLElementID of the table
     * @param {ITableToSlidesOpts} inOpts - array of options (e.g.: tabsize)
     */
    PptxGenJS.prototype.tableToSlides = function (tableElementId, opts) {
        if (opts === void 0) { opts = {}; }
        genTableToSlides(this, tableElementId, opts, opts && opts.masterSlideName
            ? this.slideLayouts.filter(function (layout) {
                return layout.name === opts.masterSlideName;
            })[0]
            : null);
    };
    return PptxGenJS;
}());

export default PptxGenJS;
