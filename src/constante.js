const LAYOUTS = {
    'LAYOUT_4x3': {
        name: 'screen4x3',
        width: 9144000,
        height: 6858000
    },
    'LAYOUT_16x9': {
        name: 'screen16x9',
        width: 9144000,
        height: 5143500
    },
    'LAYOUT_16x10': {
        name: 'screen16x10',
        width: 9144000,
        height: 5715000
    },
    'LAYOUT_WIDE': {
        name: 'custom',
        width: 12191996,
        height: 6858000
    }
};

const BASE_SHAPES = {
    RECTANGLE: {
        'displayName': 'Rectangle',
        'name': 'rect',
        'avLst': {}
    },
    LINE: {
        'displayName': 'Line',
        'name': 'line',
        'avLst': {}
    }
};

const APP_VER = "1.1.0";


export {
    LAYOUTS,
    BASE_SHAPES,
    APP_VER
};
