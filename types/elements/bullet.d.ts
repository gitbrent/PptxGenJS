export declare type BulletOptions = boolean | {
    code?: string;
    type?: string;
    style?: string;
    startAt?: number;
    color?: string;
    indent?: string | number;
};
export default class Bullet {
    enabled: boolean;
    default: boolean;
    inherit: boolean;
    code?: string;
    type?: string;
    style?: string;
    startAt?: number;
    bulletCode?: string;
    color?: string;
    indent?: string | number;
    constructor(bullet?: BulletOptions);
    renderIndentProps(presLayout: any, indentLevel: any): string;
    render(): string;
}
