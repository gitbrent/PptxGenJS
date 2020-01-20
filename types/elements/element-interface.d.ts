import Placeholder from './placeholder';
export default interface Element {
    placeholder?: string;
    render: (index: number, presLayout: any, placeholder?: Placeholder) => string;
}
