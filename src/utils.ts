export function isArrEqual<T extends string> (arr1: T[], arr2: T[]): boolean {
	return arr1.length !== arr2.length ? false : arr1.every((element, index) => element === arr2[index])
}
