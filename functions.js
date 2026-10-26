// Two custom functions. Only the metadata differs between functions.json (WITHENUM uses
// the custom enum STYLE) and functions-noenum.json (no enums); this script is shared.

/**
 * @param {any} value
 * @returns {any}
 */
function plain(value) {
	return value;
}

/**
 * @param {number} style
 * @returns {string}
 */
function withEnum(style) {
	return style === 1 ? "Name" : "Caption";
}

CustomFunctions.associate("PLAIN", plain);
CustomFunctions.associate("WITHENUM", withEnum);
