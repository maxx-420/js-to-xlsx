

const escapeCharMap = {
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;'
  };

export const toArrayOfArray = (aoo = [], columnKeys = []) => {
  return aoo.map(function (obj) {
    return columnKeys.map(function (key) {
      return (obj[key] !== null || obj[key] !== undefined)? obj[key]: null;
    });
  });
};

export const createNode = ({
  nodeName,
  attr = {},
  cellContent = "",
  children = [],
}) => {
  let attributes = Object.keys(attr).reduce((str, key) => {
    return `${str? str + ' ': ''}${key}="${attr[key]}"`;
  }, "");
  let data =
    (children.length > 0 &&
      children.map((child) => createNode(child)).join("")) ||
    cellContent;
  let node = data.toString()? `<${nodeName}${
    attributes ? " " + attributes : ""
  }>${data}</${nodeName}>`: `<${nodeName}${
    attributes ? " " + attributes : ""
  }/>`
  return node;
};

export const addToZip = ( zip, obj ) => {
	Object.keys(obj).forEach(key=> {
		if (typeof obj[key] == 'object') {
			let newDir = zip.folder( key );
			addToZip( newDir, obj[key] );
		}
		else {
			zip.file( key, obj[key] );
		}
	} );
    return zip;
}

export const escape = (str) => {

    const escapeRegexLiteral = /[&<>"']/g;
    const escapeRegex = RegExp(escapeRegexLiteral.source)

    return (str && escapeRegex.test(str))
    ? str.replace(escapeRegexLiteral, (chr) => escapeCharMap[chr])
    : (str || '')
}